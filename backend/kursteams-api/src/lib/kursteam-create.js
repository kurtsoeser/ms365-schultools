'use strict';

const {
    sleep,
    graphRequest,
    graphJson,
    isGraphDuplicateRefError,
    parseTeamsOperationPath,
    pollTeamsAsyncOperation
} = require('./graph-client');
const { getAppOnlyToken } = require('./msal-app-only');

function sanitizeEducationClassCode(t) {
    const raw = (t && (t.gruppenmail || t.teamName)) || 'Klasse';
    const s = String(raw).replace(/[^a-zA-Z0-9]/g, '');
    return s.substring(0, 50) || 'Klasse';
}

async function createEducationClassGroup(token, t, log) {
    const body = {
        '@odata.type': '#microsoft.graph.educationClass',
        displayName: t.teamName,
        mailNickname: t.gruppenmail,
        description: 'Kursteam (WebUntis / MS365-Schulverwaltung)',
        classCode: sanitizeEducationClassCode(t),
        externalSource: 'manual'
    };
    log('Education: POST /education/classes …');
    const edu = await graphJson('POST', '/education/classes', token, body);
    log('Education: Klasse angelegt, warte auf Replikation …');
    await sleep(5000);
    return edu.id;
}

async function waitForGroupOwners(token, gid, log, minOwners = 1, maxAttempts = 30) {
    for (let i = 0; i < maxAttempts; i++) {
        try {
            const data = await graphJson(
                'GET',
                '/groups/' + gid + '/owners?$select=id',
                token,
                undefined
            );
            const count = (data.value && data.value.length) || 0;
            if (count >= minOwners) {
                log('Besitzer in Graph bestätigt (' + count + ').');
                return;
            }
            const wait = 3000;
            log('Besitzer noch nicht sichtbar (' + count + '), Warte ' + wait / 1000 + ' s …');
            await sleep(wait);
        } catch (e) {
            const msg = String(e && e.message ? e.message : e);
            const retryable =
                /ResourceNotFound/i.test(msg) ||
                /does not exist/i.test(msg) ||
                /404/.test(msg);
            if (retryable && i < maxAttempts - 1) {
                const wait = 3000 + (i % 5) * 2000;
                log('Gruppe noch nicht repliziert (owners), Warte ' + wait / 1000 + ' s …');
                await sleep(wait);
                continue;
            }
            throw e;
        }
    }
    throw new Error(
        'Timeout: Gruppe hat keinen Besitzer in Graph – Team-Erstellung abgebrochen. ' +
            'Prüfen Sie UPN/E-Mail des Besitzers und Berechtigung Group.ReadWrite.All.'
    );
}

async function addGroupOwnerAndMember(token, gid, ownerId, log) {
    const refBody = {
        '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + ownerId
    };

    async function postRef(path, label) {
        for (let attempt = 0; attempt < 6; attempt++) {
            try {
                await graphJson('POST', path, token, refBody);
                return;
            } catch (e) {
                const msg = String(e && e.message ? e.message : e);
                const retryable =
                    /ResourceNotFound/i.test(msg) ||
                    /does not exist/i.test(msg) ||
                    /404/.test(msg);
                if (retryable && attempt < 5) {
                    const wait = 2000 + attempt * 3000;
                    log(label + ': Gruppe noch nicht repliziert, Warte ' + wait / 1000 + ' s …');
                    await sleep(wait);
                    continue;
                }
                if (isGraphDuplicateRefError(e)) {
                    log(label + ': bereits gesetzt.');
                    return;
                }
                throw e;
            }
        }
    }

    await sleep(2000);
    await postRef('/groups/' + gid + '/owners/$ref', 'Besitzer');
    try {
        await postRef('/groups/' + gid + '/members/$ref', 'Mitglied');
    } catch (e) {
        log('Hinweis (Besitzer als Mitglied): ' + e.message);
    }
}

async function provisionKursteamPostTeamsEducationTemplate(token, gid, log, tenantId) {
    const postBody = {
        'template@odata.bind':
            'https://graph.microsoft.com/v1.0/teamsTemplates(\'educationClass\')',
        'group@odata.bind': 'https://graph.microsoft.com/v1.0/groups(\'' + gid + '\')'
    };

    let lastPostErr = null;
    for (let attempt = 0; attempt < 3; attempt++) {
        try {
            const res = await graphRequest('POST', '/teams', token, postBody);
            const text = await res.text();
            if (res.status === 202 || res.status === 200) {
                const loc = res.headers.get('Location') || res.headers.get('Content-Location');
                const opPath = parseTeamsOperationPath(loc);
                if (opPath) {
                    log('Teams: POST /teams mit Template educationClass …');
                    await pollTeamsAsyncOperation(token, opPath, log);
                } else {
                    log('Teams: POST /teams angenommen (keine Operation-URL).');
                }
                return await getAppOnlyToken(tenantId);
            }
            if (res.status === 404 && attempt < 2) {
                log('Teams: 404 nach Klassenanlage – Replikation, Warte 10 s …');
                await sleep(10000);
                token = await getAppOnlyToken(tenantId);
                continue;
            }
            if (
                (res.status === 400 || res.status === 403) &&
                /one or more owners/i.test(text) &&
                attempt < 4
            ) {
                log('Teams: Besitzer noch nicht bei Teams angekommen, Warte 15 s …');
                await sleep(15000);
                await waitForGroupOwners(token, gid, log);
                continue;
            }
            lastPostErr = new Error('POST /teams: ' + res.status + ' ' + (text || ''));
            break;
        } catch (e) {
            lastPostErr = e;
            if (attempt < 2 && /404/.test(String(e.message))) {
                log('Teams: Wiederholung nach Wartezeit (404) …');
                await sleep(10000);
                token = await getAppOnlyToken(tenantId);
                continue;
            }
            break;
        }
    }

    const detail = lastPostErr && lastPostErr.message ? lastPostErr.message : String(lastPostErr);
    throw new Error(
        'POST /teams (Template educationClass) ist fehlgeschlagen – kein Kursteam angelegt. Details: ' +
            detail
    );
}

/**
 * @param {{ teamName: string, gruppenmail: string, besitzer: string }} team
 * @param {(msg: string) => void} log
 * @param {string} tenantId
 * @returns {Promise<{ groupId: string, ownerId: string }>}
 */
async function createSingleKursteam(team, log, tenantId) {
    let token = await getAppOnlyToken(tenantId);

    const owner = await graphJson(
        'GET',
        '/users/' + encodeURIComponent(team.besitzer),
        token,
        undefined
    );
    const ownerId = owner.id;

    const gid = await createEducationClassGroup(token, team, log);
    await addGroupOwnerAndMember(token, gid, ownerId, log);
    await waitForGroupOwners(token, gid, log);
    await provisionKursteamPostTeamsEducationTemplate(token, gid, log, tenantId);

    return { groupId: gid, ownerId };
}

module.exports = {
    sanitizeEducationClassCode,
    createSingleKursteam
};
