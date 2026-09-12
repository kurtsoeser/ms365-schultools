/**
 * Echtes Kursteam (Aufgaben, Klassennotizbuch): educationClass / EDU_Class.
 *
 * Reihenfolge (Microsoft Learn / Kursteams-Tool):
 * 1) möglichst POST /education/classes (Education-Metadaten + Assignments-fähig)
 * 2) sonst M365-Gruppe mit creationOptions classAssignments (Microsoft Support)
 * 3) POST /teams mit teamsTemplates('educationClass') – nie PUT /groups/{id}/team
 */

/** Tenant-übergreifende Education-Directory-Erweiterung (SDS / Class Teams). */
export const EDUCATION_OBJECT_TYPE_EXTENSION =
    'extension_fe2174665583431c953114ff7268b7b3_Education_ObjectType';

/** SharePoint-App-Rollen für EDU-Apps (Assignments, Forms, OneNote, …). */
export const CLASS_TEAM_RESOURCE_BEHAVIOR_OPTIONS = [
    'appRoleForSite:22d27567-b3f0-4dc2-9ec2-46ed368ba538:fullcontrol',
    'appRoleForSite:c9a559d2-7aab-4f13-a6ed-e7e9c52aec87:fullcontrol',
    'appRoleForSite:13291f5a-59ac-4c59-b0fa-d1632e8f3292:fullcontrol',
    'appRoleForSite:2d4d3d8e-2be3-4bef-9f87-7875a61c29de:fullcontrol',
    'appRoleForSite:8f348934-64be-4bb2-bc16-c54c96789f43:fullcontrol'
];

function norm(s) {
    return String(s == null ? '' : s).trim();
}

export function sanitizeEducationClassCode(raw) {
    const s = String(raw || 'Klasse').replace(/[^a-zA-Z0-9]/g, '');
    return s.substring(0, 50) || 'Klasse';
}

export function parseTeamsOperationPath(locationHeader) {
    if (!locationHeader) return null;
    const loc = String(locationHeader).trim();
    const m = loc.match(/teams\('([^']+)'\)\/operations\('([^']+)'\)/i);
    if (m) return '/teams/' + m[1] + '/operations/' + m[2];
    const m2 = loc.match(/\/teams\/([^/]+)\/operations\/([^/?\s]+)/i);
    if (m2) return '/teams/' + m2[1] + '/operations/' + m2[2];
    return null;
}

function isDup(e) {
    const msg = String((e && e.message) || e || '');
    return /already exist|added object references|Conflict|One or more added/i.test(msg);
}

/**
 * @param {{
 *   graphJson: Function,
 *   graphRequest: Function,
 *   sleep?: Function,
 *   getToken?: Function,
 *   token?: string,
 *   log?: Function,
 *   displayName: string,
 *   mailNickname: string,
 *   description?: string,
 *   classCode?: string,
 *   ownerId: string
 * }} opts
 * @returns {Promise<{ groupId: string, method: 'educationClass'|'classAssignmentsGroup', created: boolean }>}
 */
export async function createEducationClassTeam(opts) {
    const o = opts || {};
    const graphJson = o.graphJson;
    const graphRequest = o.graphRequest;
    const sleep =
        o.sleep ||
        function (ms) {
            return new Promise(function (r) {
                setTimeout(r, ms);
            });
        };
    const log =
        o.log ||
        function () {
            /* noop */
        };
    const getToken = o.getToken;
    let token = o.token || '';

    async function refreshToken() {
        if (typeof getToken === 'function') {
            token = await getToken();
            return token;
        }
        if (token) return token;
        throw new Error('Kein Graph-Token.');
    }

    const displayName = norm(o.displayName);
    const mailNickname = norm(o.mailNickname);
    const description = norm(o.description) || 'Kursteam (MS365-Schulverwaltung)';
    const classCode = sanitizeEducationClassCode(o.classCode || mailNickname || displayName);
    const ownerId = norm(o.ownerId);

    if (!graphJson || !graphRequest) throw new Error('Graph-Helfer fehlen.');
    if (!displayName) throw new Error('Anzeigename fehlt.');
    if (!mailNickname) throw new Error('Mail-Nickname fehlt.');
    if (!ownerId) throw new Error('Besitzer-ID fehlt (Kursteam braucht mind. einen Owner).');

    await refreshToken();

    let groupId = '';
    let method = 'educationClass';

    try {
        log('Education: POST /education/classes …');
        const edu = await graphJson('POST', '/education/classes', token, {
            '@odata.type': '#microsoft.graph.educationClass',
            displayName: displayName,
            mailNickname: mailNickname,
            description: description,
            classCode: classCode,
            externalSource: 'manual'
        });
        groupId = edu && edu.id ? String(edu.id) : '';
        if (!groupId) throw new Error('education/classes ohne ID.');
        log('Education-Klasse angelegt.');
    } catch (e1) {
        log('Education/classes nicht möglich: ' + (e1 && e1.message ? e1.message : e1));
        log('Fallback: Gruppe mit classAssignments (Microsoft Support) …');
        method = 'classAssignmentsGroup';
        await refreshToken();
        const body = {
            displayName: displayName,
            description: description,
            groupTypes: ['Unified'],
            mailEnabled: true,
            securityEnabled: false,
            mailNickname: mailNickname,
            visibility: 'HiddenMembership',
            'owners@odata.bind': ['https://graph.microsoft.com/v1.0/users/' + ownerId],
            'members@odata.bind': ['https://graph.microsoft.com/v1.0/users/' + ownerId],
            creationOptions: ['ExchangeProvisioningFlags:461', 'classAssignments']
        };
        body[EDUCATION_OBJECT_TYPE_EXTENSION] = 'Section';
        body.resourceBehaviorOptions = CLASS_TEAM_RESOURCE_BEHAVIOR_OPTIONS.slice();
        const g = await graphJson('POST', '/groups', token, body);
        groupId = g && g.id ? String(g.id) : '';
        if (!groupId) throw new Error('Gruppenanlage (classAssignments) ohne ID.');
        log('Assignments-fähige Gruppe angelegt.');
    }

    await sleep(2000);
    await refreshToken();
    if (method === 'educationClass') {
        await addOwnerAndMember(graphJson, token, groupId, ownerId, log, sleep);
    }

    await refreshToken();
    await waitForGroupOwners(graphJson, token, groupId, log, sleep);
    await provisionEducationClassTemplate({
        graphJson: graphJson,
        graphRequest: graphRequest,
        sleep: sleep,
        getToken: getToken,
        log: log,
        groupId: groupId,
        token: token
    });

    return { groupId: groupId, method: method, created: true };
}

async function addOwnerAndMember(graphJson, token, gid, ownerId, log, sleep) {
    await sleep(1500);
    const ref = {
        '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + ownerId
    };
    try {
        await graphJson('POST', '/groups/' + encodeURIComponent(gid) + '/owners/$ref', token, ref);
    } catch (e) {
        if (!isDup(e)) throw e;
        log('Besitzer war schon gesetzt.');
    }
    try {
        await graphJson('POST', '/groups/' + encodeURIComponent(gid) + '/members/$ref', token, ref);
    } catch (e) {
        if (!isDup(e)) log('Hinweis Mitglied: ' + (e.message || e));
    }
}

async function waitForGroupOwners(graphJson, token, gid, log, sleep) {
    for (let i = 0; i < 20; i++) {
        try {
            const data = await graphJson(
                'GET',
                '/groups/' + encodeURIComponent(gid) + '/owners?$select=id',
                token
            );
            const count = (data && data.value && data.value.length) || 0;
            if (count >= 1) {
                log('Besitzer in Graph sichtbar (' + count + ').');
                return;
            }
        } catch (e) {
            if (!/404|ResourceNotFound/i.test(String(e && e.message))) throw e;
        }
        log('Warte auf Besitzer-Replikation …');
        await sleep(3000);
    }
    throw new Error('Timeout: Gruppe hat keinen sichtbaren Besitzer – Team-Anlage abgebrochen.');
}

/**
 * POST /teams mit Template educationClass.
 */
export async function provisionEducationClassTemplate(opts) {
    const o = opts || {};
    const graphJson = o.graphJson;
    const graphRequest = o.graphRequest;
    const sleep =
        o.sleep ||
        function (ms) {
            return new Promise(function (r) {
                setTimeout(r, ms);
            });
        };
    const log = o.log || function () {};
    const getToken = o.getToken;
    let token = o.token;
    const gid = norm(o.groupId);
    if (!gid) throw new Error('groupId fehlt.');
    if (!graphJson || !graphRequest) throw new Error('Graph-Helfer fehlen.');

    const postBody = {
        'template@odata.bind': "https://graph.microsoft.com/v1.0/teamsTemplates('educationClass')",
        'group@odata.bind': "https://graph.microsoft.com/v1.0/groups('" + gid + "')"
    };

    let lastErr = null;
    for (let attempt = 0; attempt < 4; attempt++) {
        try {
            if (!token && typeof getToken === 'function') token = await getToken();
            const res = await graphRequest('POST', '/teams', token, postBody);
            const text = await res.text();
            if (res.status === 202 || res.status === 200) {
                const loc = res.headers.get('Location') || res.headers.get('Content-Location');
                const opPath = parseTeamsOperationPath(loc);
                if (opPath) {
                    log('Teams: Template educationClass (EDU_Class) – warte auf Bereitstellung …');
                    await pollTeamsOp(graphJson, token, opPath, log, sleep);
                } else {
                    log('Teams: POST /teams angenommen (keine Operation-URL).');
                }
                return;
            }
            if (res.status === 404 && attempt < 3) {
                log('Teams 404 – Replikation, warte 10 s …');
                await sleep(10000);
                if (getToken) token = await getToken();
                continue;
            }
            if (
                (res.status === 400 || res.status === 403) &&
                /one or more owners/i.test(text) &&
                attempt < 3
            ) {
                log('Teams: Besitzer noch nicht angekommen, warte 15 s …');
                await sleep(15000);
                if (getToken) token = await getToken();
                continue;
            }
            lastErr = new Error('POST /teams: ' + res.status + ' ' + (text || '').slice(0, 400));
            break;
        } catch (e) {
            lastErr = e;
            if (attempt < 3 && /404/.test(String(e && e.message))) {
                await sleep(10000);
                if (getToken) token = await getToken();
                continue;
            }
            break;
        }
    }
    throw new Error(
        'POST /teams (educationClass) fehlgeschlagen – kein Kursteam. ' +
            'PUT /groups/…/team wird bewusst nicht verwendet. ' +
            (lastErr && lastErr.message ? lastErr.message : lastErr)
    );
}

async function pollTeamsOp(graphJson, token, operationPath, log, sleep) {
    for (let i = 0; i < 90; i++) {
        await sleep(2000);
        const data = await graphJson('GET', operationPath, token);
        const st = String((data && (data.status || data.Status)) || '').toLowerCase();
        if (st === 'succeeded') {
            log('Teams: educationClass bereit (Aufgaben/Notizbuch-Template).');
            return;
        }
        if (st === 'failed') {
            const errMsg =
                (data.error && (data.error.message || JSON.stringify(data.error))) ||
                JSON.stringify(data);
            throw new Error('Team-Bereitstellung fehlgeschlagen: ' + errMsg);
        }
        if (i > 0 && i % 10 === 0) log('Teams: warte … (' + i * 2 + ' s)');
    }
    throw new Error('Timeout: Team-Bereitstellung (educationClass).');
}

export default {
    EDUCATION_OBJECT_TYPE_EXTENSION,
    CLASS_TEAM_RESOURCE_BEHAVIOR_OPTIONS,
    sanitizeEducationClassCode,
    parseTeamsOperationPath,
    createEducationClassTeam,
    provisionEducationClassTemplate
};
