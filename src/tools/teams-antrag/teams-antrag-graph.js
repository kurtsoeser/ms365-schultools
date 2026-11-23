/**
 * Graph: Team aus genehmigtem Antrag anlegen.
 * Transport: shared/graph-client.js (R22 / b11).
 */
import { graphJson as sharedGraphJson } from '../../shared/graph-client.js';

import { buildProvisionPlan } from './teams-antrag-logic.js';
import { STATUS } from './teams-antrag-model.js';
import { getAntrag, setAntragStatus } from './teams-antrag-store.js';

function graphApi() {
    return typeof window !== 'undefined' ? window.ms365GraphUnifiedGroups : null;
}

/**
 * @param {string[]} scopes
 */
async function getToken(scopes) {
    const G = graphApi();
    if (!G || typeof G.getGraphToken !== 'function') {
        throw new Error('Graph-Modul nicht geladen.');
    }
    return G.getGraphToken(scopes);
}

/**
 * @param {string} token
 * @param {string} email
 */
async function resolveUserId(token, email) {
    const G = graphApi();
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return '';
    if (typeof G.resolveUserIdByEmail === 'function') {
        return (await G.resolveUserIdByEmail(token, em)) || '';
    }
    try {
        const data = await sharedGraphJson(
            'GET',
            "/users?$filter=mail eq '" +
                em.replace(/'/g, "''") +
                "' or userPrincipalName eq '" +
                em.replace(/'/g, "''") +
                "'&$select=id&$top=1",
            token,
            undefined
        );
        const u = data && data.value && data.value[0];
        return u && u.id ? String(u.id) : '';
    } catch {
        return '';
    }
}

/**
 * @param {string} token
 * @param {string} groupId
 * @param {string} userId
 * @param {'owner'|'member'} role
 */
async function addMember(token, groupId, userId, role) {
    const G = graphApi();
    if (!userId || !groupId) return;
    const path =
        role === 'owner'
            ? '/groups/' + encodeURIComponent(groupId) + '/owners/$ref'
            : '/groups/' + encodeURIComponent(groupId) + '/members/$ref';
    try {
        await sharedGraphJson('POST', path, token, {
            '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + userId
        });
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e || '');
        if (/already exist|AlreadyExists|added object references/i.test(msg)) return;
        throw e;
    }
}

/**
 * Genehmigen und Team anlegen.
 * @param {string} antragId
 * @param {{ onLog?: (msg: string) => void }} [opts]
 */
export async function approveAndCreateTeam(antragId, opts) {
    const log = (msg) => {
        if (opts && typeof opts.onLog === 'function') opts.onLog(msg);
    };
    const antrag = getAntrag(antragId);
    if (!antrag) throw new Error('Antrag nicht gefunden.');
    const plan = buildProvisionPlan(antrag);
    if (!plan.ok) throw new Error(plan.issues.join(', '));

    setAntragStatus(antragId, STATUS.GENEHMIGT, {
        finalDisplayName: plan.displayName,
        finalMailNickname: plan.mailNickname,
        lastError: ''
    });

    const G = graphApi();
    if (!G) throw new Error('Graph-Modul nicht geladen.');

    const scopes = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.ReadWrite.All',
        'https://graph.microsoft.com/Team.Create',
        'https://graph.microsoft.com/Directory.Read.All'
    ];

    try {
        log('Token …');
        const token = await getToken(scopes);
        log('Gruppe anlegen: ' + plan.displayName);
        const group = await G.createUnifiedGroup(
            token,
            plan.displayName,
            plan.mailNickname,
            plan.description
        );
        const gid = group && group.id ? String(group.id) : '';
        if (!gid) throw new Error('Gruppe ohne ID.');
        log('Team provisionieren …');
        await G.provisionTeamForGroup(token, gid);

        for (const em of plan.ownerEmails) {
            const uid = await resolveUserId(token, em);
            if (uid) {
                log('Besitzer: ' + em);
                await addMember(token, gid, uid, 'owner');
            }
        }
        for (const em of plan.memberEmails) {
            const uid = await resolveUserId(token, em);
            if (uid) {
                log('Mitglied: ' + em);
                await addMember(token, gid, uid, 'member');
            }
        }

        let webUrl = '';
        try {
            const team = await sharedGraphJson(
                'GET',
                '/teams/' + encodeURIComponent(gid) + '?$select=id,webUrl',
                token,
                undefined
            );
            webUrl = team && team.webUrl ? String(team.webUrl) : '';
        } catch {
            webUrl = '';
        }

        return setAntragStatus(antragId, STATUS.ANGELEGT, {
            teamId: gid,
            teamWebUrl: webUrl,
            finalDisplayName: plan.displayName,
            finalMailNickname: plan.mailNickname,
            lastError: ''
        });
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e || 'Unbekannter Fehler');
        setAntragStatus(antragId, STATUS.ANLAGE_FEHLGESCHLAGEN, { lastError: msg });
        throw e;
    }
}
