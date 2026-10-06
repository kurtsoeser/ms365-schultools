/**
 * Schüler-Klasse aus Entra-Klassengruppen (classTeams / classGroupMatchByKey).
 */
import { getGraphToken, graphJson, fetchAllPages } from '../../shared/graph-client.js';
import {
    studentKlasseFromRecord,
    persistDemoKlasseCode,
    prefillStudentFreistellungForm,
    classesForStudentPicker,
    refreshStudentMatchFromStammdaten,
    isStudentKlasseLocked
} from './freistellung-planer-state.js';
import { loadClassTeamsContext, classTeamLinksFromAppData } from './freistellung-planer-class-context.js';
import { loadEffectivePermissionsConfig } from './freistellung-planer-permissions.js';

export { loadClassTeamsContext } from './freistellung-planer-class-context.js';

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Graph-Gruppen-IDs aller verknüpften Klassenteams (für checkMemberGroups).
 * @param {object} setup
 * @param {object[]} classTeams
 * @returns {string[]}
 */
export function listClassGraphGroupIds(setup, classTeams, classTeamLinks) {
    const seen = new Set();
    const out = [];
    const push = (id) => {
        const g = String(id || '').trim();
        if (!GUID_RE.test(g)) return;
        const key = g.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(g);
    };
    (classTeams || []).forEach((t) => push(t && t.graphGroupId));
    const map = setup && setup.classGroupMatchByKey && typeof setup.classGroupMatchByKey === 'object'
        ? setup.classGroupMatchByKey
        : {};
    Object.values(map).forEach((entry) => {
        if (!entry || typeof entry !== 'object') return;
        push(entry.groupId || entry.graphGroupId);
    });
    (classTeamLinks || []).forEach((row) => {
        if (!row || typeof row !== 'object') return;
        push(row.groupId || row.graphGroupId);
    });
    return out;
}

function classTeamsForKlasseMatch(classTeams, classTeamLinks) {
    const teams = (classTeams || []).slice();
    (classTeamLinks || []).forEach((row) => {
        if (!row || !row.groupId) return;
        teams.push({
            graphGroupId: row.groupId,
            classCode: row.code,
            displayName: row.name || row.code
        });
    });
    return teams;
}

/**
 * @param {Set<string>|string[]} memberIdSet – IDs aus checkMemberGroups (lowercase ok)
 * @param {object[]} classTeams
 * @param {Record<string, object>} classGroupMatchByKey
 * @returns {{ klasse: string }|null}
 */
export function matchStudentKlasseFromMemberGroupIds(memberIdSet, classTeams, classGroupMatchByKey) {
    const member = new Set();
    if (memberIdSet instanceof Set) {
        memberIdSet.forEach((id) => member.add(String(id).toLowerCase()));
    } else {
        (memberIdSet || []).forEach((id) => member.add(String(id).toLowerCase()));
    }
    if (!member.size) return null;

    for (const t of classTeams || []) {
        if (!t) continue;
        const gid = String(t.graphGroupId || '').trim().toLowerCase();
        if (!gid || !member.has(gid)) continue;
        const code = String(t.classCode || t.displayName || '').trim();
        if (code) return { klasse: code, graphGroupId: gid };
    }

    const map =
        classGroupMatchByKey && typeof classGroupMatchByKey === 'object' ? classGroupMatchByKey : {};
    for (const [key, entry] of Object.entries(map)) {
        const gid = String((entry && (entry.groupId || entry.graphGroupId)) || '')
            .trim()
            .toLowerCase();
        if (!gid || !member.has(gid)) continue;
        const code = String(key || '').trim();
        if (code) return { klasse: code, graphGroupId: gid };
    }
    return null;
}

/**
 * Stammdaten-Schülerzeile hat Vorrang; sonst Klasse aus Klassen-Entra-Gruppe.
 * @param {object} state
 * @param {Set<string>|string[]|undefined} memberIdSet
 */
export function applyStudentKlasseFromEntraMembership(state, memberIdSet) {
    if (!state || !memberIdSet) return;
    if (studentKlasseFromRecord(state.studentMatch)) return;

    const { classTeams, setup } = loadClassTeamsContext();
    const perms = loadEffectivePermissionsConfig();
    const teams = classTeamsForKlasseMatch(classTeams, perms.classTeamLinks);
    const hit = matchStudentKlasseFromMemberGroupIds(
        memberIdSet,
        teams,
        setup && setup.classGroupMatchByKey
    );
    if (!hit || !hit.klasse) return;

    state.studentMatch = {
        ...(state.studentMatch && typeof state.studentMatch === 'object' ? state.studentMatch : {}),
        email: state.accountEmail,
        name: state.accountName || (state.studentMatch && state.studentMatch.name) || '',
        klasse: hit.klasse,
        klasseSource: 'entra-class-group'
    };
    state.demoKlasseCode = '';
    persistDemoKlasseCode('');
    prefillStudentFreistellungForm(state);
}

const USER_READ_SCOPES = ['https://graph.microsoft.com/User.Read'];

const MEMBER_CHECK_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/GroupMember.Read.All',
    'https://graph.microsoft.com/Group.Read.All'
];

/**
 * @param {string[]} groupIds
 * @returns {Promise<Set<string>>}
 */
async function tryGraphTokenSilent(scopes) {
    if (typeof window !== 'undefined' && typeof window.ms365AuthAcquireTokenSilent === 'function') {
        try {
            return await window.ms365AuthAcquireTokenSilent(scopes);
        } catch {
            return null;
        }
    }
    try {
        return await getGraphToken(scopes);
    } catch {
        return null;
    }
}

/**
 * Klasse aus Gruppennamen (z. B. „Klasse 1B“, Jahrgangsgruppe).
 * @param {object[]} groups
 * @param {string[]} knownCodes
 */
export function inferStudentKlasseFromGroupLabels(groups, knownCodes) {
    const canonical = new Map();
    (knownCodes || []).forEach((c) => {
        const x = String(c || '').trim();
        if (!x) return;
        canonical.set(x.toLowerCase(), x);
    });
    const accept = (candidate) => {
        const hit = String(candidate || '').trim();
        if (!hit || !/^[0-9]{1,2}[A-Za-z]{1,4}$/.test(hit)) return '';
        const low = hit.toLowerCase();
        if (canonical.has(low)) return canonical.get(low);
        if (!canonical.size) return hit;
        return '';
    };
    const fromLabel = (raw) => {
        const s = String(raw || '').trim();
        if (!s) return '';
        const low = s.toLowerCase();
        for (const [codeLow, code] of canonical) {
            if (low === codeLow || low.includes('klasse ' + codeLow)) return code;
        }
        const m =
            s.match(/klasse\s+([0-9]{1,2}[A-Za-z]{1,4})/i) || s.match(/\b([0-9]{1,2}[A-Za-z]{1,4})\b/);
        return m ? accept(m[1]) : '';
    };
    for (const g of groups || []) {
        const fromDn = fromLabel(g && g.displayName);
        if (fromDn) return fromDn;
        const fromNick = fromLabel(g && g.mailNickname);
        if (fromNick) return fromNick;
    }
    return '';
}

async function fetchMemberGroupsForLabelInference() {
    const tok = await tryGraphTokenSilent(USER_READ_SCOPES);
    if (!tok) return [];
    try {
        const page = await fetchAllPages(
            tok,
            '/me/memberOf/microsoft.graph.group?$select=id,displayName,mailNickname',
            { maxItems: 256, maxPages: 6 }
        );
        return page.items || [];
    } catch {
        return [];
    }
}

async function memberGroupsWithToken(tok, ids, want) {
    if (!tok) return new Set();
    try {
        const data = await graphJson('POST', '/me/checkMemberGroups', tok, { groupIds: ids });
        const matched = (data && data.value) || [];
        const out = new Set(matched.map((g) => String(g).toLowerCase()));
        if (out.size) return out;
    } catch {
        /* memberOf */
    }
    try {
        const page = await fetchAllPages(tok, '/me/memberOf/microsoft.graph.group?$select=id', {
            maxItems: 512,
            maxPages: 8
        });
        const out = new Set();
        (page.items || []).forEach((g) => {
            const id = String((g && g.id) || '').toLowerCase();
            if (id && want.has(id)) out.add(id);
        });
        return out;
    } catch {
        return new Set();
    }
}

async function fetchUserMemberGroupIds(groupIds) {
    const ids = (groupIds || []).filter((id) => GUID_RE.test(String(id || '').trim()));
    if (!ids.length) return new Set();
    const want = new Set(ids.map((g) => String(g).toLowerCase()));
    let tok = await tryGraphTokenSilent(USER_READ_SCOPES);
    let out = await memberGroupsWithToken(tok, ids, want);
    if (out.size) return out;
    tok = await tryGraphTokenSilent(MEMBER_CHECK_SCOPES);
    out = await memberGroupsWithToken(tok, ids, want);
    return out;
}

/**
 * Nach Login / list-access: Klasse aus Klassen-Entra-Gruppe (z. B. 1A), wenn Stammdaten fehlen.
 * @param {object} state
 * @returns {Promise<boolean>}
 */
/**
 * Stammdaten → M365-Klassengruppe → gespeicherte Wahl.
 * @param {object} state
 */
export async function resolveStudentKlasseForAccount(state) {
    if (!state || state.role !== 'schueler') return false;
    refreshStudentMatchFromStammdaten(state);
    if (!isStudentKlasseLocked(state)) {
        await ensureStudentKlasseFromEntraForAccount(state);
    }
    prefillStudentFreistellungForm(state);
    return isStudentKlasseLocked(state) || !!studentKlasseFromRecord(state.studentMatch);
}

export async function ensureStudentKlasseFromEntraForAccount(state) {
    if (!state || state.role !== 'schueler') return false;
    if (studentKlasseFromRecord(state.studentMatch)) return true;

    const { classTeams, setup } = loadClassTeamsContext();
    const perms = loadEffectivePermissionsConfig();
    const linkRows = (perms.classTeamLinks || []).concat(classTeamLinksFromAppData());
    const classIds = listClassGraphGroupIds(setup, classTeams, linkRows);
    if (!classIds.length) return false;

    const member = await fetchUserMemberGroupIds(classIds);
    if (member.size) {
        applyStudentKlasseFromEntraMembership(state, member);
        if (studentKlasseFromRecord(state.studentMatch)) return true;
    }

    const picker = classesForStudentPicker(state);
    const known = picker.map((c) => String(c.code || c.name || '').trim()).filter(Boolean);
    const groups = await fetchMemberGroupsForLabelInference();
    const inferred = inferStudentKlasseFromGroupLabels(groups, known);
    if (!inferred) return false;

    state.studentMatch = {
        ...(state.studentMatch && typeof state.studentMatch === 'object' ? state.studentMatch : {}),
        email: state.accountEmail,
        name: state.accountName || '',
        klasse: inferred,
        klasseSource: 'entra-group-label'
    };
    state.demoKlasseCode = '';
    persistDemoKlasseCode('');
    prefillStudentFreistellungForm(state);
    return true;
}
