/**
 * Schüler-Klasse aus Entra-Klassengruppen (classTeams / classGroupMatchByKey).
 */
import { studentKlasseFromRecord, persistDemoKlasseCode, prefillStudentFreistellungForm } from './freistellung-planer-state.js';
import { loadClassTeamsContext } from './freistellung-planer-class-context.js';

export { loadClassTeamsContext } from './freistellung-planer-class-context.js';

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Graph-Gruppen-IDs aller verknüpften Klassenteams (für checkMemberGroups).
 * @param {object} setup
 * @param {object[]} classTeams
 * @returns {string[]}
 */
export function listClassGraphGroupIds(setup, classTeams) {
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
    return out;
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
    const hit = matchStudentKlasseFromMemberGroupIds(
        memberIdSet,
        classTeams,
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
