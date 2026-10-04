/**
 * Planer-Rolle aus Microsoft Entra-Gruppen, Stammdaten und Setup (Direktion-E-Mail).
 */
import { getGraphToken, graphJson, fetchAllPages } from '../../shared/graph-client.js';
import { userIsEntraGlobalAdministrator } from '../schularbeiten-planer/schularbeiten-planer-entra-role.js';
import {
    loadPermissionsConfig,
    normalizePermissionsConfig,
    entraGroupsConfigured
} from './freistellung-planer-permissions.js';
import { ROLE_STORAGE_KEY, resolveRole } from './freistellung-planer-state.js';
import { accountIsDirektionPlannerUser, accountIsPlannerUserInList } from './freistellung-planer-direktion-users.js';
import { applyStudentKlasseFromEntraMembership, listClassGraphGroupIds } from './freistellung-planer-student-klasse.js';
import { loadClassTeamsContext } from './freistellung-planer-class-context.js';

/** @typedef {'direktion'|'kv'|'schueler'} FrPlanerRole */

export const PLANER_ROLE_ORDER = ['direktion', 'kv', 'schueler'];

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const MEMBER_CHECK_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/GroupMember.Read.All',
    'https://graph.microsoft.com/Group.Read.All'
];

/**
 * @param {string[]} groupIds
 * @returns {Promise<Set<string>>}
 */
async function fetchUserMemberGroupIds(groupIds) {
    const ids = (groupIds || []).filter((id) => GUID_RE.test(String(id || '').trim()));
    if (!ids.length) return new Set();
    const want = new Set(ids.map((id) => String(id).toLowerCase()));
    try {
        const tok = await getGraphToken(MEMBER_CHECK_SCOPES);
        const data = await graphJson('POST', '/me/checkMemberGroups', tok, { groupIds: ids });
        const matched = (data && data.value) || [];
        const out = new Set(matched.map((g) => String(g).toLowerCase()));
        if (out.size) return out;
    } catch {
        /* fallback memberOf */
    }
    try {
        const tok = await getGraphToken(MEMBER_CHECK_SCOPES);
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

/**
 * @param {FrPlanerRole[]} roles
 * @returns {FrPlanerRole[]}
 */
export function sortPlanerRoles(roles) {
    const set = new Set((roles || []).map((r) => String(r)));
    return PLANER_ROLE_ORDER.filter((r) => set.has(r));
}

/**
 * @param {FrPlanerRole[]} available
 * @param {{ preferredActiveRole?: FrPlanerRole|null, preferredDemoRole?: FrPlanerRole|null }} [opts]
 * @returns {FrPlanerRole}
 */
export function resolveActivePlanerRole(available, opts) {
    const list = sortPlanerRoles(available || []);
    const set = new Set(list);
    const cur = opts && opts.preferredActiveRole;
    const demo = opts && opts.preferredDemoRole;
    if (cur && set.has(cur)) return cur;
    if (demo && set.has(demo)) return demo;
    const stored = resolveRole();
    if (stored && set.has(stored)) return stored;
    try {
        const raw = String(localStorage.getItem(ROLE_STORAGE_KEY) || '').toLowerCase();
        if (raw === 'direktion' || raw === 'kv' || raw === 'schueler') {
            if (set.has(raw)) return raw;
        }
    } catch {
        /* ignore */
    }
    return list[0] || 'schueler';
}

/**
 * @param {ReturnType<typeof normalizePermissionsConfig>} config
 * @returns {string[]}
 */
export function planerEntraGroupIds(config) {
    const c = normalizePermissionsConfig(config || loadPermissionsConfig());
    const ids = [c.groupDirektionId, c.groupKvId, c.groupSchuelerId]
        .map((id) => String(id || '').trim())
        .filter((id) => GUID_RE.test(id));
    return [...new Set(ids)];
}

/**
 * @param {Set<string>|string[]} memberIds
 * @param {ReturnType<typeof normalizePermissionsConfig>} [config]
 * @returns {FrPlanerRole[]}
 */
export function listRolesFromEntraGroups(memberIds, config) {
    const c = normalizePermissionsConfig(config || loadPermissionsConfig());
    const member = memberIds instanceof Set ? memberIds : new Set((memberIds || []).map((x) => String(x).toLowerCase()));
    const inGroup = (idKey) => {
        const id = String(c[idKey] || '').trim().toLowerCase();
        return id && member.has(id);
    };
    const roles = [];
    if (inGroup('groupDirektionId')) roles.push('direktion');
    if (inGroup('groupKvId')) roles.push('kv');
    if (inGroup('groupSchuelerId')) roles.push('schueler');
    return roles;
}

/**
 * @param {{ studentMatch?: object|null, kvMatch?: object|null, direktionMatch?: boolean, accountEmail?: string }} scope
 * @returns {FrPlanerRole[]}
 */
export function listRolesFromStammdaten(scope) {
    const roles = [];
    if (scope && scope.direktionMatch) roles.push('direktion');
    if (scope && scope.kvMatch) roles.push('kv');
    if (scope && scope.studentMatch) roles.push('schueler');
    return roles;
}

/**
 * @param {FrPlanerRole|string} role
 * @param {string} [source]
 */
export function roleSourceLabel(source) {
    const s = String(source || '');
    if (s === 'entra') return 'Entra-Gruppe';
    if (s === 'global-admin') return 'Globaler Administrator';
    if (s === 'direktion-schueler') return 'Schüler-Ansicht (Direktion)';
    if (s === 'stammdaten') return 'Stammdaten';
    if (s === 'setup-direktion') return 'Setup (Direktion-E-Mail)';
    if (s === 'direktion-user') return 'Verwaltung (Einzelperson)';
    if (s === 'kv-user') return 'Klassenvorstand (Einzelperson)';
    if (s === 'schueler-user') return 'Schüler/in (Einzelperson)';
    if (s === 'list-access') return 'Freistellungsliste (SharePoint)';
    if (s === 'demo') return 'Demo';
    return '';
}

/**
 * @param {FrPlanerRole[]} roles
 * @param {Record<string, string>} sources
 */
export function finalizePlanerRoles(roles, sources) {
    const list = sortPlanerRoles(roles || []);
    const src = { ...(sources || {}) };
    if (list.includes('direktion') && !list.includes('schueler')) {
        list.push('schueler');
        if (!src.schueler) src.schueler = 'direktion-schueler';
    }
    return { roles: sortPlanerRoles(list), sources: src };
}

const ROLE_HINT_PUBLIC_DENIED =
    'Ihr Konto ist für den Freistellungs-Planer noch nicht freigeschaltet. Bitte wenden Sie sich an Klassenvorstand oder Sekretariat.';
const ROLE_HINT_STAFF_NO_GROUPS =
    'Keine Planer-Berechtigung: IT muss im Freistellungen-Setup Entra-Gruppen wählen und „Alles speichern“ (veröffentlicht die Gruppen auf SharePoint und in der Listen-Beschreibung). Danach Planer neu laden. IT-Vorschau: ?demoRole=1';
const ROLE_HINT_STAFF_NOT_IN_GROUP =
    'Keine Berechtigung: Konto ist in keiner konfigurierten Entra-Gruppe (Schüler, Klassenvorstand oder Direktion) und nicht in den Stammdaten zugeordnet.';

function setPlanerRoleHints(state, publicHint, staffHint) {
    state.roleHintPublic = String(publicHint || '');
    state.roleHintStaff = String(staffHint || '');
    state.roleHint = state.roleHintStaff || state.roleHintPublic;
}

function applyRolesToState(state, roles, sources, opts) {
    const finalized = finalizePlanerRoles(roles, sources);
    state.planerRoles = finalized.roles;
    state.planerRoleSources = finalized.sources;
    if (finalized.roles.length === 0) {
        state.role = '';
        state.planerAccessDenied = true;
        return;
    }
    state.planerAccessDenied = false;
    state.role = resolveActivePlanerRole(finalized.roles, {
        preferredActiveRole: opts && opts.preferredActiveRole,
        preferredDemoRole: opts && opts.preferredDemoRole
    });
    state.roleSource = state.planerRoleSources[state.role] || state.roleSource || 'entra';
}

/**
 * Rollen-Umschalter nur für IT-Vorschau (?demoRole=1), nicht wenn Entra fehlt.
 */
export function isPlanerDemoRoleUiEnabled(_entraConfigured, demoRoleOverride) {
    return !!demoRoleOverride;
}

export function readDemoRoleOverrideFromUrl(demoRoleOverride) {
    if (demoRoleOverride) return true;
    try {
        const p = new URLSearchParams(typeof window !== 'undefined' ? window.location.search || '' : '');
        return p.get('demoRole') === '1';
    } catch {
        return false;
    }
}

function isLoggedIn() {
    try {
        return typeof window !== 'undefined' &&
            typeof window.ms365AuthIsLoggedIn === 'function' &&
            window.ms365AuthIsLoggedIn();
    } catch {
        return false;
    }
}

/**
 * @param {object} state
 * @param {{ demoRoleOverride?: boolean, preferredDemoRole?: FrPlanerRole|null, preferredActiveRole?: FrPlanerRole|null }} [opts]
 */
export async function applyPlanerRoleFromEntra(state, opts) {
    const config = loadPermissionsConfig();
    const entra = entraGroupsConfigured(config);
    state.entraGroupsConfigured = entra;
    const demoOverride = readDemoRoleOverrideFromUrl(opts && opts.demoRoleOverride);
    state.demoRoleOverride = demoOverride;

    const pickOpts = {
        preferredActiveRole: (opts && opts.preferredActiveRole) || state.role,
        preferredDemoRole: opts && opts.preferredDemoRole
    };

    if (demoOverride) {
        applyRolesToState(
            state,
            ['direktion', 'kv', 'schueler'],
            { direktion: 'demo', kv: 'demo', schueler: 'demo' },
            pickOpts
        );
        state.roleSource = 'demo';
        state.planerAccessDenied = false;
        setPlanerRoleHints(state, '', '');
        return;
    }

    if (!isLoggedIn()) {
        state.planerRoles = [];
        state.planerRoleSources = {};
        state.role = '';
        if (entra) {
            state.roleSource = 'entra';
            state.planerAccessDenied = false;
            setPlanerRoleHints(
                state,
                'Bitte mit Ihrem Schul-Microsoft-Konto anmelden.',
                'Nicht angemeldet – Rolle nach Anmeldung über Entra-Gruppen aus dem Freistellungen-Setup.'
            );
        } else {
            state.planerAccessDenied = true;
            setPlanerRoleHints(state, ROLE_HINT_PUBLIC_DENIED, ROLE_HINT_STAFF_NO_GROUPS);
        }
        return;
    }

    const scope = {
        studentMatch: state.studentMatch,
        kvMatch: state.kvMatch,
        direktionMatch: state.direktionMatch,
        accountEmail: state.accountEmail
    };

    if (!entra) {
        const collected = await collectPlanerRolesForUser(scope, config, false, true);
        if (collected.roles.length) {
            applyRolesToState(state, collected.roles, collected.sources, pickOpts);
            applyStudentKlasseFromEntraMembership(state, collected.memberIdSet);
            state.planerAccessDenied = false;
            setPlanerRoleHints(
                state,
                '',
                'Entra-Gruppen im Freistellungen-Setup fehlen – Rolle vorläufig aus Stammdaten oder Direktion-E-Mail. Bitte Gruppen eintragen und „Alles speichern“.'
            );
            return;
        }
        state.role = '';
        state.planerRoles = [];
        state.planerRoleSources = {};
        state.planerAccessDenied = true;
        setPlanerRoleHints(state, ROLE_HINT_PUBLIC_DENIED, ROLE_HINT_STAFF_NO_GROUPS);
        return;
    }

    try {
        const collected = await collectPlanerRolesForUser(scope, config, entra, false);

        if (collected.roles.length) {
            applyRolesToState(state, collected.roles, collected.sources, pickOpts);
            applyStudentKlasseFromEntraMembership(state, collected.memberIdSet);
            state.planerAccessDenied = false;
            const staffOnly =
                collected.roles.length > 1
                    ? ''
                    : collected.sources[state.role] === 'stammdaten'
                      ? 'Rolle aus Stammdaten (E-Mail-Zuordnung).'
                      : collected.sources[state.role] === 'setup-direktion'
                        ? 'Rolle als Direktion (E-Mail im Setup).'
                        : '';
            setPlanerRoleHints(state, '', staffOnly);
            return;
        }

        state.role = '';
        state.roleSource = 'entra';
        state.planerRoles = [];
        state.planerRoleSources = {};
        state.planerAccessDenied = true;
        setPlanerRoleHints(state, ROLE_HINT_PUBLIC_DENIED, ROLE_HINT_STAFF_NOT_IN_GROUP);
    } catch (e) {
        const collected = await collectPlanerRolesForUser(scope, config, entra, true);
        if (collected.roles.length) {
            applyRolesToState(state, collected.roles, collected.sources, pickOpts);
            applyStudentKlasseFromEntraMembership(state, collected.memberIdSet);
            state.planerAccessDenied = false;
            setPlanerRoleHints(
                state,
                '',
                'Entra-Gruppen konnten nicht geprüft werden – vorläufig aus anderen Quellen. ' +
                    (e && e.message ? e.message : String(e))
            );
            return;
        }
        state.planerAccessDenied = true;
        state.role = '';
        setPlanerRoleHints(
            state,
            ROLE_HINT_PUBLIC_DENIED,
            'Entra-Gruppen konnten nicht geprüft werden. ' + (e && e.message ? e.message : String(e))
        );
    }
}

/**
 * @param {{ studentMatch?: object|null, kvMatch?: object|null, direktionMatch?: boolean, accountEmail?: string }} scope
 * @param {ReturnType<typeof normalizePermissionsConfig>} config
 * @param {boolean} entra
 * @param {boolean} [skipEntraGroups]
 */
async function collectPlanerRolesForUser(scope, config, entra, skipEntraGroups) {
    /** @type {FrPlanerRole[]} */
    const roles = [];
    /** @type {Record<string, string>} */
    const sources = {};
    /** @type {Set<string>} */
    let memberIdSet = new Set();

    const add = (role, source) => {
        if (!roles.includes(role)) roles.push(role);
        if (!sources[role]) sources[role] = source;
    };

    if (await userIsEntraGlobalAdministrator()) {
        add('direktion', 'global-admin');
        add('kv', 'global-admin');
    }

    if (entra && !skipEntraGroups) {
        const planerIds = planerEntraGroupIds(config);
        const planerMember = await fetchUserMemberGroupIds(planerIds);
        listRolesFromEntraGroups(planerMember, config).forEach((r) => add(r, 'entra'));
        const { classTeams, setup } = loadClassTeamsContext();
        const classIds = listClassGraphGroupIds(setup, classTeams);
        const classMember = classIds.length ? await fetchUserMemberGroupIds(classIds) : new Set();
        memberIdSet = new Set([...planerMember, ...classMember]);
    }

    if (scope && scope.direktionMatch) add('direktion', 'setup-direktion');
    if (accountIsDirektionPlannerUser(scope && scope.accountEmail, config.direktionUsers)) {
        add('direktion', 'direktion-user');
    }
    if (accountIsPlannerUserInList(scope && scope.accountEmail, config.kvUsers)) {
        add('kv', 'kv-user');
    }
    if (accountIsPlannerUserInList(scope && scope.accountEmail, config.schuelerUsers)) {
        add('schueler', 'schueler-user');
    }
    listRolesFromStammdaten(scope).forEach((r) => {
        if (r === 'direktion' && sources.direktion) return;
        add(r, 'stammdaten');
    });

    const finalized = finalizePlanerRoles(roles, sources);
    finalized.memberIdSet = memberIdSet;
    return finalized;
}

export { entraGroupsConfigured };
