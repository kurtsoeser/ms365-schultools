/**
 * Planer-Rolle aus Microsoft-Entra-Gruppen (Administration → SharePoint-Berechtigungen).
 */
import { getGraphToken, graphJson, fetchAllPages } from '../../shared/graph-client.js';
import {
    loadPermissionsConfig,
    normalizePermissionsConfig
} from './schularbeiten-planer-permissions.js';
import { accountIsPlannerUserInList } from '../freistellung-planer/freistellung-planer-direktion-users.js';
import { resolveRole, ROLE_STORAGE_KEY } from './schularbeiten-planer-state.js';

const CHECK_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/GroupMember.Read.All',
    'https://graph.microsoft.com/Group.Read.All'
];

const GLOBAL_ADMIN_ROLE_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/RoleManagement.Read.Directory'
];

/** Entra: Globaler Administrator */
export const GLOBAL_ADMINISTRATOR_ROLE_TEMPLATE_ID = '62e90394-69f5-4237-9190-012177145e10';

/** @type {{ key: string, at: number, value: boolean } | null} */
let globalAdminCache = null;
const GLOBAL_ADMIN_CACHE_MS = 5 * 60 * 1000;

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** @typedef {'admin'|'lehrer'|'schueler'} PlanerRole */

export const PLANER_ROLE_ORDER = ['admin', 'lehrer', 'schueler'];

/**
 * @param {PlanerRole[]} roles
 * @returns {PlanerRole[]}
 */
export function sortPlanerRoles(roles) {
    const set = new Set((roles || []).map((r) => String(r)));
    return PLANER_ROLE_ORDER.filter((r) => set.has(r));
}

/**
 * @param {PlanerRole[]} available
 * @param {{ preferredActiveRole?: PlanerRole|null, preferredDemoRole?: PlanerRole|null }} [opts]
 * @returns {PlanerRole}
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
        if (raw === 'admin' || raw === 'lehrer' || raw === 'schueler') {
            if (set.has(raw)) return raw;
        }
    } catch {
        /* ignore */
    }
    return list[0] || 'lehrer';
}

/**
 * @param {ReturnType<typeof normalizePermissionsConfig>|null|undefined} [config]
 */
export function entraGroupsConfigured(config) {
    const c = normalizePermissionsConfig(config || loadPermissionsConfig());
    return !!(
        c.groupAdminId ||
        c.groupLehrerId ||
        c.groupSchuelerId ||
        c.adminUsers.length ||
        c.lehrerUsers.length ||
        c.schuelerUsers.length
    );
}

/**
 * @param {ReturnType<typeof normalizePermissionsConfig>} config
 * @returns {string[]}
 */
export function planerEntraGroupIds(config) {
    const c = normalizePermissionsConfig(config || loadPermissionsConfig());
    const ids = [c.groupAdminId, c.groupLehrerId, c.groupSchuelerId]
        .map((id) => String(id || '').trim())
        .filter((id) => GUID_RE.test(id));
    return [...new Set(ids)];
}

/**
 * @param {string[]} groupIds
 * @returns {Promise<Set<string>>}
 */
export async function fetchUserMemberGroupIds(groupIds) {
    const ids = (groupIds || []).filter((id) => GUID_RE.test(String(id || '').trim()));
    if (!ids.length) return new Set();
    const want = new Set(ids.map((id) => String(id).toLowerCase()));
    try {
        const tok = await getGraphToken(CHECK_SCOPES);
        const data = await graphJson('POST', '/me/checkMemberGroups', tok, { groupIds: ids });
        const matched = (data && data.value) || [];
        const out = new Set(matched.map((g) => String(g).toLowerCase()));
        if (out.size) return out;
    } catch {
        /* fallback memberOf */
    }
    try {
        const tok = await getGraphToken(CHECK_SCOPES);
        const page = await fetchAllPages(
            tok,
            '/me/memberOf/microsoft.graph.group?$select=id',
            { maxItems: 512, maxPages: 8 }
        );
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
 * @param {Set<string>|string[]} memberIds
 * @param {ReturnType<typeof normalizePermissionsConfig>} [config]
 * @returns {PlanerRole[]}
 */
export function listRolesFromEntraGroups(memberIds, config) {
    const c = normalizePermissionsConfig(config || loadPermissionsConfig());
    const member = memberIds instanceof Set ? memberIds : new Set((memberIds || []).map((x) => String(x).toLowerCase()));
    const inGroup = (idKey) => {
        const id = String(c[idKey] || '').trim().toLowerCase();
        return id && member.has(id);
    };
    const roles = [];
    if (inGroup('groupAdminId')) roles.push('admin');
    if (inGroup('groupLehrerId')) roles.push('lehrer');
    if (inGroup('groupSchuelerId')) roles.push('schueler');
    return roles;
}

/**
 * @param {Set<string>|string[]} memberIds
 * @param {ReturnType<typeof normalizePermissionsConfig>} [config]
 * @returns {PlanerRole|''}
 */
export function pickRoleFromEntraGroups(memberIds, config) {
    const roles = listRolesFromEntraGroups(memberIds, config);
    return roles[0] || '';
}

/**
 * @param {{ teacherMatch?: object|null, studentMatch?: object|null }}
 * @returns {PlanerRole[]}
 */
export function listRolesFromStammdaten(scope) {
    const roles = [];
    if (scope && scope.teacherMatch) roles.push('lehrer');
    if (scope && scope.studentMatch) roles.push('schueler');
    return roles;
}

/**
 * @param {{ teacherMatch?: object|null, studentMatch?: object|null }} scope
 * @returns {PlanerRole|''}
 */
export function pickRoleFromStammdaten(scope) {
    const roles = listRolesFromStammdaten(scope);
    return roles[0] || '';
}

/**
 * @param {PlanerRole|string} role
 */
export function roleSourceLabel(source) {
    const s = String(source || '');
    if (s === 'entra') return 'Entra-Gruppe';
    if (s === 'global-admin') return 'Globaler Administrator';
    if (s === 'admin-schueler') return 'Schüler-Ansicht (Admin)';
    if (s === 'stammdaten') return 'Stammdaten';
    if (s === 'admin-user') return 'Verwaltung (Einzelperson)';
    if (s === 'lehrer-user') return 'Lehrkraft (Einzelperson)';
    if (s === 'schueler-user') return 'Schüler/in (Einzelperson)';
    if (s === 'demo') return 'Demo';
    return '';
}

/**
 * @param {Array<{ roleTemplateId?: string }>} directoryRoles
 * @param {string[]} [templateIds]
 */
export function hasGlobalAdministratorDirectoryRole(directoryRoles, templateIds) {
    const allow = new Set(
        (templateIds && templateIds.length ? templateIds : [GLOBAL_ADMINISTRATOR_ROLE_TEMPLATE_ID]).map((id) =>
            String(id).toLowerCase()
        )
    );
    return (directoryRoles || []).some((role) =>
        allow.has(String(role && role.roleTemplateId ? role.roleTemplateId : '').toLowerCase())
    );
}

function cacheKeyForAccount() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            const oid = String((info && info.oid) || '').trim().toLowerCase();
            const upn = String((info && (info.upn || info.username)) || '').trim().toLowerCase();
            return oid || upn || '';
        }
    } catch {
        /* ignore */
    }
    return '';
}

/**
 * @returns {Promise<boolean>}
 */
export async function userIsEntraGlobalAdministrator() {
    const key = cacheKeyForAccount();
    if (!key) return false;
    const now = Date.now();
    if (globalAdminCache && globalAdminCache.key === key && now - globalAdminCache.at <= GLOBAL_ADMIN_CACHE_MS) {
        return globalAdminCache.value;
    }
    try {
        const tok = await getGraphToken(GLOBAL_ADMIN_ROLE_SCOPES);
        const page = await fetchAllPages(
            tok,
            '/me/transitiveMemberOf/microsoft.graph.directoryRole?$select=roleTemplateId,displayName',
            { maxItems: 64, maxPages: 3 }
        );
        const ok = hasGlobalAdministratorDirectoryRole(page.items);
        globalAdminCache = { key, at: now, value: ok };
        return ok;
    } catch {
        globalAdminCache = { key, at: now, value: false };
        return false;
    }
}

/**
 * @param {{ teacherMatch?: object|null, studentMatch?: object|null }} scope
 * @param {ReturnType<typeof normalizePermissionsConfig>} config
 * @param {boolean} entra
 * @param {boolean} [skipEntraGroups]
 * @returns {Promise<{ roles: PlanerRole[], sources: Record<string, string> }>}
 */
async function collectPlanerRolesForUser(scope, config, entra, skipEntraGroups) {
    /** @type {PlanerRole[]} */
    const roles = [];
    /** @type {Record<string, string>} */
    const sources = {};

    const add = (role, source) => {
        if (!roles.includes(role)) roles.push(role);
        if (!sources[role]) sources[role] = source;
    };

    if (await userIsEntraGlobalAdministrator()) {
        add('admin', 'global-admin');
    }

    if (entra && !skipEntraGroups) {
        const memberIds = await fetchUserMemberGroupIds(planerEntraGroupIds(config));
        listRolesFromEntraGroups(memberIds, config).forEach((r) => add(r, 'entra'));
    }

    listRolesFromStammdaten(scope).forEach((r) => add(r, 'stammdaten'));

    const mail = scope && scope.accountEmail;
    if (accountIsPlannerUserInList(mail, config.adminUsers)) add('admin', 'admin-user');
    if (accountIsPlannerUserInList(mail, config.lehrerUsers)) add('lehrer', 'lehrer-user');
    if (accountIsPlannerUserInList(mail, config.schuelerUsers)) add('schueler', 'schueler-user');

    return finalizePlanerRoles(roles, sources);
}

/**
 * Admins dürfen immer in die Schüler-Ansicht wechseln (Klasse ggf. manuell wählen).
 * @param {PlanerRole[]} roles
 * @param {Record<string, string>} sources
 */
export function finalizePlanerRoles(roles, sources) {
    const list = sortPlanerRoles(roles || []);
    const src = { ...(sources || {}) };
    if (list.includes('admin') && !list.includes('schueler')) {
        list.push('schueler');
        if (!src.schueler) src.schueler = 'admin-schueler';
    }
    return { roles: sortPlanerRoles(list), sources: src };
}

function applyRolesToState(state, roles, sources, opts) {
    const finalized = finalizePlanerRoles(roles, sources);
    state.planerRoles = finalized.roles;
    state.planerRoleSources = finalized.sources;
    state.role = resolveActivePlanerRole(finalized.roles, {
        preferredActiveRole: opts && opts.preferredActiveRole,
        preferredDemoRole: opts && opts.preferredDemoRole
    });
    state.roleSource = state.planerRoleSources[state.role] || state.roleSource || 'entra';
    state.planerAccessDenied = finalized.roles.length === 0;
}

/**
 * Rollen-Umschalter nur für IT-Vorschau (?demoRole=1), nicht wenn Entra fehlt.
 */
export function isPlanerDemoRoleUiEnabled(_entraConfigured, demoRoleOverride) {
    return !!demoRoleOverride;
}

/**
 * Lehrkräfte/Schüler: keine Rolle wählen. Nur Planer-Admins mit mehreren Rollen oder Demo-Vorschau.
 * @param {PlanerRole[]} planerRoles
 * @param {boolean} demoRoleUi
 */
export function canUsePlanerRoleSwitcher(planerRoles, demoRoleUi) {
    if (demoRoleUi) return true;
    const roles = sortPlanerRoles(planerRoles || []);
    return roles.includes('admin') && roles.length > 1;
}

/** Stammdaten / Listen einrichten – nur echte Planer-Verwaltung in Admin-Rolle. */
export function canShowPlanerItToolbar(planerRoles, activeRole) {
    const roles = sortPlanerRoles(planerRoles || []);
    return roles.includes('admin') && activeRole === 'admin';
}

/**
 * @param {boolean} [demoRoleOverride]
 */
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
 * Setzt Rolle/Hinweise auf dem Planer-State (mutiert state).
 * @param {object} state
 * @param {{ demoRoleOverride?: boolean, preferredDemoRole?: PlanerRole|null, preferredActiveRole?: PlanerRole|null }} [opts]
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
        const demoRoles = ['admin', 'lehrer', 'schueler'];
        applyRolesToState(
            state,
            demoRoles,
            { admin: 'demo', lehrer: 'demo', schueler: 'demo' },
            pickOpts
        );
        state.roleSource = 'demo';
        state.planerAccessDenied = false;
        state.roleHint = '';
        return;
    }

    if (!isLoggedIn()) {
        state.planerRoles = [];
        state.planerRoleSources = {};
        if (entra) {
            state.roleSource = 'entra';
            state.planerAccessDenied = false;
            state.roleHint =
                'Bitte mit Ihrem Schul-Microsoft-Konto anmelden (links unten). Ihre Rolle im Planer richtet sich nach den Entra-Gruppen unter Administration → SharePoint-Berechtigungen.';
        }
        return;
    }

    const scope = {
        teacherMatch: state.teacherMatch,
        studentMatch: state.studentMatch,
        accountEmail: state.accountEmail
    };

    if (!entra) {
        const collected = await collectPlanerRolesForUser(scope, config, false, true);
        if (collected.roles.length) {
            applyRolesToState(state, collected.roles, collected.sources, pickOpts);
            state.planerAccessDenied = false;
            state.roleHint =
                'Entra-Gruppen im Planer-Setup fehlen – Rolle vorläufig aus Stammdaten oder Einzelpersonen. Bitte unter Administration → SharePoint-Berechtigungen Gruppen eintragen.';
            return;
        }
        state.role = 'lehrer';
        state.planerRoles = [];
        state.planerRoleSources = {};
        state.planerAccessDenied = true;
        state.roleHint =
            'Keine Planer-Berechtigung: Entra-Gruppen fehlen und keine Zuordnung als Lehrkraft/Schüler/Verwaltung. Bitte wenden Sie sich an die IT.';
        return;
    }

    try {
        const collected = await collectPlanerRolesForUser(scope, config, entra, false);

        if (collected.roles.length) {
            applyRolesToState(state, collected.roles, collected.sources, pickOpts);
            state.planerAccessDenied = false;
            state.roleHint =
                collected.roles.length > 1
                    ? ''
                    : collected.sources[state.role] === 'stammdaten'
                      ? 'Rolle aus Stammdaten (E-Mail-Zuordnung).'
                      : '';
            return;
        }

        state.role = 'lehrer';
        state.roleSource = 'entra';
        state.planerRoles = [];
        state.planerRoleSources = {};
        state.planerAccessDenied = true;
        state.roleHint =
            'Keine Planer-Berechtigung: Sie sind in keiner konfigurierten Entra-Gruppe und nicht als Einzelperson eingetragen (Verwaltung, Lehrkräfte oder Schüler). Bitte wenden Sie sich an die IT.';
    } catch (e) {
        const collected = await collectPlanerRolesForUser(scope, config, entra, true);
        if (collected.roles.length) {
            applyRolesToState(state, collected.roles, collected.sources, pickOpts);
            state.planerAccessDenied = false;
            state.roleHint =
                'Entra-Gruppen konnten nicht geprüft werden – vorläufig aus anderen Quellen. ' +
                (e && e.message ? e.message : String(e));
            return;
        }
        state.planerAccessDenied = false;
        state.roleHint =
            'Entra-Gruppen konnten nicht geprüft werden. ' + (e && e.message ? e.message : String(e));
    }
}
