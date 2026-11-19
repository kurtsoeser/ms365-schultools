/**
 * Schulaktivitäten: Rollen aus SA-/FR-Gruppen, Stammdaten-KV und Lehrer-Match.
 * Demo nur mit ?demoRole=1.
 */
import {
    canUseAppRoleSwitcher,
    isDemoRoleUiEnabled,
    readDemoRoleOverrideFromUrl
} from '../../shared/app-role-switcher-policy.js';
import { userIsDesignatedSchoolIt } from '../../shared/designated-school-it.js';
import { readItShellChoice } from '../../shared/dashboard-layout-mode.js';
import { loadSchoolAudienceGroups } from '../../shared/school-audience-groups.js';
import {
    fetchUserMemberGroupIds,
    userIsEntraGlobalAdministrator
} from '../schularbeiten-planer/schularbeiten-planer-entra-role.js';
import { loadEffectivePermissionsConfig as loadSaPerms } from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import {
    loadEffectivePermissionsConfig as loadFrPerms
} from '../freistellung-planer/freistellung-planer-permissions.js';
import { loadEffectiveAktPermissionsConfig as loadAktPerms } from './schulaktivitaeten-planer-permissions.js';
import {
    accountIsPlannerUserInList,
    direktionUsersFromTenantStammdaten
} from '../freistellung-planer/freistellung-planer-direktion-users.js';
import {
    kvKlasseCodesForEmail,
    matchTeacherByEmail,
    persistRole,
    resolveRole
} from './schulaktivitaeten-planer-state.js';

/** @typedef {'admin'|'kv'|'lehrer'} AktRole */

export const AKT_ROLE_ORDER = ['admin', 'kv', 'lehrer'];

export { readDemoRoleOverrideFromUrl };

/**
 * @returns {boolean}
 */
export function aktEntraGroupsConfigured() {
    try {
        const akt = loadAktPerms();
        if (akt.groupAdminId || akt.groupLehrerId) return true;
        if ((akt.adminUsers && akt.adminUsers.length) || (akt.lehrerUsers && akt.lehrerUsers.length)) {
            return true;
        }
    } catch {
        /* ignore */
    }
    try {
        const sa = loadSaPerms();
        if (sa.groupAdminId || sa.groupLehrerId) return true;
    } catch {
        /* ignore */
    }
    try {
        const fr = loadFrPerms();
        if (fr.groupKvId || fr.groupDirektionId) return true;
    } catch {
        /* ignore */
    }
    try {
        const stamm = loadSchoolAudienceGroups();
        if (stamm.groupLehrerId) return true;
    } catch {
        /* ignore */
    }
    return false;
}

/**
 * @param {AktRole[]|string[]} roles
 * @returns {AktRole[]}
 */
export function sortAktRoles(roles) {
    const set = new Set((roles || []).map((r) => String(r)));
    return AKT_ROLE_ORDER.filter((r) => set.has(r));
}

function currentPortalIsSchulIt() {
    try {
        const p =
            typeof window !== 'undefined' ? window.__ms365DashboardPersonasLast : null;
        return !!(p && p.loggedIn && p.isIt);
    } catch {
        return false;
    }
}

/**
 * Shell-bewusste Standard-App-Rolle bei Mehrfachrollen.
 * Schul-IT: Portal → Lehrer (Alltag), IT-Backend → Admin.
 * Direktion / Genehmiger (Admin, aber nicht Schul-IT): immer Admin – sonst keine Freigabe-Inbox.
 * @param {AktRole[]|string[]} available
 * @returns {AktRole|null}
 */
export function shellPreferredAktRole(available) {
    const set = new Set(sortAktRoles(available || []));
    if (!set.size) return null;
    let shell = 'portal';
    try {
        shell = readItShellChoice();
    } catch {
        shell = 'portal';
    }
    if (shell === 'full' && set.has('admin')) return 'admin';
    // Portal-Vorschau nur für Schul-IT – Direktion darf nicht auf Lehrer-Alltag fallen
    if (shell !== 'full' && currentPortalIsSchulIt() && set.has('lehrer')) return 'lehrer';
    if (set.has('admin')) return 'admin';
    if (set.has('lehrer')) return 'lehrer';
    if (set.has('kv')) return 'kv';
    return null;
}

/**
 * @param {AktRole[]} available
 * @param {{ preferredActiveRole?: AktRole|null, preferredDemoRole?: AktRole|null, preferShellDefault?: boolean }} [opts]
 */
export function resolveActiveAktRole(available, opts) {
    const list = sortAktRoles(available || []);
    const set = new Set(list);
    const cur = opts && opts.preferredActiveRole;
    const demo = opts && opts.preferredDemoRole;
    if (cur && set.has(cur)) return cur;
    if (demo && set.has(demo)) return demo;
    if (opts && opts.preferShellDefault) {
        const shellRole = shellPreferredAktRole(list);
        if (shellRole && set.has(shellRole)) return shellRole;
    }
    const stored = resolveRole();
    if (stored && set.has(stored)) return stored;
    const softShell = shellPreferredAktRole(list);
    if (softShell && set.has(softShell)) return softShell;
    return list[0] || 'lehrer';
}

export function isAktDemoRoleUiEnabled(_entraConfigured, demoRoleOverride) {
    return isDemoRoleUiEnabled(demoRoleOverride);
}

export function canUseAktRoleSwitcher(roles, demoRoleUi) {
    // Nie Admin/KV/Lehrer-Chips in der produktiven Lehrerwelt.
    // Nur explizite IT-Vorschau mit ?demoRole=1.
    if (!demoRoleUi) return false;
    return canUseAppRoleSwitcher(roles, { demoRoleUi: true });
}

export function roleSourceLabel(source) {
    const s = String(source || '');
    if (s === 'entra') return 'Entra-Gruppe';
    if (s === 'global-admin') return 'Globaler Administrator';
    if (s === 'designated-it') return 'Schul-IT';
    if (s === 'stammdaten-kv') return 'Klassenvorstand (Stammdaten)';
    if (s === 'stammdaten-direktion') return 'Schulleitung (Stammdaten)';
    if (s === 'stammdaten') return 'Stammdaten';
    if (s === 'admin-user') return 'Verwaltung (Einzelperson)';
    if (s === 'direktion-user') return 'Direktion (Freistellung)';
    if (s === 'demo') return 'IT-Vorschau';
    return s || '–';
}

/**
 * @param {object} state
 * @param {{ demoRoleOverride?: boolean, preferredDemoRole?: AktRole|null, preferredActiveRole?: AktRole|null, preferShellDefault?: boolean }} [opts]
 */
export async function applyAktRoleFromEntra(state, opts) {
    const entra = aktEntraGroupsConfigured();
    state.entraGroupsConfigured = entra;
    const demoOverride = readDemoRoleOverrideFromUrl(opts && opts.demoRoleOverride);
    state.demoRoleOverride = demoOverride;

    const preferShellDefault = !!(opts && opts.preferShellDefault);
    const hasPreferred =
        opts && Object.prototype.hasOwnProperty.call(opts, 'preferredActiveRole');
    const pickOpts = {
        preferredActiveRole: hasPreferred
            ? opts.preferredActiveRole
            : preferShellDefault
              ? null
              : state.role,
        preferredDemoRole: opts && opts.preferredDemoRole,
        preferShellDefault
    };

    if (demoOverride) {
        const roles = sortAktRoles(['admin', 'kv', 'lehrer']);
        state.planerRoles = roles;
        state.planerRoleSources = { admin: 'demo', kv: 'demo', lehrer: 'demo' };
        state.planerAccessDenied = false;
        state.role = resolveActiveAktRole(roles, pickOpts);
        persistRole(state.role);
        state.roleHint = 'IT-Vorschau (?demoRole=1).';
        return;
    }

    /** @type {AktRole[]} */
    let roles = [];
    /** @type {Record<string, string>} */
    const sources = {};

    let sa = { groupAdminId: '', groupLehrerId: '', adminUsers: [], lehrerUsers: [] };
    let akt = { groupAdminId: '', groupLehrerId: '', adminUsers: [], lehrerUsers: [] };
    let fr = {
        groupKvId: '',
        groupDirektionId: '',
        kvUsers: [],
        direktionUsers: []
    };
    try {
        sa = loadSaPerms();
    } catch {
        /* ignore */
    }
    try {
        akt = loadAktPerms();
    } catch {
        /* ignore */
    }
    try {
        fr = loadFrPerms();
    } catch {
        /* ignore */
    }
    let stammLehrer = '';
    try {
        stammLehrer = String(loadSchoolAudienceGroups().groupLehrerId || '').trim();
    } catch {
        /* ignore */
    }

    try {
        if (await userIsEntraGlobalAdministrator()) {
            roles.push('admin', 'kv', 'lehrer');
            sources.admin = 'global-admin';
            sources.kv = 'global-admin';
            sources.lehrer = 'global-admin';
        }
    } catch {
        /* ignore */
    }

    try {
        if (await userIsDesignatedSchoolIt()) {
            if (!roles.includes('admin')) roles.push('admin');
            if (!sources.admin) sources.admin = 'designated-it';
        }
    } catch {
        /* ignore */
    }

    const ids = [
        sa.groupAdminId,
        sa.groupLehrerId,
        akt.groupAdminId,
        akt.groupLehrerId,
        fr.groupKvId,
        fr.groupDirektionId,
        stammLehrer
    ].filter(Boolean);
    if (ids.length) {
        try {
            const member = await fetchUserMemberGroupIds(ids);
            const adminId = String(sa.groupAdminId || akt.groupAdminId || '').toLowerCase();
            const dirId = String(fr.groupDirektionId || '').toLowerCase();
            const lehrerId = String(sa.groupLehrerId || akt.groupLehrerId || '').toLowerCase();
            const stammId = String(stammLehrer || '').toLowerCase();
            const kvId = String(fr.groupKvId || '').toLowerCase();
            if ((adminId && member.has(adminId)) || (dirId && member.has(dirId))) {
                if (!roles.includes('admin')) roles.push('admin');
                if (!sources.admin) sources.admin = 'entra';
            }
            if (kvId && member.has(kvId)) {
                if (!roles.includes('kv')) roles.push('kv');
                if (!sources.kv) sources.kv = 'entra';
            }
            if ((lehrerId && member.has(lehrerId)) || (stammId && member.has(stammId))) {
                if (!roles.includes('lehrer')) roles.push('lehrer');
                if (!sources.lehrer) sources.lehrer = 'entra';
            }
        } catch {
            /* ignore */
        }
    }

    const em = String(state.accountEmail || '')
        .trim()
        .toLowerCase();
    if (
        accountIsPlannerUserInList(em, sa.adminUsers || []) ||
        accountIsPlannerUserInList(em, akt.adminUsers || [])
    ) {
        if (!roles.includes('admin')) roles.push('admin');
        if (!sources.admin) sources.admin = 'admin-user';
    }
    if (
        accountIsPlannerUserInList(em, sa.lehrerUsers || []) ||
        accountIsPlannerUserInList(em, akt.lehrerUsers || [])
    ) {
        if (!roles.includes('lehrer')) roles.push('lehrer');
        if (!sources.lehrer) sources.lehrer = 'lehrer-user';
    }
    if (accountIsPlannerUserInList(em, fr.kvUsers || [])) {
        if (!roles.includes('kv')) roles.push('kv');
        if (!sources.kv) sources.kv = 'kv-user';
    }
    if (accountIsPlannerUserInList(em, fr.direktionUsers || [])) {
        if (!roles.includes('admin')) roles.push('admin');
        if (!sources.admin) sources.admin = 'direktion-user';
    }

    const kvCodes = kvKlasseCodesForEmail(state.stammdaten, em);
    if (kvCodes && kvCodes.length) {
        if (!roles.includes('kv')) roles.push('kv');
        if (!sources.kv) sources.kv = 'stammdaten-kv';
    }
    const teacher = matchTeacherByEmail((state.stammdaten && state.stammdaten.teachers) || [], em);
    if (teacher) {
        if (!roles.includes('lehrer')) roles.push('lehrer');
        if (!sources.lehrer) sources.lehrer = 'stammdaten';
    }
    // Schulleitung / Direktion aus Stammdaten → Genehmiger (App-Admin), unabhängig von Lehrerliste
    try {
        const direktionStamm = direktionUsersFromTenantStammdaten();
        if (accountIsPlannerUserInList(em, direktionStamm)) {
            if (!roles.includes('admin')) roles.push('admin');
            if (!sources.admin) sources.admin = 'stammdaten-direktion';
        }
    } catch {
        /* ignore */
    }

    // Admin sieht auch KV-/Lehrer-Views (ohne Fake-Identität)
    if (roles.includes('admin')) {
        if (!roles.includes('kv')) roles.push('kv');
        if (!sources.kv) sources.kv = 'admin';
        if (!roles.includes('lehrer')) roles.push('lehrer');
        if (!sources.lehrer) sources.lehrer = 'admin';
    }

    roles = sortAktRoles(roles);

    if (!roles.length) {
        if (!entra) {
            roles = ['lehrer'];
            sources.lehrer = 'default';
            state.planerAccessDenied = false;
            state.roleHint = '';
        } else {
            state.planerRoles = [];
            state.planerAccessDenied = true;
            state.role = 'lehrer';
            state.roleHint =
                'Keine Berechtigung für Schulaktivitäten. IT: Lehrer-/Admin-Gruppen (Schularbeiten) bzw. KV-Stammdaten. Vorschau: ?demoRole=1';
            persistRole(state.role);
            return;
        }
    } else {
        state.planerAccessDenied = false;
        state.roleHint = '';
    }

    state.planerRoles = roles;
    state.planerRoleSources = sources;
    state.role = resolveActiveAktRole(roles, pickOpts);
    persistRole(state.role);
}
