/**
 * Lehrer-Freistellung: Rollen aus Entra-Gruppen / Setup-Mails / Stammdaten.
 * Demo-Umschalter nur mit ?demoRole=1 (docs/rollen-vertrag.md).
 */
import {
    canUseAppRoleSwitcher,
    isDemoRoleUiEnabled,
    readDemoRoleOverrideFromUrl
} from '../../shared/app-role-switcher-policy.js';
import { loadSchoolAudienceGroups } from '../../shared/school-audience-groups.js';
import {
    fetchUserMemberGroupIds,
    userIsEntraGlobalAdministrator
} from '../schularbeiten-planer/schularbeiten-planer-entra-role.js';
import {
    accountIsLfrDirektionAccess,
    accountIsLfrLehrerAccess,
    loadLfrPermissions,
    resolveLfrApproverEmail
} from './lfr-permissions.js';
import { matchTeacherByEmail, persistRole, resolveRole } from './lfr-state.js';

/** @typedef {'direktion'|'lehrer'} LfrRole */

export const LFR_ROLE_ORDER = ['direktion', 'lehrer'];

/**
 * @param {ReturnType<typeof loadLfrPermissions>} [perms]
 */
export function entraGroupsConfigured(perms) {
    const p = perms || loadLfrPermissions();
    return !!(
        p.groupLehrerId ||
        p.groupDirektionId ||
        (p.lehrerUsers && p.lehrerUsers.length) ||
        (p.direktionUsers && p.direktionUsers.length) ||
        resolveLfrApproverEmail(p)
    );
}

/**
 * @param {LfrRole[]|string[]} roles
 * @returns {LfrRole[]}
 */
export function sortLfrRoles(roles) {
    const set = new Set((roles || []).map((r) => String(r)));
    return LFR_ROLE_ORDER.filter((r) => set.has(r));
}

/**
 * @param {LfrRole[]} available
 * @param {{ preferredActiveRole?: LfrRole|null, preferredDemoRole?: LfrRole|null }} [opts]
 */
export function resolveActiveLfrRole(available, opts) {
    const list = sortLfrRoles(available || []);
    const set = new Set(list);
    const cur = opts && opts.preferredActiveRole;
    const demo = opts && opts.preferredDemoRole;
    if (cur && set.has(cur)) return cur;
    if (demo && set.has(demo)) return demo;
    const stored = resolveRole();
    if (stored && set.has(stored)) return stored;
    return list[0] || 'lehrer';
}

/**
 * @param {boolean} [_entraConfigured]
 * @param {boolean} [demoRoleOverride]
 */
export function isLfrDemoRoleUiEnabled(_entraConfigured, demoRoleOverride) {
    return isDemoRoleUiEnabled(demoRoleOverride);
}

export { readDemoRoleOverrideFromUrl };

/**
 * @param {LfrRole[]} roles
 * @param {boolean} demoRoleUi
 */
export function canUseLfrRoleSwitcher(roles, demoRoleUi) {
    return canUseAppRoleSwitcher(roles, {
        demoRoleUi,
        privilegedRoles: ['direktion']
    });
}

/**
 * @param {ReturnType<typeof loadLfrPermissions>} perms
 * @returns {string[]}
 */
function groupIdsForCheck(perms) {
    /** @type {string[]} */
    const ids = [];
    if (perms.groupDirektionId) ids.push(perms.groupDirektionId);
    if (perms.groupLehrerId) ids.push(perms.groupLehrerId);
    try {
        const stamm = loadSchoolAudienceGroups();
        if (stamm && stamm.groupLehrerId) ids.push(stamm.groupLehrerId);
    } catch {
        /* ignore */
    }
    return [...new Set(ids.map((x) => String(x).trim()).filter(Boolean))];
}

/**
 * @param {Set<string>} member
 * @param {ReturnType<typeof loadLfrPermissions>} perms
 * @returns {LfrRole[]}
 */
export function listRolesFromEntraGroups(member, perms) {
    /** @type {LfrRole[]} */
    const roles = [];
    const dirId = String(perms.groupDirektionId || '')
        .trim()
        .toLowerCase();
    const lehrerId = String(perms.groupLehrerId || '')
        .trim()
        .toLowerCase();
    let stammLehrer = '';
    try {
        const stamm = loadSchoolAudienceGroups();
        stammLehrer = String((stamm && stamm.groupLehrerId) || '')
            .trim()
            .toLowerCase();
    } catch {
        /* ignore */
    }
    if (dirId && member.has(dirId)) roles.push('direktion');
    if ((lehrerId && member.has(lehrerId)) || (stammLehrer && member.has(stammLehrer))) {
        roles.push('lehrer');
    }
    return sortLfrRoles(roles);
}

/**
 * @param {object} state
 * @param {{ demoRoleOverride?: boolean, preferredDemoRole?: LfrRole|null, preferredActiveRole?: LfrRole|null }} [opts]
 */
export async function applyLfrRoleFromEntra(state, opts) {
    const perms = loadLfrPermissions();
    const entra = entraGroupsConfigured(perms);
    state.entraGroupsConfigured = entra;
    const demoOverride = readDemoRoleOverrideFromUrl(opts && opts.demoRoleOverride);
    state.demoRoleOverride = demoOverride;

    const pickOpts = {
        preferredActiveRole: (opts && opts.preferredActiveRole) || state.role,
        preferredDemoRole: opts && opts.preferredDemoRole
    };

    if (demoOverride) {
        const roles = sortLfrRoles(['direktion', 'lehrer']);
        state.planerRoles = roles;
        state.planerAccessDenied = false;
        state.role = resolveActiveLfrRole(roles, pickOpts);
        persistRole(state.role);
        state.roleHint = 'IT-Vorschau (?demoRole=1).';
        return;
    }

    /** @type {LfrRole[]} */
    let roles = [];
    /** @type {Record<string, string>} */
    const sources = {};

    try {
        if (await userIsEntraGlobalAdministrator()) {
            roles.push('direktion', 'lehrer');
            sources.direktion = 'global-admin';
            sources.lehrer = 'global-admin';
        }
    } catch {
        /* ignore */
    }

    const ids = groupIdsForCheck(perms);
    if (ids.length) {
        try {
            const member = await fetchUserMemberGroupIds(ids);
            const fromEntra = listRolesFromEntraGroups(member, perms);
            fromEntra.forEach((r) => {
                if (!roles.includes(r)) roles.push(r);
                if (!sources[r]) sources[r] = 'entra';
            });
        } catch {
            /* ignore */
        }
    }

    const em = String(state.accountEmail || '')
        .trim()
        .toLowerCase();
    if (accountIsLfrDirektionAccess(em, perms) || resolveLfrApproverEmail(perms) === em) {
        if (!roles.includes('direktion')) roles.push('direktion');
        if (!sources.direktion) sources.direktion = 'setup-direktion';
    }
    if (accountIsLfrLehrerAccess(em, perms)) {
        if (!roles.includes('lehrer')) roles.push('lehrer');
        if (!sources.lehrer) sources.lehrer = 'lehrer-user';
    }
    const match = matchTeacherByEmail(
        (state.stammdaten && state.stammdaten.teachers) || [],
        em
    );
    if (match) {
        if (!roles.includes('lehrer')) roles.push('lehrer');
        if (!sources.lehrer) sources.lehrer = 'stammdaten';
    }

    roles = sortLfrRoles(roles);

    if (!roles.length) {
        // Ohne konfigurierte Gruppen: weicher Lehrer-Default (wie bisher)
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
                'Keine Berechtigung für Lehrer-Freistellungen. IT: Entra-Gruppen im Setup zuordnen. Vorschau: ?demoRole=1';
            return;
        }
    } else {
        state.planerAccessDenied = false;
        state.roleHint = '';
    }

    state.planerRoles = roles;
    state.planerRoleSources = sources;
    state.role = resolveActiveLfrRole(roles, pickOpts);
    persistRole(state.role);
}
