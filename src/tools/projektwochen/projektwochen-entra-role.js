/**
 * Projektwochen: Rollen aus Schul-Sammelgruppen / SA-Admin / Stammdaten.
 * Demo nur mit ?demoRole=1.
 */
import {
    canUseAppRoleSwitcher,
    isDemoRoleUiEnabled,
    readDemoRoleOverrideFromUrl
} from '../../shared/app-role-switcher-policy.js';
import { loadSchoolAudienceGroups } from '../../shared/school-audience-groups.js';
import { userIsDesignatedSchoolIt } from '../../shared/designated-school-it.js';
import {
    fetchUserMemberGroupIds,
    userIsEntraGlobalAdministrator
} from '../schularbeiten-planer/schularbeiten-planer-entra-role.js';
import { loadEffectivePermissionsConfig as loadSaPerms } from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import { persistRole, loadRole } from './projektwochen-state.js';

/** @typedef {'admin'|'lehrer'|'schueler'} PwRole */

export const PW_ROLE_ORDER = ['admin', 'lehrer', 'schueler'];

export { readDemoRoleOverrideFromUrl };

export function pwEntraGroupsConfigured() {
    try {
        const sa = loadSaPerms();
        if (sa.groupAdminId || sa.groupLehrerId || sa.groupSchuelerId) return true;
    } catch {
        /* ignore */
    }
    try {
        const stamm = loadSchoolAudienceGroups();
        if (stamm.groupLehrerId || stamm.groupSchuelerId) return true;
    } catch {
        /* ignore */
    }
    return false;
}

export function sortPwRoles(roles) {
    const set = new Set((roles || []).map((r) => String(r)));
    return PW_ROLE_ORDER.filter((r) => set.has(r));
}

export function resolveActivePwRole(available, opts) {
    const list = sortPwRoles(available || []);
    const set = new Set(list);
    const cur = opts && opts.preferredActiveRole;
    const demo = opts && opts.preferredDemoRole;
    if (cur && set.has(cur)) return cur;
    if (demo && set.has(demo)) return demo;
    const stored = loadRole();
    if (stored && set.has(stored)) return stored;
    return list[0] || 'lehrer';
}

export function isPwDemoRoleUiEnabled(_entraConfigured, demoRoleOverride) {
    return isDemoRoleUiEnabled(demoRoleOverride);
}

export function canUsePwRoleSwitcher(roles, demoRoleUi) {
    return canUseAppRoleSwitcher(roles, {
        demoRoleUi,
        privilegedRoles: ['admin']
    });
}

/**
 * @param {object} state
 * @param {{ demoRoleOverride?: boolean, preferredDemoRole?: PwRole|null, preferredActiveRole?: PwRole|null }} [opts]
 */
export async function applyPwRoleFromEntra(state, opts) {
    const entra = pwEntraGroupsConfigured();
    state.entraGroupsConfigured = entra;
    const demoOverride = readDemoRoleOverrideFromUrl(opts && opts.demoRoleOverride);
    state.demoRoleOverride = demoOverride;

    const pickOpts = {
        preferredActiveRole: (opts && opts.preferredActiveRole) || state.role,
        preferredDemoRole: opts && opts.preferredDemoRole
    };

    if (demoOverride) {
        const roles = sortPwRoles(['admin', 'lehrer', 'schueler']);
        state.planerRoles = roles;
        state.planerAccessDenied = false;
        state.role = resolveActivePwRole(roles, pickOpts);
        persistRole(state.role);
        state.roleHint = 'IT-Vorschau (?demoRole=1).';
        return;
    }

    /** @type {PwRole[]} */
    let roles = [];
    let sa = { groupAdminId: '', groupLehrerId: '', groupSchuelerId: '' };
    try {
        sa = loadSaPerms();
    } catch {
        /* ignore */
    }
    let stamm = { groupLehrerId: '', groupSchuelerId: '' };
    try {
        stamm = loadSchoolAudienceGroups();
    } catch {
        /* ignore */
    }

    try {
        if (await userIsEntraGlobalAdministrator()) {
            roles.push('admin', 'lehrer');
        }
    } catch {
        /* ignore */
    }
    try {
        if (await userIsDesignatedSchoolIt()) {
            if (!roles.includes('admin')) roles.push('admin');
            if (!roles.includes('lehrer')) roles.push('lehrer');
        }
    } catch {
        /* ignore */
    }

    const ids = [
        sa.groupAdminId,
        sa.groupLehrerId,
        sa.groupSchuelerId,
        stamm.groupLehrerId,
        stamm.groupSchuelerId
    ].filter(Boolean);
    if (ids.length) {
        try {
            const member = await fetchUserMemberGroupIds(ids);
            const adminId = String(sa.groupAdminId || '').toLowerCase();
            const lehrerId = String(sa.groupLehrerId || stamm.groupLehrerId || '').toLowerCase();
            const schId = String(sa.groupSchuelerId || stamm.groupSchuelerId || '').toLowerCase();
            if (adminId && member.has(adminId) && !roles.includes('admin')) roles.push('admin');
            if (lehrerId && member.has(lehrerId) && !roles.includes('lehrer')) roles.push('lehrer');
            if (schId && member.has(schId) && !roles.includes('schueler')) roles.push('schueler');
        } catch {
            /* ignore */
        }
    }

    if (state.teacherMatch && !roles.includes('lehrer')) roles.push('lehrer');

    if (roles.includes('admin') && !roles.includes('lehrer')) roles.push('lehrer');

    roles = sortPwRoles(roles);

    if (!roles.length) {
        if (!entra) {
            roles = ['lehrer'];
            state.planerAccessDenied = false;
            state.roleHint = '';
        } else {
            state.planerRoles = [];
            state.planerAccessDenied = true;
            state.role = 'lehrer';
            state.roleHint =
                'Keine Berechtigung für Projektwochen. IT: Lehrer-/Schüler-Sammelgruppen. Vorschau: ?demoRole=1';
            persistRole(state.role);
            return;
        }
    } else {
        state.planerAccessDenied = false;
        state.roleHint = '';
    }

    state.planerRoles = roles;
    state.role = resolveActivePwRole(roles, pickOpts);
    persistRole(state.role);
}
