/**
 * sRDP-Rollen: Schüler, Lehrkraft, Direktion, Schul-IT.
 * Umschalter nur mit ?demoRole=1.
 */
import {
    canUseAppRoleSwitcher,
    isDemoRoleUiEnabled,
    readDemoRoleOverrideFromUrl as readDemoRoleOverrideFromUrlShared
} from '../../shared/app-role-switcher-policy.js';
import { userIsDesignatedSchoolIt } from '../../shared/designated-school-it.js';
import { loadSchoolAudienceGroups } from '../../shared/school-audience-groups.js';
import {
    fetchUserMemberGroupIds,
    userIsEntraGlobalAdministrator
} from '../schularbeiten-planer/schularbeiten-planer-entra-role.js';
import {
    accountIsPlannerUserInList,
    direktionUsersFromTenantStammdaten
} from '../freistellung-planer/freistellung-planer-direktion-users.js';
import { loadEffectivePermissionsConfig as loadFrPerms } from '../freistellung-planer/freistellung-planer-permissions.js';
import { loadEffectiveSrdpPermissionsConfig } from '../sharepoint/sharepoint-liste-srdp-permissions-ui.js';

export const SRDP_ROLE_ORDER = ['schueler', 'lehrer', 'admin'];

export function readDemoRoleOverrideFromUrl() {
    return readDemoRoleOverrideFromUrlShared(false);
}

/**
 * Role-Chips nur bei explizitem ?demoRole=1.
 * @param {boolean} [_entraConfigured]
 * @param {boolean} [demoOverride]
 */
export function isSrdpDemoRoleUiEnabled(_entraConfigured, demoOverride) {
    return isDemoRoleUiEnabled(demoOverride);
}

/**
 * Produktiv nie umschalten: eine Persona, eine Ansicht.
 * @param {string[]} [_roles]
 * @param {boolean} [demoUi]
 */
export function canUseSrdpRoleSwitcher(_roles, demoUi) {
    if (!demoUi) return false;
    return canUseAppRoleSwitcher(SRDP_ROLE_ORDER, { demoRoleUi: true });
}

export function sortSrdpRoles(roles) {
    const set = new Set((Array.isArray(roles) ? roles : []).map((r) => String(r || '').toLowerCase()));
    return SRDP_ROLE_ORDER.filter((r) => set.has(r));
}

function normEmail(email) {
    return String(email || '')
        .trim()
        .toLowerCase();
}

function emailInPeople(people, email) {
    const em = normEmail(email);
    if (!em) return false;
    return (Array.isArray(people) ? people : []).some(
        (p) => normEmail(p && (p.email || p.upn || p.mail)) === em
    );
}

function portalIsSchoolIt() {
    try {
        const p = typeof window !== 'undefined' ? window.__ms365DashboardPersonasLast : null;
        return !!(p && p.loggedIn && p.isIt);
    } catch {
        return false;
    }
}

function groupIdsOf(groups) {
    return (Array.isArray(groups) ? groups : [])
        .map((g) => String((g && g.groupId) || '').trim())
        .filter(Boolean);
}

/**
 * @param {{ isIt?: boolean, isDirektion?: boolean, isLehrer?: boolean }} flags
 * @returns {'admin'|'lehrer'|'schueler'}
 */
export function roleFromFlags(flags) {
    const f = flags || {};
    if (f.isIt || f.isDirektion) return 'admin';
    if (f.isLehrer) return 'lehrer';
    return 'schueler';
}

/**
 * Heuristik ohne Entra: Lehrkraft aus Stammdaten, sonst Schüler.
 * Direktion/IT setzt nur resolveSrdpRolesForAccount.
 * @param {object} state
 */
export function inferSrdpRolesFromStammdaten(state) {
    const em = normEmail(state && state.accountEmail);
    if (!em) return ['schueler'];
    const teachers = (state.stammdaten && state.stammdaten.teachers) || [];
    if (emailInPeople(teachers, em)) return ['lehrer'];
    return ['schueler'];
}

/**
 * @param {string[]} roles
 * @param {string} [saved]
 * @param {{ allowSaved?: boolean }} [opts]
 */
export function shellPreferredSrdpRole(roles, saved, opts) {
    const list = (Array.isArray(roles) ? roles : []).map((r) => String(r || '').toLowerCase());
    const allowSaved = !!(opts && opts.allowSaved);
    const s = String(saved || '').toLowerCase();
    if (allowSaved && s && list.includes(s)) return s;
    if (list.includes('admin')) return 'admin';
    if (list.includes('lehrer')) return 'lehrer';
    if (list.includes('schueler')) return 'schueler';
    return 'schueler';
}

/**
 * Eine produktive Rolle: Schul-IT und Direktion sehen die Verwaltung,
 * Lehrkräfte die Lehreransicht, alle anderen die Anmeldung.
 * @param {object} state
 * @returns {Promise<{ roles: string[], sources: Record<string, string>, active: string }>}
 */
export async function resolveSrdpRolesForAccount(state) {
    const em = normEmail(state && state.accountEmail);
    /** @type {Record<string, string>} */
    const sources = {};
    let isIt = portalIsSchoolIt();
    let isDirektion = false;
    let isLehrer = emailInPeople((state && state.stammdaten && state.stammdaten.teachers) || [], em);

    if (isIt) sources.admin = 'designated-it';

    try {
        if (await userIsEntraGlobalAdministrator()) {
            isIt = true;
            sources.admin = 'global-admin';
        }
    } catch {
        /* ignore */
    }
    try {
        if (await userIsDesignatedSchoolIt()) {
            isIt = true;
            if (!sources.admin) sources.admin = 'designated-it';
        }
    } catch {
        /* ignore */
    }

    let srdp = {
        adminGroups: [],
        lehrerGroups: [],
        kandidatenGroups: [],
        adminUsers: [],
        lehrerUsers: []
    };
    let fr = { groupDirektionId: '', direktionUsers: [] };
    let stamm = { groupLehrerId: '', groupSchuelerId: '' };
    try {
        srdp = loadEffectiveSrdpPermissionsConfig() || srdp;
    } catch {
        /* ignore */
    }
    try {
        fr = loadFrPerms() || fr;
    } catch {
        /* ignore */
    }
    try {
        stamm = loadSchoolAudienceGroups() || stamm;
    } catch {
        /* ignore */
    }

    const adminIds = groupIdsOf(srdp.adminGroups);
    const lehrerIds = groupIdsOf(srdp.lehrerGroups);
    const dirId = String(fr.groupDirektionId || '').trim();
    const stammLehrer = String(stamm.groupLehrerId || '').trim();
    const watch = [...adminIds, ...lehrerIds];
    if (dirId) watch.push(dirId);
    if (stammLehrer) watch.push(stammLehrer);

    let member = new Set();
    if (watch.length) {
        try {
            member = await fetchUserMemberGroupIds(watch);
        } catch {
            member = new Set();
        }
    }
    const has = (id) => {
        const key = String(id || '').trim().toLowerCase();
        if (!key) return false;
        if (member && typeof member.has === 'function' && member.has(key)) return true;
        if (member && typeof member.has === 'function' && member.has(id)) return true;
        return false;
    };

    if (adminIds.some(has) || accountIsPlannerUserInList(em, srdp.adminUsers || [])) {
        isDirektion = true;
        if (!sources.admin) sources.admin = 'entra';
    }
    if (dirId && has(dirId)) {
        isDirektion = true;
        if (!sources.admin) sources.admin = 'entra';
    }
    if (accountIsPlannerUserInList(em, fr.direktionUsers || [])) {
        isDirektion = true;
        if (!sources.admin) sources.admin = 'direktion-user';
    }
    try {
        if (accountIsPlannerUserInList(em, direktionUsersFromTenantStammdaten())) {
            isDirektion = true;
            if (!sources.admin) sources.admin = 'stammdaten-direktion';
        }
    } catch {
        /* ignore */
    }
    if (
        lehrerIds.some(has) ||
        (stammLehrer && has(stammLehrer)) ||
        accountIsPlannerUserInList(em, srdp.lehrerUsers || [])
    ) {
        isLehrer = true;
        sources.lehrer = 'entra';
    } else if (isLehrer) {
        sources.lehrer = 'stammdaten';
    }

    const active = roleFromFlags({ isIt, isDirektion, isLehrer });
    if (active === 'admin' && isIt && !isDirektion) sources.admin = sources.admin || 'designated-it';
    if (active === 'lehrer' && !sources.lehrer) sources.lehrer = 'stammdaten';
    if (active === 'schueler') sources.schueler = 'default';
    return { roles: [active], sources, active };
}
