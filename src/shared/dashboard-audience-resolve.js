/**
 * Persona-Signale für das Dashboard.
 *
 * Hierarchie (strikt):
 * 1. Global Admin oder Plattform-Betreiber → voller Katalog + Stammdaten
 * 2. Mitglieder der konfigurierten Dashboard-Lehrer-Entra-Gruppe → Lehrkraft-Ansicht
 * 3. Mitglieder der konfigurierten Dashboard-Schüler-Entra-Gruppe → Schüler-Ansicht
 * 4. Sonst (Planer-Rollen / Stammdaten-E-Mail) → Lehrkraft, wenn erkennbar
 * 5. Jeder andere angemeldete Nutzer → Schüler-Ansicht (Minimum)
 */
import { resolveDashboardPersonaFromEntraGroups } from './dashboard-audience-entra.js';
import { dashboardAudienceGroupsConfigured } from './dashboard-audience-groups-store.js';
import {
    userIsEntraGlobalAdministrator,
    fetchUserMemberGroupIds,
    listRolesFromEntraGroups as listSaRolesFromEntra,
    finalizePlanerRoles as finalizeSaRoles,
    planerEntraGroupIds as saGroupIds,
    entraGroupsConfigured as saEntraConfigured
} from '../tools/schularbeiten-planer/schularbeiten-planer-entra-role.js';
import {
    loadPermissionsConfig as loadSaPermissions,
    normalizePermissionsConfig as normalizeSaPermissions
} from '../tools/schularbeiten-planer/schularbeiten-planer-permissions.js';
import { accountIsPlannerUserInList } from '../tools/freistellung-planer/freistellung-planer-direktion-users.js';
import {
    matchTeacherByEmail,
    matchStudentByEmail
} from '../tools/schularbeiten-planer/schularbeiten-planer-state.js';
import { listRolesFromStammdaten as listSaRolesFromStammdaten } from '../tools/schularbeiten-planer/schularbeiten-planer-entra-role.js';
import {
    entraGroupsConfigured as frEntraConfigured,
    loadPermissionsConfig as loadFrPermissions,
    normalizePermissionsConfig as normalizeFrPermissions
} from '../tools/freistellung-planer/freistellung-planer-permissions.js';
import {
    listRolesFromEntraGroups as listFrRolesFromEntra,
    finalizePlanerRoles as finalizeFrRoles,
    planerEntraGroupIds as frGroupIds
} from '../tools/freistellung-planer/freistellung-planer-entra-role.js';

function accountEmail() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetUserPrincipalName === 'function') {
            return String(window.ms365AuthGetUserPrincipalName() || '').trim().toLowerCase();
        }
    } catch {
        /* ignore */
    }
    return '';
}

function isLoggedIn() {
    try {
        return (
            typeof window !== 'undefined' &&
            typeof window.ms365AuthIsLoggedIn === 'function' &&
            window.ms365AuthIsLoggedIn()
        );
    } catch {
        return false;
    }
}

function loadStammdatenScope() {
    const mail = accountEmail();
    /** @type {{ teacherMatch?: object|null, studentMatch?: object|null }} */
    const scope = { teacherMatch: null, studentMatch: null };
    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad();
            scope.teacherMatch = matchTeacherByEmail(mail, s && s.teachers);
            scope.studentMatch = matchStudentByEmail(mail, s && s.students);
        }
    } catch {
        /* ignore */
    }
    return scope;
}

function isOperatorSync() {
    try {
        return (
            typeof window !== 'undefined' &&
            window.ms365OperatorAccess &&
            typeof window.ms365OperatorAccess.isCurrentUserOperator === 'function' &&
            window.ms365OperatorAccess.isCurrentUserOperator()
        );
    } catch {
        return false;
    }
}

async function resolveOperatorFlag() {
    if (isOperatorSync()) return true;
    try {
        const oa = window.ms365OperatorAccess;
        if (oa && typeof oa.refreshOperatorStatus === 'function') {
            return !!(await oa.refreshOperatorStatus({ force: false }));
        }
    } catch {
        /* ignore */
    }
    return false;
}

async function checkMemberGroups(groupIds) {
    const ids = (groupIds || []).filter(Boolean);
    if (!ids.length) return new Set();
    return fetchUserMemberGroupIds(ids);
}

async function resolveSchularbeitenRoles() {
    const config = normalizeSaPermissions(loadSaPermissions());
    const entra = saEntraConfigured(config);
    const mail = accountEmail();
    /** @type {string[]} */
    const roles = [];
    const sources = {};

    const add = (role, source) => {
        if (!roles.includes(role)) roles.push(role);
        if (!sources[role]) sources[role] = source;
    };

    if (await userIsEntraGlobalAdministrator()) add('admin', 'global-admin');

    if (entra) {
        try {
            const member = await checkMemberGroups(saGroupIds(config));
            listSaRolesFromEntra(member, config).forEach((r) => add(r, 'entra'));
        } catch {
            /* ignore */
        }
    }

    if (accountIsPlannerUserInList(mail, config.adminUsers)) add('admin', 'admin-user');
    if (accountIsPlannerUserInList(mail, config.lehrerUsers)) add('lehrer', 'lehrer-user');
    if (accountIsPlannerUserInList(mail, config.schuelerUsers)) add('schueler', 'schueler-user');

    const stammdatenScope = loadStammdatenScope();
    listSaRolesFromStammdaten(stammdatenScope).forEach((r) => add(r, 'stammdaten'));

    if (!entra && !roles.length) return [];
    return finalizeSaRoles(roles, sources).roles;
}

async function resolveFreistellungRoles() {
    const config = normalizeFrPermissions(loadFrPermissions());
    const entra = frEntraConfigured(config);
    const mail = accountEmail();
    /** @type {string[]} */
    const roles = [];
    const sources = {};

    const add = (role, source) => {
        if (!roles.includes(role)) roles.push(role);
        if (!sources[role]) sources[role] = source;
    };

    if (await userIsEntraGlobalAdministrator()) {
        add('direktion', 'global-admin');
        add('kv', 'global-admin');
    }

    if (entra) {
        try {
            const member = await checkMemberGroups(frGroupIds(config));
            listFrRolesFromEntra(member, config).forEach((r) => add(r, 'entra'));
        } catch {
            /* ignore */
        }
    }

    if (accountIsPlannerUserInList(mail, config.direktionUsers)) add('direktion', 'direktion-user');
    if (accountIsPlannerUserInList(mail, config.kvUsers)) add('kv', 'kv-user');
    if (accountIsPlannerUserInList(mail, config.schuelerUsers)) add('schueler', 'schueler-user');

    const stammdatenScopeFr = loadStammdatenScope();
    if (stammdatenScopeFr.studentMatch) add('schueler', 'stammdaten');

    if (!entra && !roles.length) return [];
    return finalizeFrRoles(roles, sources).roles;
}

/**
 * Planer-„Admin“/Direktion = erweiterte Lehrkraft-Ansicht, kein voller IT-Katalog.
 */
function isTeacherDashboardRole(schularbeitenRoles, freistellungRoles, stammdatenScope) {
    if (schularbeitenRoles.includes('lehrer') || schularbeitenRoles.includes('admin')) return true;
    if (freistellungRoles.includes('kv') || freistellungRoles.includes('direktion')) return true;
    if (stammdatenScope && stammdatenScope.teacherMatch) return true;
    return false;
}

/**
 * @typedef {{
 *   filtering: boolean,
 *   loggedIn: boolean,
 *   unknownUser?: boolean,
 *   isIt: boolean,
 *   isLehrer: boolean,
 *   isSchueler: boolean,
 *   globalAdmin?: boolean,
 *   operator?: boolean,
 *   schularbeitenRoles: string[],
 *   freistellungRoles: string[]
 * }} DashboardPersonas
 */

/**
 * @returns {Promise<DashboardPersonas>}
 */
export async function resolveDashboardPersonas() {
    if (!isLoggedIn()) {
        return {
            filtering: false,
            loggedIn: false,
            isIt: false,
            isLehrer: false,
            isSchueler: false,
            schularbeitenRoles: [],
            freistellungRoles: []
        };
    }

    const [schularbeitenRoles, freistellungRoles, globalAdmin, operator] = await Promise.all([
        resolveSchularbeitenRoles(),
        resolveFreistellungRoles(),
        userIsEntraGlobalAdministrator(),
        resolveOperatorFlag()
    ]);

    const stammdatenScope = loadStammdatenScope();

    /** Nur Global Admin + Plattform-Betreiber: alles sichtbar */
    const fullAccess = globalAdmin || operator;

    if (fullAccess) {
        return {
            filtering: false,
            loggedIn: true,
            isIt: true,
            isLehrer: false,
            isSchueler: false,
            globalAdmin,
            operator,
            schularbeitenRoles,
            freistellungRoles
        };
    }

    const groupsConfigured = dashboardAudienceGroupsConfigured();
    let entraDashPersona = null;
    if (groupsConfigured) {
        try {
            entraDashPersona = await resolveDashboardPersonaFromEntraGroups();
        } catch {
            entraDashPersona = null;
        }
    }

    if (entraDashPersona === 'lehrer') {
        return {
            filtering: true,
            loggedIn: true,
            isIt: false,
            isLehrer: true,
            isSchueler: false,
            personaSource: 'dashboard-entra-lehrer',
            globalAdmin,
            operator,
            schularbeitenRoles,
            freistellungRoles
        };
    }

    if (entraDashPersona === 'schueler' || (groupsConfigured && !entraDashPersona)) {
        return {
            filtering: true,
            loggedIn: true,
            unknownUser: groupsConfigured && !entraDashPersona,
            isIt: false,
            isLehrer: false,
            isSchueler: true,
            personaSource: entraDashPersona === 'schueler' ? 'dashboard-entra-schueler' : 'dashboard-entra-none',
            globalAdmin,
            operator,
            schularbeitenRoles,
            freistellungRoles
        };
    }

    const isTeacher = isTeacherDashboardRole(schularbeitenRoles, freistellungRoles, stammdatenScope);

    const explicitPlanerSchueler =
        schularbeitenRoles.includes('schueler') || freistellungRoles.includes('schueler');

    return {
        filtering: true,
        loggedIn: true,
        unknownUser: !isTeacher && !explicitPlanerSchueler && !stammdatenScope.studentMatch,
        isIt: false,
        isLehrer: isTeacher,
        isSchueler: !isTeacher,
        globalAdmin,
        operator,
        schularbeitenRoles,
        freistellungRoles
    };
}
