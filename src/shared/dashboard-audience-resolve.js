/**
 * Persona-Signale für das Dashboard.
 *
 * Hierarchie (strikt):
 * 1. Global Admin, designierter IT-Kontakt (Lizenz) oder Plattform-Betreiber → IT
 * 2. Mitglieder der konfigurierten Dashboard-Lehrer-Entra-Gruppe → Lehrkraft
 * 3. Planer-Rollen / Stammdaten-Lehrer-E-Mail → Lehrkraft
 *    (auch wenn nur in Schüler-Gruppe oder in keiner Dashboard-Gruppe)
 * 4. Mitglieder der konfigurierten Dashboard-Schüler-Entra-Gruppe → Schüler
 * 5. Explizite Planer-/Stammdaten-Schüler → Schüler
 * 6. Sonst (vor Verzeichnis-Sync) → unknownUser, neutrale Begrüßung (nicht Schüler/in)
 */
import { resolveDashboardPersonaFromEntraGroups } from './dashboard-audience-entra.js';
import { dashboardAudienceGroupsConfigured } from './dashboard-audience-groups-store.js';
import {
    loadEffectiveTeachersForPersona,
    loadLiveSchoolStammdatenCache
} from './live-school-stammdaten-cache.js';
import {
    userIsEntraGlobalAdministrator,
    fetchUserMemberGroupIds,
    listRolesFromEntraGroups as listSaRolesFromEntra,
    finalizePlanerRoles as finalizeSaRoles,
    planerEntraGroupIds as saGroupIds,
    entraGroupsConfigured as saEntraConfigured
} from '../tools/schularbeiten-planer/schularbeiten-planer-entra-role.js';
import {
    loadEffectivePermissionsConfig as loadSaEffectivePermissions,
    normalizePermissionsConfig as normalizeSaPermissions
} from '../tools/schularbeiten-planer/schularbeiten-planer-permissions.js';
import {
    accountIsPlannerUserInList,
    direktionUsersFromTenantStammdaten
} from '../tools/freistellung-planer/freistellung-planer-direktion-users.js';
import { userIsDesignatedSchoolIt } from './designated-school-it.js';
import {
    matchTeacherByEmail,
    matchStudentByEmail
} from '../tools/schularbeiten-planer/schularbeiten-planer-state.js';
import { listRolesFromStammdaten as listSaRolesFromStammdaten } from '../tools/schularbeiten-planer/schularbeiten-planer-entra-role.js';
import {
    entraGroupsConfigured as frEntraConfigured,
    loadEffectivePermissionsConfig as loadFrEffectivePermissions,
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

function readSchoolDomainHint() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365GetSchoolDomainNoAt === 'function') {
            const d = String(window.ms365GetSchoolDomainNoAt() || '')
                .trim()
                .replace(/^@+/, '');
            if (d) return d;
        }
    } catch {
        /* ignore */
    }
    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad() || {};
            const d = String(s.domain || '')
                .trim()
                .replace(/^@+/, '');
            if (d) return d;
        }
    } catch {
        /* ignore */
    }
    const mail = accountEmail();
    const at = mail.lastIndexOf('@');
    return at > 0 ? mail.slice(at + 1) : '';
}

function loadStammdatenScope() {
    const mail = accountEmail();
    /** @type {{
     *   teacherMatch?: object|null,
     *   studentMatch?: object|null,
     *   kvMatch?: boolean,
     *   kvKlassen?: string[],
     *   teacherCount?: number,
     *   domain?: string,
     *   schulleitungMatch?: boolean
     * }} */
    const scope = {
        teacherMatch: null,
        studentMatch: null,
        kvMatch: false,
        kvKlassen: [],
        teacherCount: 0,
        domain: '',
        schulleitungMatch: false
    };
    try {
        let teachers = [];
        let classes = [];
        let students = [];
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad();
            teachers = Array.isArray(s && s.teachers) ? s.teachers : [];
            classes = Array.isArray(s && s.classes) ? s.classes : [];
            students = Array.isArray(s && s.students) ? s.students : [];
            scope.domain = String((s && s.domain) || '')
                .trim()
                .replace(/^@+/, '');
        }
        // Lehrer-Browser: App-Data / Live-SharePoint-Cache, wenn Tenant-Liste noch leer
        if (!teachers.length) {
            teachers = loadEffectiveTeachersForPersona();
            if (!classes.length) {
                const live = loadLiveSchoolStammdatenCache();
                if (Array.isArray(live.classes) && live.classes.length) classes = live.classes;
            }
        }
        if (!teachers.length && window.ms365AppDataV2 && typeof window.ms365AppDataV2.getCore === 'function') {
            try {
                const core = window.ms365AppDataV2.getCore() || {};
                if (Array.isArray(core.teachers) && core.teachers.length) teachers = core.teachers;
                if (!classes.length && Array.isArray(core.classes) && core.classes.length) {
                    classes = core.classes;
                }
            } catch {
                /* ignore */
            }
        }
        scope.teacherCount = Array.isArray(teachers) ? teachers.length : 0;
        scope.teacherMatch = matchTeacherByEmail(mail, teachers);
        scope.studentMatch = matchStudentByEmail(mail, students);
        const kvKlassen = (classes || [])
            .filter((c) => {
                const head = String(
                    (c && (c.headEmail || c.kvEmail || c.classTeacherEmail)) || ''
                )
                    .trim()
                    .toLowerCase();
                return !!mail && head === mail;
            })
            .map((c) => String((c && c.code) || '').trim())
            .filter(Boolean);
        scope.kvKlassen = kvKlassen;
        scope.kvMatch = kvKlassen.length > 0;
        try {
            scope.schulleitungMatch = accountIsPlannerUserInList(
                mail,
                direktionUsersFromTenantStammdaten()
            );
        } catch {
            scope.schulleitungMatch = false;
        }
    } catch {
        /* ignore */
    }
    if (!scope.domain) scope.domain = readSchoolDomainHint();
    return scope;
}

/**
 * @param {{
 *   teacherMatch?: object|null,
 *   teacherCount?: number,
 *   domain?: string
 * }} stammdatenScope
 * @param {boolean} groupsConfigured
 */
function buildPersonaDiagnostics(stammdatenScope, groupsConfigured) {
    const domain = String((stammdatenScope && stammdatenScope.domain) || '').trim();
    return {
        hasDomain: !!domain,
        teacherCount: Number((stammdatenScope && stammdatenScope.teacherCount) || 0) || 0,
        groupsConfigured: !!groupsConfigured,
        inTeacherList: !!(stammdatenScope && stammdatenScope.teacherMatch)
    };
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
    const config = normalizeSaPermissions(loadSaEffectivePermissions());
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
    const config = normalizeFrPermissions(loadFrEffectivePermissions());
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
    if (stammdatenScopeFr.kvMatch) add('kv', 'stammdaten-kv');
    if (stammdatenScopeFr.schulleitungMatch) add('direktion', 'stammdaten-direktion');

    if (!entra && !roles.length) return [];
    return finalizeFrRoles(roles, sources).roles;
}

/**
 * Planer-„Admin“/Direktion = erweiterte Lehrkraft-Ansicht, kein voller IT-Katalog.
 */
export function isTeacherDashboardRole(schularbeitenRoles, freistellungRoles, stammdatenScope) {
    if (schularbeitenRoles.includes('lehrer') || schularbeitenRoles.includes('admin')) return true;
    if (freistellungRoles.includes('kv') || freistellungRoles.includes('direktion')) return true;
    if (stammdatenScope && stammdatenScope.teacherMatch) return true;
    if (stammdatenScope && stammdatenScope.kvMatch) return true;
    // Schulleitung ohne Lehrerlisten-Eintrag: trotzdem Portal-Lehrkraft (nicht „Noch nicht zugeordnet“)
    if (stammdatenScope && stammdatenScope.schulleitungMatch) return true;
    return false;
}

/**
 * Nicht-IT-Persona: Entra-Lehrer gewinnt; Stammdaten/Planer-Lehrkraft schlägt
 * „nur Schüler-Gruppe“ bzw. „Gruppen konfiguriert, aber nicht Mitglied“.
 *
 * @param {{
 *   entraDashPersona?: 'lehrer'|'schueler'|null,
 *   groupsConfigured?: boolean,
 *   isTeacher?: boolean,
 *   explicitPlanerSchueler?: boolean,
 *   studentMatch?: boolean
 * }} opts
 * @returns {{ isLehrer: boolean, isSchueler: boolean, unknownUser?: boolean, personaSource: string }}
 */
export function resolveNonItDashboardPersona(opts) {
    const entra = opts && opts.entraDashPersona != null ? opts.entraDashPersona : null;
    const groupsConfigured = !!(opts && opts.groupsConfigured);
    const isTeacher = !!(opts && opts.isTeacher);
    const explicitPlanerSchueler = !!(opts && opts.explicitPlanerSchueler);
    const studentMatch = !!(opts && opts.studentMatch);

    if (entra === 'lehrer') {
        return {
            isLehrer: true,
            isSchueler: false,
            personaSource: 'dashboard-entra-lehrer'
        };
    }

    // Lehrkraft per Stammdaten / Planer-Rollen hat Vorrang vor Schüler-Fallback
    if (isTeacher) {
        let personaSource = 'stammdaten-planer';
        if (entra === 'schueler') personaSource = 'stammdaten-lehrer-over-entra-schueler';
        else if (groupsConfigured && !entra) personaSource = 'stammdaten-lehrer-entra-miss';
        return {
            isLehrer: true,
            isSchueler: false,
            personaSource
        };
    }

    if (entra === 'schueler') {
        return {
            isLehrer: false,
            isSchueler: true,
            personaSource: 'dashboard-entra-schueler'
        };
    }

    // Explizit Schüler (Planer/Stammdaten) → Schüler-Portal.
    // Sonst: noch nicht zugeordnet (z. B. vor Verzeichnis-Sync) – nicht als Schüler/in begrüßen.
    if (explicitPlanerSchueler || studentMatch) {
        return {
            isLehrer: false,
            isSchueler: true,
            personaSource: groupsConfigured && !entra ? 'dashboard-entra-none' : 'default-schueler'
        };
    }

    return {
        isLehrer: false,
        isSchueler: false,
        unknownUser: true,
        personaSource: groupsConfigured && !entra ? 'dashboard-entra-none' : 'pending-directory'
    };
}

/**
 * @typedef {{
 *   filtering: boolean,
 *   loggedIn: boolean,
 *   unknownUser?: boolean,
 *   isIt: boolean,
 *   isLehrer: boolean,
 *   isSchueler: boolean,
 *   personaSource?: string,
 *   personaDiagnostics?: {
 *     hasDomain?: boolean,
 *     teacherCount?: number,
 *     groupsConfigured?: boolean,
 *     inTeacherList?: boolean
 *   },
 *   globalAdmin?: boolean,
 *   operator?: boolean,
 *   schularbeitenRoles: string[],
 *   freistellungRoles: string[]
 * }} DashboardPersonas
 */

/**
 * Nur für Homepage-Screenshots (`?landingShot=1&persona=it|lehrer|schueler`).
 * Echte Anmeldung und Stammdaten bleiben unberührt, solange der Parameter fehlt.
 *
 * @param {string} [search]
 * @returns {DashboardPersonas|null}
 */
export function landingShotDashboardPersonas(search) {
    let raw = search;
    if (raw == null) {
        try {
            raw = typeof location !== 'undefined' ? location.search || '' : '';
        } catch {
            raw = '';
        }
    }
    let who = '';
    try {
        const q = new URLSearchParams(String(raw || '').replace(/^\?/, ''));
        if (q.get('landingShot') !== '1') return null;
        who = String(q.get('persona') || '')
            .trim()
            .toLowerCase();
    } catch {
        return null;
    }
    if (who !== 'it' && who !== 'lehrer' && who !== 'schueler') return null;

    const shellValue = who === 'it' ? 'full' : who === 'schueler' ? 'schueler' : 'portal';
    try {
        localStorage.setItem('ms365-dashboard-it-preview-v1', shellValue);
        localStorage.setItem('ms365-dashboard-it-preview-v1:demo-tenant', shellValue);
    } catch {
        /* ignore */
    }

    if (who === 'it') {
        return {
            filtering: false,
            loggedIn: true,
            isIt: true,
            isLehrer: false,
            isSchueler: false,
            globalAdmin: true,
            personaSource: 'landing-shot',
            schularbeitenRoles: ['admin'],
            freistellungRoles: ['direktion']
        };
    }

    return {
        filtering: true,
        loggedIn: true,
        isIt: false,
        isLehrer: who === 'lehrer',
        isSchueler: who === 'schueler',
        personaSource: 'landing-shot',
        schularbeitenRoles: who === 'schueler' ? ['schueler'] : ['lehrer'],
        freistellungRoles: who === 'schueler' ? ['schueler'] : ['lehrer']
    };
}

/**
 * @returns {Promise<DashboardPersonas>}
 */
export async function resolveDashboardPersonas() {
    const landingShot = landingShotDashboardPersonas();
    if (landingShot) return landingShot;

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

    const [schularbeitenRoles, freistellungRoles, globalAdmin, operator, designatedSchoolIt] = await Promise.all([
        resolveSchularbeitenRoles(),
        resolveFreistellungRoles(),
        userIsEntraGlobalAdministrator(),
        resolveOperatorFlag(),
        userIsDesignatedSchoolIt()
    ]);

    const stammdatenScope = loadStammdatenScope();

    /** Global Admin, Lizenz-IT-Kontakt oder Plattform-Betreiber: alles sichtbar */
    const fullAccess = globalAdmin || operator || designatedSchoolIt;

    if (fullAccess) {
        return {
            filtering: false,
            loggedIn: true,
            isIt: true,
            isLehrer: false,
            isSchueler: false,
            globalAdmin: !!globalAdmin,
            operator: !!operator,
            designatedSchoolIt: !!designatedSchoolIt,
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

    const isTeacher = isTeacherDashboardRole(schularbeitenRoles, freistellungRoles, stammdatenScope);
    const explicitPlanerSchueler =
        schularbeitenRoles.includes('schueler') || freistellungRoles.includes('schueler');
    const decided = resolveNonItDashboardPersona({
        entraDashPersona,
        groupsConfigured,
        isTeacher,
        explicitPlanerSchueler,
        studentMatch: !!(stammdatenScope && stammdatenScope.studentMatch)
    });

    return {
        filtering: true,
        loggedIn: true,
        unknownUser: !!decided.unknownUser,
        isIt: false,
        isLehrer: decided.isLehrer,
        isSchueler: decided.isSchueler,
        personaSource: decided.personaSource,
        personaDiagnostics: buildPersonaDiagnostics(stammdatenScope, groupsConfigured),
        globalAdmin: !!globalAdmin,
        operator: !!operator,
        designatedSchoolIt: !!designatedSchoolIt,
        schularbeitenRoles,
        freistellungRoles
    };
}
