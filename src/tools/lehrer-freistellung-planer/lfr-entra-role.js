/**
 * Lehrer-Freistellung: Rollen aus Entra-Gruppen / Setup-Mails / Stammdaten.
 * Demo-Umschalter nur mit ?demoRole=1 (docs/rollen-vertrag.md).
 */
import {
    canUseAppRoleSwitcher,
    isDemoRoleUiEnabled,
    readDemoRoleOverrideFromUrl
} from '../../shared/app-role-switcher-policy.js';
import { readItShellChoice } from '../../shared/dashboard-layout-mode.js';
import { loadSchoolAudienceGroups } from '../../shared/school-audience-groups.js';
import {
    fetchUserDirectMemberGroupIds,
    fetchUserMemberGroupIds,
    userIsEntraGlobalAdministrator
} from '../schularbeiten-planer/schularbeiten-planer-entra-role.js';
import {
    accountIsPlannerUserInList,
    direktionUsersFromTenantStammdaten
} from '../freistellung-planer/freistellung-planer-direktion-users.js';
import {
    accountIsLfrDirektionAccess,
    accountIsLfrLehrerAccess,
    loadLfrPermissions,
    resolveLfrApproverEmail
} from './lfr-permissions.js';
import { matchTeacherByEmail, persistRole, resolveRole } from './lfr-state.js';
import { MS365_DEMO_DIREKTOR_EMAIL } from '../../shared/demo-portal-teacher.js';

/** @typedef {'direktion'|'lehrer'} LfrRole */

export const LFR_ROLE_ORDER = ['direktion', 'lehrer'];

export { readDemoRoleOverrideFromUrl };

/**
 * @param {ReturnType<typeof loadLfrPermissions>} [perms]
 */
export function entraGroupsConfigured(perms) {
    const p = perms || loadLfrPermissions();
    if (
        p.groupLehrerId ||
        p.groupDirektionId ||
        (p.lehrerUsers && p.lehrerUsers.length) ||
        (p.direktionUsers && p.direktionUsers.length) ||
        resolveLfrApproverEmail(p)
    ) {
        return true;
    }
    try {
        const stamm = loadSchoolAudienceGroups();
        if (stamm && stamm.groupLehrerId) return true;
    } catch {
        /* ignore */
    }
    return false;
}

/**
 * @param {LfrRole[]|string[]} roles
 * @returns {LfrRole[]}
 */
export function sortLfrRoles(roles) {
    const set = new Set((roles || []).map((r) => String(r)));
    return LFR_ROLE_ORDER.filter((r) => set.has(r));
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
 * Schul-IT: Portal → Lehrer (Alltag), IT-Backend → Direktion.
 * Direktion / Genehmiger (nicht Schul-IT): immer Direktion – sonst keine Freigabe-Inbox.
 * @param {LfrRole[]|string[]} available
 * @returns {LfrRole|null}
 */
export function shellPreferredLfrRole(available) {
    const set = new Set(sortLfrRoles(available || []));
    if (!set.size) return null;
    let shell = 'portal';
    try {
        shell = readItShellChoice();
    } catch {
        shell = 'portal';
    }
    if (shell === 'full' && set.has('direktion')) return 'direktion';
    // Portal-Vorschau nur für Schul-IT – Direktion darf nicht auf Lehrer-Alltag fallen
    if (shell !== 'full' && currentPortalIsSchulIt() && set.has('lehrer')) return 'lehrer';
    if (set.has('direktion')) return 'direktion';
    if (set.has('lehrer')) return 'lehrer';
    return null;
}

/**
 * @param {LfrRole[]} available
 * @param {{ preferredActiveRole?: LfrRole|null, preferredDemoRole?: LfrRole|null, preferShellDefault?: boolean }} [opts]
 */
export function resolveActiveLfrRole(available, opts) {
    const list = sortLfrRoles(available || []);
    const set = new Set(list);
    const cur = opts && opts.preferredActiveRole;
    const demo = opts && opts.preferredDemoRole;
    if (cur && set.has(cur)) return cur;
    if (demo && set.has(demo)) return demo;
    if (opts && opts.preferShellDefault) {
        const shellRole = shellPreferredLfrRole(list);
        if (shellRole && set.has(shellRole)) return shellRole;
    }
    const stored = resolveRole();
    if (stored && set.has(stored)) return stored;
    const softShell = shellPreferredLfrRole(list);
    if (softShell && set.has(softShell)) return softShell;
    return list[0] || 'lehrer';
}

/**
 * @param {boolean} [_entraConfigured]
 * @param {boolean} [demoRoleOverride]
 */
export function isLfrDemoRoleUiEnabled(_entraConfigured, demoRoleOverride) {
    return isDemoRoleUiEnabled(demoRoleOverride);
}

/**
 * @param {LfrRole[]} roles
 * @param {boolean} demoRoleUi
 */
export function canUseLfrRoleSwitcher(roles, demoRoleUi) {
    // Nie Direktion/Lehrkraft-Chips in der produktiven Lehrerwelt.
    // Nur explizite IT-Vorschau mit ?demoRole=1 (wie Schulaktivitäten).
    if (!demoRoleUi) return false;
    return canUseAppRoleSwitcher(roles, { demoRoleUi: true });
}

export function roleSourceLabel(source) {
    const s = String(source || '');
    if (s === 'entra') return 'Entra-Gruppe';
    if (s === 'global-admin') return 'Globaler Administrator';
    if (s === 'demo-direktor') return 'Demo-Direktor';
    if (s === 'setup-direktion') return 'Genehmiger (Setup)';
    if (s === 'stammdaten-direktion') return 'Schulleitung (Stammdaten)';
    if (s === 'direktion-user') return 'Direktion (Freistellung)';
    if (s === 'stammdaten') return 'Stammdaten';
    if (s === 'lehrer-user') return 'Lehrkraft (Einzelperson)';
    if (s === 'demo') return 'IT-Vorschau';
    if (s === 'konto-offen') return 'Konto wird gelesen';
    if (s === 'default') return 'Standard, ohne Zuordnung';
    return s || '–';
}

/**
 * Nur LFR-eigene + Stammdaten-Lehrkräfte-Gruppe.
 * Schüler-Freistellung-Direktion darf hier keine LFR-Direktion auslösen
 * (sonst sehen reine Lehrkräfte Freigabe/System).
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
 * Entra-Direktion nur behalten bei direkter Mitgliedschaft oder starkem Signal
 * (Genehmiger / Einzelperson / Schulleitung / Global Admin).
 * checkMemberGroups ist transitiv – verschachtelte Lehrergruppe unter Direktion
 * würde sonst alle Lehrkräfte zu Direktion machen.
 * @param {{
 *   roles: LfrRole[],
 *   sources: Record<string, string>,
 *   hasDirectDirektion: boolean,
 *   hasStrongDirektionSignal: boolean
 * }} opts
 */
export function refineLfrDirektionFromEntra(opts) {
    const roles = Array.isArray(opts && opts.roles) ? opts.roles.slice() : [];
    const sources = Object.assign({}, (opts && opts.sources) || {});
    if (!roles.includes('direktion')) {
        return { roles: sortLfrRoles(roles), sources };
    }
    if (opts && opts.hasStrongDirektionSignal) {
        return { roles: sortLfrRoles(roles), sources };
    }
    if (sources.direktion === 'entra' && !(opts && opts.hasDirectDirektion)) {
        const next = roles.filter((r) => r !== 'direktion');
        delete sources.direktion;
        return { roles: sortLfrRoles(next), sources };
    }
    return { roles: sortLfrRoles(roles), sources };
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
    const isLehrerGroup =
        (lehrerId && member.has(lehrerId)) || (stammLehrer && member.has(stammLehrer));
    // Gleiche GUID für Direktion + Lehrkräfte → nur Lehrkraft (kein Admin-Nav für alle Lehrer)
    const direktionIsDistinctLehrerGroup =
        !!dirId && dirId !== lehrerId && dirId !== stammLehrer;
    if (direktionIsDistinctLehrerGroup && member.has(dirId)) {
        roles.push('direktion');
    }
    if (isLehrerGroup) {
        roles.push('lehrer');
    }
    return sortLfrRoles(roles);
}

/**
 * @param {object} state
 * @param {{ demoRoleOverride?: boolean, preferredDemoRole?: LfrRole|null, preferredActiveRole?: LfrRole|null, preferShellDefault?: boolean }} [opts]
 */
export async function applyLfrRoleFromEntra(state, opts) {
    const perms = loadLfrPermissions();
    const entra = entraGroupsConfigured(perms);
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
        const roles = sortLfrRoles(['direktion', 'lehrer']);
        state.planerRoles = roles;
        state.planerRoleSources = { direktion: 'demo', lehrer: 'demo' };
        state.planerAccessDenied = false;
        state.role = resolveActiveLfrRole(roles, pickOpts);
        persistRole(state.role);
        state.roleHint = 'IT-Vorschau (?demoRole=1).';
        return;
    }

    const emEarly = String(state.accountEmail || '')
        .trim()
        .toLowerCase();
    // Ohne Konto keine Lehrkraft-Entscheidung speichern. Sonst gewinnt ein
    // früher Lauf (Portal noch nicht angemeldet) und Direktion bleibt unsichtbar.
    if (!emEarly) {
        state.planerRoles = ['lehrer'];
        state.planerRoleSources = { lehrer: 'konto-offen' };
        state.planerAccessDenied = false;
        state.role = 'lehrer';
        state.roleHint = 'Konto wird gelesen …';
        return;
    }

    /** @type {LfrRole[]} */
    let roles = [];
    /** @type {Record<string, string>} */
    let sources = {};

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
    const dirId = String(perms.groupDirektionId || '')
        .trim()
        .toLowerCase();
    let hasDirectDirektion = false;
    if (ids.length) {
        try {
            const member = await fetchUserMemberGroupIds(ids);
            const fromEntra = listRolesFromEntraGroups(member, perms);
            fromEntra.forEach((r) => {
                if (!roles.includes(r)) roles.push(r);
                if (!sources[r]) sources[r] = 'entra';
            });
            if (dirId && fromEntra.includes('direktion')) {
                try {
                    const direct = await fetchUserDirectMemberGroupIds([dirId]);
                    hasDirectDirektion = direct.has(dirId);
                } catch {
                    hasDirectDirektion = false;
                }
            }
        } catch {
            /* ignore */
        }
    }

    const em = String(state.accountEmail || '')
        .trim()
        .toLowerCase();
    // Nur LFR-Setup/Perms – nicht Schüler-Freistellung-Direktion-Liste
    let hasStrongDirektionSignal = sources.direktion === 'global-admin';
    if (lfrAccountIsDirektion(em, perms)) {
        if (!roles.includes('direktion')) roles.push('direktion');
        sources.direktion = em === MS365_DEMO_DIREKTOR_EMAIL ? 'demo-direktor' : 'setup-direktion';
        hasStrongDirektionSignal = true;
    }
    // Schulleitung aus Tenant-Stammdaten → Direktion, unabhängig von Lehrerliste / LFR-Perms
    try {
        const direktionStamm = direktionUsersFromTenantStammdaten();
        if (accountIsPlannerUserInList(em, direktionStamm)) {
            if (!roles.includes('direktion')) roles.push('direktion');
            sources.direktion = 'stammdaten-direktion';
            hasStrongDirektionSignal = true;
        }
    } catch {
        /* ignore */
    }

    // Transitive Entra-Direktion (z. B. Lehrergruppe unter Direktion) verwerfen
    ({ roles, sources } = refineLfrDirektionFromEntra({
        roles,
        sources,
        hasDirectDirektion,
        hasStrongDirektionSignal
    }));

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

    // Direktion sieht auch Lehrkraft-Views (Antrag stellen / Meine Anträge)
    if (roles.includes('direktion')) {
        if (!roles.includes('lehrer')) roles.push('lehrer');
        if (!sources.lehrer) sources.lehrer = 'direktion';
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
            state.planerRoleSources = sources;
            state.planerAccessDenied = true;
            state.role = 'lehrer';
            state.roleHint =
                'Keine Berechtigung für Lehrer-Freistellungen. IT: Entra-Gruppen im Setup zuordnen. Vorschau: ?demoRole=1';
            persistRole(state.role);
            return;
        }
    } else {
        state.planerAccessDenied = false;
        state.roleHint = '';
    }

    state.planerRoles = roles;
    state.planerRoleSources = sources;
    state.role = enforceLfrDirektionRole(
        resolveActiveLfrRole(roles, pickOpts),
        roles,
        em,
        perms
    );
    persistRole(state.role);
}

/**
 * Genehmiger dieser App bleibt Direktion, auch wenn die Portal-Rolle „Lehrkraft“ ist
 * oder zuletzt eine Lehrer-Ansicht gespeichert wurde.
 * @param {LfrRole|string} resolved
 * @param {LfrRole[]} roles
 * @param {string} accountEmail
 * @param {ReturnType<typeof loadLfrPermissions>} [perms]
 * @returns {LfrRole}
 */
export function enforceLfrDirektionRole(resolved, roles, accountEmail, perms) {
    if ((roles || []).includes('direktion') && lfrAccountIsDirektion(accountEmail, perms)) {
        return 'direktion';
    }
    return resolved === 'direktion' ? 'direktion' : 'lehrer';
}

/**
 * Diese App: Genehmiger, Direktion-Person, Schulleitung oder Demo-Direktor.
 * Unabhängig von der Portal-Rolle (die bleibt oft „Lehrkraft“).
 * @param {string} accountEmail
 * @param {ReturnType<typeof loadLfrPermissions>} [perms]
 * @param {boolean} [schulleitungMatch]
 */
export function lfrAccountIsDirektion(accountEmail, perms, schulleitungMatch) {
    const em = String(accountEmail || '')
        .trim()
        .toLowerCase();
    if (!em) return false;
    if (em === MS365_DEMO_DIREKTOR_EMAIL) return true;
    if (accountIsLfrDirektionAccess(em, perms)) return true;
    return !!schulleitungMatch;
}
