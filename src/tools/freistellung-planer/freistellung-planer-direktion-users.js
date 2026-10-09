/**
 * Einzelpersonen mit Planer-Rolle (zusätzlich zur Entra-Gruppe).
 */
import {
    ADMIN_TIER_SCHULLEITUNG,
    inferAdminTierForPersonRow,
    inferAdminTierForRole
} from '../../shared/administration-audience-logic.js';

/**
 * @typedef {{ id?: string, displayName?: string, mail?: string, email?: string, name?: string }} PlannerUser
 */

/** @alias normalizePlannerUsers */
export function normalizeDirektionUsers(raw) {
    return normalizePlannerUsers(raw);
}

/**
 * @param {PlannerUser[]|null|undefined} raw
 * @returns {Array<{ id: string, displayName: string, mail: string }>}
 */
export function normalizePlannerUsers(raw) {
    const arr = Array.isArray(raw) ? raw : [];
    const out = [];
    const seen = new Set();
    arr.forEach((u) => {
        if (!u || typeof u !== 'object') return;
        const mail = String(u.mail || u.email || '').trim().toLowerCase();
        const id = String(u.id || '').trim();
        const displayName = String(u.displayName || u.name || mail || '').trim();
        const key = (id || mail).toLowerCase();
        if (!key || seen.has(key)) return;
        if (!mail && !id) return;
        seen.add(key);
        out.push({ id, displayName, mail });
    });
    return out;
}

/**
 * @param {DirektionUser[]|null|undefined} list
 * @param {DirektionUser} entry
 */
export function mergeDirektionUser(list, entry) {
    return mergePlannerUser(list, entry);
}

export function mergePlannerUser(list, entry) {
    return normalizePlannerUsers([...(list || []), entry]);
}

/**
 * @param {string} accountEmail
 * @param {PlannerUser[]|null|undefined} users
 */
export function accountIsPlannerUserInList(accountEmail, users) {
    const mail = String(accountEmail || '').trim().toLowerCase();
    if (!mail) return false;
    return normalizePlannerUsers(users).some((u) => u.mail && u.mail === mail);
}

/** @param {string} accountEmail @param {PlannerUser[]|null|undefined} users */
export function accountIsDirektionPlannerUser(accountEmail, users) {
    return accountIsPlannerUserInList(accountEmail, users);
}

/**
 * Schulleitung (z. B. Direktion) aus Tenant-Stammdaten – nicht das gesamte Verwaltungs-Personal.
 */
export function direktionUsersFromTenantStammdaten() {
    /** @type {Map<string, { displayName: string, mail: string }>} */
    const map = new Map();
    const add = (name, email) => {
        const mail = String(email || '').trim().toLowerCase();
        if (!mail || mail.indexOf('@') < 0) return;
        if (!map.has(mail)) {
            map.set(mail, { displayName: String(name || mail).trim(), mail });
        }
    };
    try {
        const core =
            typeof window !== 'undefined' && window.ms365TenantSettingsLoad
                ? window.ms365TenantSettingsLoad()
                : null;
        const data = (core && core.data) || core || {};
        const roles = Array.isArray(data.adminRoles) ? data.adminRoles : [];
        (data.admin || []).forEach((p) => {
            if (inferAdminTierForPersonRow(p, roles) !== ADMIN_TIER_SCHULLEITUNG) return;
            add(p.name, p.email);
        });
        (data.administration || []).forEach((entry) => {
            if (!entry) return;
            if (Array.isArray(entry.people)) {
                if (inferAdminTierForRole(entry) !== ADMIN_TIER_SCHULLEITUNG) return;
                entry.people.forEach((person) => add(person && person.name, person && person.email));
                return;
            }
            if (entry.kind !== 'person') return;
            if (inferAdminTierForPersonRow(entry, roles) !== ADMIN_TIER_SCHULLEITUNG) return;
            add(entry.name, entry.email);
        });
    } catch {
        /* ignore */
    }
    return Array.from(map.values()).map((u) => ({ id: '', displayName: u.displayName, mail: u.mail }));
}
