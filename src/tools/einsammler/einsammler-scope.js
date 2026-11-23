/**
 * Sichtbarkeit Einsammler: Lehrkraft nur eigene Vorgänge, Admin optional alle.
 */

import { loadEinsammlerSettings } from './einsammler-settings.js';

/**
 * @returns {{ email: string, name: string, loggedIn: boolean }}
 */
export function getEinsammlerAccount() {
    try {
        if (typeof window !== 'undefined') {
            const q = new URLSearchParams(window.location.search || '');
            if (q.get('landingShot') === '1') {
                return {
                    email: 'mueller@schule.demo',
                    name: 'Maria Müller',
                    loggedIn: true
                };
            }
        }
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            if (info) {
                const email = String(info.username || info.upn || '')
                    .trim()
                    .toLowerCase();
                return {
                    email,
                    name: String(info.name || '').trim(),
                    loggedIn: !!email
                };
            }
        }
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetUserPrincipalName === 'function') {
            const email = String(window.ms365AuthGetUserPrincipalName() || '')
                .trim()
                .toLowerCase();
            return { email, name: '', loggedIn: !!email };
        }
    } catch {
        /* ignore */
    }
    return { email: '', name: '', loggedIn: false };
}

/**
 * @param {string} text
 * @returns {string[]}
 */
export function parseAdminMails(text) {
    return String(text || '')
        .split(/[\s,;]+/)
        .map((s) => s.trim().toLowerCase())
        .filter((s) => s.includes('@'));
}

/** @type {{ email: string, isAdmin: boolean, sources: string[], demo: boolean, resolved: boolean }} */
let adminAccessCache = {
    email: '',
    isAdmin: false,
    sources: [],
    demo: false,
    resolved: false
};

/**
 * @param {{ email?: string, isAdmin: boolean, sources?: string[], demo?: boolean }} result
 */
export function setEinsammlerAdminAccessCache(result) {
    adminAccessCache = {
        email: String((result && result.email) || '')
            .trim()
            .toLowerCase(),
        isAdmin: !!(result && result.isAdmin),
        sources: Array.isArray(result && result.sources) ? result.sources.slice() : [],
        demo: !!(result && result.demo),
        resolved: true
    };
}

export function getEinsammlerAdminAccessCache() {
    return {
        email: adminAccessCache.email,
        isAdmin: adminAccessCache.isAdmin,
        sources: adminAccessCache.sources.slice(),
        demo: adminAccessCache.demo,
        resolved: adminAccessCache.resolved
    };
}

/**
 * Sync-Admin-Check: nach Entra-Resolve Cache, sonst Legacy bis Resolve fertig.
 * @param {string} [email]
 */
export function isEinsammlerAdmin(email) {
    const em = String(email || getEinsammlerAccount().email || '')
        .trim()
        .toLowerCase();
    if (!em) return false;

    if (adminAccessCache.resolved && adminAccessCache.email === em) {
        return adminAccessCache.isAdmin;
    }

    const settings = loadEinsammlerSettings();
    if (parseAdminMails(settings.adminEmails).includes(em)) return true;

    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const core = window.ms365TenantSettingsLoad();
            const data = (core && core.data) || core || {};
            const admin = Array.isArray(data.admin) ? data.admin : [];
            for (const a of admin) {
                if (String((a && a.email) || '').trim().toLowerCase() === em) return true;
            }
            const administration = Array.isArray(data.administration) ? data.administration : [];
            for (const a of administration) {
                const role = String((a && (a.role || a.group || a.code)) || '').toLowerCase();
                const mail = String((a && a.email) || '').trim().toLowerCase();
                if (mail !== em) continue;
                if (
                    !role ||
                    role.includes('schulleitung') ||
                    role.includes('direktion') ||
                    role.includes('admin') ||
                    role.includes('it') ||
                    role.includes('verwaltung') ||
                    role.includes('sekretariat')
                ) {
                    return true;
                }
            }
        }
    } catch {
        /* ignore */
    }

    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365OperatorAccess &&
            typeof window.ms365OperatorAccess.isOperatorCached === 'function' &&
            window.ms365OperatorAccess.isOperatorCached()
        ) {
            return true;
        }
    } catch {
        /* ignore */
    }

    return false;
}

/**
 * @param {object} head
 * @param {string} email
 */
export function isOwnedBy(head, email) {
    const owner = String((head && head.erstellerEmail) || '')
        .trim()
        .toLowerCase();
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return false;
    if (!owner) return false;
    return owner === em;
}

/**
 * @param {object[]} list
 * @param {{ email: string, admin: boolean, showAll: boolean }} opts
 */
export function filterEinsammlerForViewer(list, opts) {
    const rows = Array.isArray(list) ? list : [];
    const email = String((opts && opts.email) || '')
        .trim()
        .toLowerCase();
    const showAll = !!(opts && opts.admin && opts.showAll);
    if (showAll) return rows.slice();
    if (!email) return [];
    return rows.filter((h) => isOwnedBy(h, email));
}

/**
 * Antworten nur zu sichtbaren Einsammlern.
 * @param {object[]} antworten
 * @param {object[]} visibleHeads
 */
export function filterAntwortenForHeads(antworten, visibleHeads) {
    const ids = new Set((visibleHeads || []).map((h) => h.id));
    return (antworten || []).filter((a) => a && ids.has(a.einsammlerId));
}
