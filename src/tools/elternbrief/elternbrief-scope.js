/**
 * Persönliche Sicht Elternbrief (Lehrkraft nur eigene).
 */

import { loadElternbriefSettings } from './elternbrief-settings.js';

/**
 * @returns {{ email: string, name: string, loggedIn: boolean }}
 */
export function getElternbriefAccount() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            if (info) {
                const email = String(info.username || info.upn || '')
                    .trim()
                    .toLowerCase();
                return { email, name: String(info.name || '').trim(), loggedIn: !!email };
            }
        }
    } catch {
        /* ignore */
    }
    return { email: '', name: '', loggedIn: false };
}

/**
 * @param {string} text
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
export function setElternbriefAdminAccessCache(result) {
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

export function getElternbriefAdminAccessCache() {
    return {
        email: adminAccessCache.email,
        isAdmin: adminAccessCache.isAdmin,
        sources: adminAccessCache.sources.slice(),
        demo: adminAccessCache.demo,
        resolved: adminAccessCache.resolved
    };
}

/**
 * @param {string} [email]
 */
export function isElternbriefAdmin(email) {
    const em = String(email || getElternbriefAccount().email || '')
        .trim()
        .toLowerCase();
    if (!em) return false;
    if (adminAccessCache.resolved && adminAccessCache.email === em) {
        return adminAccessCache.isAdmin;
    }
    if (parseAdminMails(loadElternbriefSettings().adminEmails).includes(em)) return true;
    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const core = window.ms365TenantSettingsLoad();
            const data = (core && core.data) || core || {};
            for (const a of Array.isArray(data.admin) ? data.admin : []) {
                if (String((a && a.email) || '').trim().toLowerCase() === em) return true;
            }
            for (const a of Array.isArray(data.administration) ? data.administration : []) {
                const mail = String((a && a.email) || '').trim().toLowerCase();
                const role = String((a && (a.role || a.group)) || '').toLowerCase();
                if (mail !== em) continue;
                if (
                    !role ||
                    role.includes('schulleitung') ||
                    role.includes('direktion') ||
                    role.includes('admin') ||
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
    return false;
}

/**
 * @param {object} brief
 * @param {string} email
 */
export function isOwnedBy(brief, email) {
    const owner = String((brief && brief.erstellerEmail) || '')
        .trim()
        .toLowerCase();
    const em = String(email || '')
        .trim()
        .toLowerCase();
    return !!em && !!owner && owner === em;
}

/**
 * @param {object[]} list
 * @param {{ email: string, admin: boolean, showAll: boolean }} opts
 */
export function filterBriefeForViewer(list, opts) {
    const rows = Array.isArray(list) ? list : [];
    if (opts && opts.admin && opts.showAll) return rows.slice();
    const email = String((opts && opts.email) || '')
        .trim()
        .toLowerCase();
    if (!email) return [];
    return rows.filter((b) => isOwnedBy(b, email));
}

/**
 * @param {object[]} empfaenger
 * @param {object[]} visibleBriefe
 */
export function filterEmpfaengerForBriefe(empfaenger, visibleBriefe) {
    const ids = new Set((visibleBriefe || []).map((b) => b.id));
    return (empfaenger || []).filter((e) => e && ids.has(e.briefId));
}
