/**
 * Betrieb-Apps (Aufgaben, Technik-Meldungen, Raumplan, Ausgaben):
 * Admin = Schulleitung / Verwaltung / SA-Admin / Schul-IT (wie Ablauf).
 * Docs: rollen-vertrag.md Welle 4.
 */
import { resolveAblaufAdminAccess } from './ablauf-admin-entra.js';

/**
 * @returns {{ email: string, name: string, loggedIn: boolean }}
 */
export function getBetriebToolAccount() {
    try {
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
        if (
            typeof window !== 'undefined' &&
            typeof window.ms365AuthGetUserPrincipalName === 'function'
        ) {
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
 * Sync-Admin-Cache für UI-Render vor/nach async Resolve.
 */
export function createBetriebAdminCache() {
    /** @type {{ email: string, isAdmin: boolean, sources: string[], demo: boolean, resolved: boolean }} */
    let cache = {
        email: '',
        isAdmin: false,
        sources: [],
        demo: false,
        resolved: false
    };

    return {
        /**
         * @param {{ email?: string, isAdmin: boolean, sources?: string[], demo?: boolean }} result
         */
        set(result) {
            cache = {
                email: String((result && result.email) || '')
                    .trim()
                    .toLowerCase(),
                isAdmin: !!(result && result.isAdmin),
                sources: Array.isArray(result && result.sources) ? result.sources.slice() : [],
                demo: !!(result && result.demo),
                resolved: true
            };
        },
        get() {
            return {
                email: cache.email,
                isAdmin: cache.isAdmin,
                sources: cache.sources.slice(),
                demo: cache.demo,
                resolved: cache.resolved
            };
        },
        /**
         * @param {string} [email]
         */
        isAdmin(email) {
            const em = String(email || '')
                .trim()
                .toLowerCase();
            if (cache.resolved && (!em || cache.email === em)) return cache.isAdmin;
            return false;
        }
    };
}

/**
 * @param {string} [email]
 * @param {{ legacyAdminEmails?: string[], includeIt?: boolean, includeVerwaltung?: boolean }} [opts]
 */
export async function resolveBetriebAdminAccess(email, opts) {
    return resolveAblaufAdminAccess(email, {
        includeIt: true,
        includeVerwaltung: true,
        ...(opts || {})
    });
}
