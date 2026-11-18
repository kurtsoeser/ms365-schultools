/**
 * Ausgaben: Admin = Verwaltung/Schulleitung/IT; Lehrkraft nur „meine“ Vorgänge.
 */
import {
    createBetriebAdminCache,
    getBetriebToolAccount,
    resolveBetriebAdminAccess
} from '../../shared/betrieb-admin-entra.js';
import { loadAusgabenSettings } from './ausgaben-settings.js';

const cache = createBetriebAdminCache();

/**
 * @param {string} text
 * @returns {string[]}
 */
export function parseAusgabenAdminMails(text) {
    return String(text || '')
        .split(/[\s,;]+/)
        .map((s) => s.trim().toLowerCase())
        .filter((s) => s.includes('@'));
}

export function getAusgabenAccount() {
    return getBetriebToolAccount();
}

/**
 * @param {string} [email]
 */
export async function resolveAusgabenAdmin(email) {
    const em = String(email || getAusgabenAccount().email || '')
        .trim()
        .toLowerCase();
    const settings = loadAusgabenSettings();
    const result = await resolveBetriebAdminAccess(em, {
        legacyAdminEmails: parseAusgabenAdminMails(settings.adminEmails)
    });
    cache.set({
        email: em,
        isAdmin: !!result.isAdmin,
        sources: result.sources || [],
        demo: !!result.demo
    });
    return result;
}

export function getAusgabenAdminAccessCache() {
    return cache.get();
}

/**
 * @param {string} [email]
 */
export function isAusgabenAdmin(email) {
    const em = String(email || getAusgabenAccount().email || '')
        .trim()
        .toLowerCase();
    if (cache.isAdmin(em)) return true;
    if (!em) return false;
    try {
        return parseAusgabenAdminMails(loadAusgabenSettings().adminEmails).includes(em);
    } catch {
        return false;
    }
}
