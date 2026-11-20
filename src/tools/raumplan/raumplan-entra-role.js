/**
 * Raumplan: Schreibrechte nur Betrieb-Admin.
 */
import {
    createBetriebAdminCache,
    getBetriebToolAccount,
    resolveBetriebAdminAccess
} from '../../shared/betrieb-admin-entra.js';
import { announceStaffAccess } from '../../shared/app-role-chrome.js';

const cache = createBetriebAdminCache();

export function getRaumplanAccount() {
    return getBetriebToolAccount();
}

/**
 * @param {string} [email]
 */
export async function resolveRaumplanAdmin(email) {
    const em = String(email || getRaumplanAccount().email || '')
        .trim()
        .toLowerCase();
    const result = await resolveBetriebAdminAccess(em);
    cache.set({
        email: em,
        isAdmin: !!result.isAdmin,
        sources: result.sources || [],
        demo: !!result.demo
    });
    announceStaffAccess('raumplan', {
        email: em,
        isAdmin: !!result.isAdmin,
        sources: result.sources || []
    });
    return result;
}

export function getRaumplanAdminAccessCache() {
    return cache.get();
}

/**
 * @param {string} [email]
 */
export function isRaumplanAdmin(email) {
    return cache.isAdmin(email || getRaumplanAccount().email);
}
