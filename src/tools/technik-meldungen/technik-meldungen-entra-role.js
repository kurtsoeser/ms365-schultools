/**
 * Technik-Meldungen: Admin = IT/Schulleitung/Verwaltung (Betrieb-Helper).
 * Demo-Vorschau nur mit ?demoRole=1.
 */
import {
    createBetriebAdminCache,
    getBetriebToolAccount,
    resolveBetriebAdminAccess
} from '../../shared/betrieb-admin-entra.js';
import {
    canUseAppRoleSwitcher,
    readDemoRoleOverrideFromUrl
} from '../../shared/app-role-switcher-policy.js';
import { announceStaffAccess } from '../../shared/app-role-chrome.js';

const cache = createBetriebAdminCache();

/** @type {'auto'|'melder'|'admin'|null} */
let demoRoleOverride = null;

export function getTechnikAccount() {
    return getBetriebToolAccount();
}

export function isTechnikDemoRoleUiEnabled() {
    return readDemoRoleOverrideFromUrl();
}

/**
 * @param {'auto'|'melder'|'admin'|null} role
 */
export function setTechnikDemoRoleOverride(role) {
    if (!isTechnikDemoRoleUiEnabled()) {
        demoRoleOverride = null;
        return;
    }
    demoRoleOverride = role === 'melder' || role === 'admin' || role === 'auto' ? role : null;
}

export function getTechnikDemoRoleOverride() {
    return isTechnikDemoRoleUiEnabled() ? demoRoleOverride : null;
}

/**
 * @param {string} [email]
 */
export async function resolveTechnikAdmin(email) {
    const em = String(email || getTechnikAccount().email || '')
        .trim()
        .toLowerCase();
    const result = await resolveBetriebAdminAccess(em);
    cache.set({
        email: em,
        isAdmin: !!result.isAdmin,
        sources: result.sources || [],
        demo: !!result.demo
    });
    announceStaffAccess('technik', {
        email: em,
        isAdmin: !!result.isAdmin,
        role: result.isAdmin ? 'leitung' : 'melder',
        sources: result.sources || []
    });
    return result;
}

export function getTechnikAdminAccessCache() {
    return cache.get();
}

/**
 * Effektive Admin-Rolle inkl. optionaler Demo-Vorschau.
 * @param {string} [email]
 */
export function isTechnikAdmin(email) {
    const override = getTechnikDemoRoleOverride();
    if (override === 'admin') return true;
    if (override === 'melder') return false;
    return cache.isAdmin(email || getTechnikAccount().email);
}

/**
 * Wann Role-Chips zeigen (nur Demo-URL).
 */
export function canShowTechnikRoleSwitcher() {
    if (!isTechnikDemoRoleUiEnabled()) return false;
    return canUseAppRoleSwitcher(['melder', 'admin'], {
        demoRoleUi: true,
        privilegedRoles: ['admin']
    });
}
