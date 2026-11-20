/**
 * Aufgaben: Admin aus Entra/Schulleitung (Betrieb-Helper).
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

export function getAufgabenAccount() {
    return getBetriebToolAccount();
}

export function isAufgabenDemoRoleUiEnabled() {
    return readDemoRoleOverrideFromUrl();
}

/**
 * @param {'auto'|'melder'|'admin'|null} role
 */
export function setAufgabenDemoRoleOverride(role) {
    if (!isAufgabenDemoRoleUiEnabled()) {
        demoRoleOverride = null;
        return;
    }
    demoRoleOverride = role === 'melder' || role === 'admin' || role === 'auto' ? role : null;
}

export function getAufgabenDemoRoleOverride() {
    return isAufgabenDemoRoleUiEnabled() ? demoRoleOverride : null;
}

/**
 * @param {string} [email]
 */
export async function resolveAufgabenAdmin(email) {
    const em = String(email || getAufgabenAccount().email || '')
        .trim()
        .toLowerCase();
    const result = await resolveBetriebAdminAccess(em);
    cache.set({
        email: em,
        isAdmin: !!result.isAdmin,
        sources: result.sources || [],
        demo: !!result.demo
    });
    announceStaffAccess('aufgaben', {
        email: em,
        isAdmin: !!result.isAdmin,
        role: result.isAdmin ? 'leitung' : 'melder',
        sources: result.sources || []
    });
    return result;
}

export function getAufgabenAdminAccessCache() {
    return cache.get();
}

/**
 * Effektive Admin-Rolle inkl. optionaler Demo-Vorschau.
 * @param {string} [email]
 */
export function isAufgabenAdmin(email) {
    const override = getAufgabenDemoRoleOverride();
    if (override === 'admin') return true;
    if (override === 'melder') return false;
    return cache.isAdmin(email || getAufgabenAccount().email);
}

/**
 * Wann Role-Chips zeigen (nur Demo-URL).
 */
export function canShowAufgabenRoleSwitcher() {
    if (!isAufgabenDemoRoleUiEnabled()) return false;
    return canUseAppRoleSwitcher(['melder', 'admin'], {
        demoRoleUi: true,
        privilegedRoles: ['admin']
    });
}
