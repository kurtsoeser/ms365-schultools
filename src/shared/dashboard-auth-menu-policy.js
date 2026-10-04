/**
 * Konto-Menü: Tenant-ID und Aktionsprotokoll nur für Schul-IT (Global Admin / Betreiber).
 */
import { getDashboardPersonas } from './dashboard-persona-session.js';

export function applyAuthMenuItAudience(showItExtras) {
    const tenantRow = document.querySelector('.ms365-auth-menu__ctx-row--tenant');
    const actionLog = document.getElementById('ms365AuthActionLogLink');
    const hide = !showItExtras;
    if (tenantRow) tenantRow.hidden = hide;
    if (actionLog) actionLog.hidden = hide;
}

export async function refreshAuthMenuAudiencePolicy() {
    try {
        if (typeof window.ms365AuthIsLoggedIn !== 'function' || !window.ms365AuthIsLoggedIn()) {
            applyAuthMenuItAudience(false);
            return;
        }
    } catch {
        applyAuthMenuItAudience(false);
        return;
    }

    try {
        const personas = await getDashboardPersonas({ force: false });
        applyAuthMenuItAudience(!!(personas && personas.isIt));
    } catch {
        applyAuthMenuItAudience(false);
    }
}

export function bootAuthMenuAudiencePolicy() {
    if (typeof window !== 'undefined' && window.__ms365AuthMenuPolicyBoot) return;
    if (typeof window !== 'undefined') window.__ms365AuthMenuPolicyBoot = true;

    refreshAuthMenuAudiencePolicy();
    window.addEventListener('ms365-auth-state-changed', () => {
        refreshAuthMenuAudiencePolicy();
    });
    window.addEventListener('ms365-dashboard-persona-ready', () => {
        refreshAuthMenuAudiencePolicy();
    });
    window.addEventListener('ms365-auth-widget-ready', () => {
        refreshAuthMenuAudiencePolicy();
    });
}
