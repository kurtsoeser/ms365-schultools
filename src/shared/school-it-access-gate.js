/**
 * Stammdaten und Dashboard-Verwaltung: nur Schul-IT (Global Admin / Plattform-Betreiber).
 */
import { getDashboardPersonas, clearDashboardPersonaCache } from './dashboard-persona-session.js';

/**
 * @param {{ redirectTo?: string, force?: boolean }} [opts]
 */
export async function enforceSchoolItAccess(opts) {
    const redirectTo = (opts && opts.redirectTo) || 'index.html';
    const personas = await getDashboardPersonas({ force: !!(opts && opts.force) });
    if (!personas.loggedIn) return { allowed: false, reason: 'not-logged-in', personas };
    if (!personas.isIt) {
        const url = new URL(redirectTo, window.location.href);
        url.searchParams.set('access', 'denied');
        url.searchParams.set('reason', 'school-it');
        window.location.replace(url.pathname + url.search);
        return { allowed: false, reason: 'not-it', personas };
    }
    return { allowed: true, personas };
}

export function mountSchoolItAccessGate(opts) {
    const run = () => {
        enforceSchoolItAccess(opts).catch(() => {
            /* ignore */
        });
    };

    if (typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn()) {
        run();
    } else {
        const onAuth = () => {
            clearDashboardPersonaCache();
            run();
        };
        window.addEventListener('ms365-auth-state-changed', onAuth, { once: true });
        window.addEventListener('ms365-auth-widget-ready', () => {
            if (typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn()) {
                onAuth();
            }
        });
    }
}
