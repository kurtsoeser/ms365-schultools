/**
 * Frontend-Planer (Freistellungen, Schularbeiten, …): Sticky-Header wie Dashboard,
 * aber Suche / Datensicherung / Schulregister nur für Schul-IT bzw. Global Admin.
 */
import { getDashboardPersonas } from './dashboard-persona-session.js';

const FRONTEND_PLANNER_RE =
    /\/tools\/(?:freistellung-planer|lehrer-freistellung-planer|schularbeiten-planer|schulaktivitaeten-planer|projektwochen)\.html(?:\?|#|$)/i;

/**
 * @param {string} [pathname]
 */
export function isFrontendPlannerPage(pathname) {
    const p = String(
        pathname != null ? pathname : typeof window !== 'undefined' ? window.location.pathname : ''
    ).replace(/\\/g, '/');
    return FRONTEND_PLANNER_RE.test(p);
}

/**
 * @param {import('./dashboard-audience-resolve.js').DashboardPersonas|null|undefined} personas
 */
export function showItAppChrome(personas) {
    if (!personas || !personas.loggedIn) return false;
    if (personas.isIt || personas.globalAdmin || personas.designatedSchoolIt) return true;
    try {
        if (
            window.ms365OperatorAccess &&
            typeof window.ms365OperatorAccess.isCurrentUserOperator === 'function' &&
            window.ms365OperatorAccess.isCurrentUserOperator()
        ) {
            return true;
        }
    } catch {
        /* ignore */
    }
    return false;
}

/** Synchron (Header-Bau in msal-auth-ui) – auf Planer-Seiten IT-Chrome erst nach Persona. */
export function shouldShowItAppChromeSync() {
    if (typeof document === 'undefined' || !isFrontendPlannerPage()) return true;
    let loggedIn = false;
    try {
        loggedIn =
            typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
    } catch {
        loggedIn = false;
    }
    if (!loggedIn) return false;
    try {
        const cached = window.__ms365DashboardPersonasLast;
        if (cached && cached.loggedIn) return showItAppChrome(cached);
    } catch {
        /* ignore */
    }
    return false;
}

const HEADER_IT_IDS = [
    'dashHeaderSearchSlot',
    'ms365BackupHeader',
    'ms365HeaderSchulregister',
    'dashRegisterToolbar'
];

/**
 * @param {boolean} showIt
 */
export function applyFrontendPlannerItChrome(showIt) {
    const hide = !showIt;
    HEADER_IT_IDS.forEach(function (id) {
        const el = document.getElementById(id);
        if (!el) return;
        el.hidden = hide;
        el.setAttribute('aria-hidden', hide ? 'true' : 'false');
    });
    document.querySelectorAll('[data-ms365-frontend-it-only], [data-sa-it-only]').forEach(function (el) {
        el.hidden = hide;
    });
    if (typeof document !== 'undefined' && document.documentElement) {
        document.documentElement.toggleAttribute('data-frontend-planner-it', showIt);
    }
}

export async function refreshFrontendPlannerChromeAudience() {
    if (typeof document === 'undefined' || !isFrontendPlannerPage()) return;

    let loggedIn = false;
    try {
        loggedIn =
            typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
    } catch {
        loggedIn = false;
    }
    if (!loggedIn) {
        applyFrontendPlannerItChrome(false);
        return;
    }

    try {
        const cached = window.__ms365DashboardPersonasLast;
        if (cached && cached.loggedIn) {
            applyFrontendPlannerItChrome(showItAppChrome(cached));
        }
    } catch {
        /* ignore */
    }

    try {
        const personas = await getDashboardPersonas({ force: false });
        applyFrontendPlannerItChrome(showItAppChrome(personas));
    } catch {
        applyFrontendPlannerItChrome(false);
    }
}

/** Von msal-auth-ui nach jedem Header-Rebuild aufrufen. */
export function syncFrontendPlannerHeaderChromeNow() {
    if (!isFrontendPlannerPage()) return;
    applyFrontendPlannerItChrome(shouldShowItAppChromeSync());
}

if (typeof window !== 'undefined') {
    window.ms365SyncFrontendPlannerHeaderChrome = syncFrontendPlannerHeaderChromeNow;
}

export function bootFrontendPlannerChromePolicy() {
    if (typeof window === 'undefined' || window.__ms365FrontendPlannerChromeBoot) return;
    window.__ms365FrontendPlannerChromeBoot = true;

    window.ms365SyncFrontendPlannerHeaderChrome = syncFrontendPlannerHeaderChromeNow;

    const run = function () {
        refreshFrontendPlannerChromeAudience();
    };

    run();
    window.addEventListener('ms365-auth-widget-ready', run);
    window.addEventListener('ms365-auth-state-changed', function () {
        run();
        import('./dashboard-persona-session.js')
            .then(function (m) {
                if (m && typeof m.refreshDashboardPersonaSession === 'function') {
                    return m.refreshDashboardPersonaSession({ force: true });
                }
            })
            .then(run)
            .catch(function () {
                run();
            });
    });
    window.addEventListener('ms365-dashboard-persona-ready', run);
    if (typeof document !== 'undefined') {
        document.addEventListener('DOMContentLoaded', run);
    }
}
