/**
 * Tool-/Planer-Seiten: Sticky-Header wie Dashboard,
 * aber Suche / Datensicherung / Organigramm / Stammdaten nur im bewussten IT-Bereich
 * (isIt + Shell „full“), nie für reine Lehrkräfte/Schüler.
 */
import { getDashboardPersonas } from './dashboard-persona-session.js';
import { readItShellChoice } from './dashboard-layout-mode.js';

const FRONTEND_PLANNER_RE =
    /\/tools\/(?:freistellung-planer|lehrer-freistellung-planer|schularbeiten-planer|schulaktivitaeten-planer|projektwochen|srdp-anmeldung|termine-hub)\.html(?:\?|#|$)/i;

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
 * Alle Werkzeug-Seiten (nicht Dashboard/Admin) – IT-Header nur bei IT-Shell.
 * @param {string} [pathname]
 */
export function isToolChromePage(pathname) {
    if (typeof document !== 'undefined' && document.body) {
        if (document.body.classList.contains('page-dashboard')) return false;
        if (document.body.classList.contains('admin-app-page')) return false;
        if (document.body.classList.contains('app-shell-chrome')) return true;
    }
    const p = String(
        pathname != null ? pathname : typeof window !== 'undefined' ? window.location.pathname : ''
    ).replace(/\\/g, '/');
    if (/\/(?:admin|index|tenant|hilfe|landing)\.html(?:\?|#|$)/i.test(p)) return false;
    return /\/tools\/[^/]+\.html(?:\?|#|$)/i.test(p) || isFrontendPlannerPage(p);
}

/**
 * @param {import('./dashboard-audience-resolve.js').DashboardPersonas|null|undefined} personas
 */
export function showItAppChrome(personas) {
    if (!personas || !personas.loggedIn) return false;
    let canIt = !!(personas.isIt || personas.globalAdmin || personas.designatedSchoolIt);
    if (!canIt) {
        try {
            if (
                window.ms365OperatorAccess &&
                typeof window.ms365OperatorAccess.isCurrentUserOperator === 'function' &&
                window.ms365OperatorAccess.isCurrentUserOperator()
            ) {
                canIt = true;
            }
        } catch {
            /* ignore */
        }
    }
    if (!canIt) return false;
    // Portal-first: Organigramm/Stammdaten nur nach bewusstem IT-Bereich
    try {
        return readItShellChoice() === 'full';
    } catch {
        return false;
    }
}

/** Synchron (Header-Bau in msal-auth-ui) – IT-Chrome erst nach Persona + IT-Shell. */
export function shouldShowItAppChromeSync() {
    if (typeof document === 'undefined' || !isToolChromePage()) return true;
    let loggedIn = false;
    try {
        loggedIn =
            typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
    } catch {
        loggedIn = false;
    }
    if (!loggedIn) return false;
    try {
        if (readItShellChoice() !== 'full') return false;
    } catch {
        return false;
    }
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
    'ms365HeaderSchulOrganigramm',
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
        // Gleiches Signal wie Dashboard-IT-Shell – CSS blendet Organigramm/Stammdaten aus
        document.documentElement.toggleAttribute('data-dash-it-shell', showIt);
    }
}

export async function refreshFrontendPlannerChromeAudience() {
    if (typeof document === 'undefined' || !isToolChromePage()) return;

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
        } else {
            applyFrontendPlannerItChrome(false);
        }
    } catch {
        applyFrontendPlannerItChrome(false);
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
    if (!isToolChromePage()) return;
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

    // Portal-Schale (b5): kompakte Kopfzeile, kein IT-Flash
    if (typeof document !== 'undefined' && document.body && isFrontendPlannerPage()) {
        document.body.classList.add('ms365-portal-shell');
        document.documentElement.setAttribute('data-ms365-portal-shell', '1');
    }

    // Portal-first: sofort IT-Chrome verstecken, bis Persona + Shell geklärt sind
    if (isToolChromePage()) applyFrontendPlannerItChrome(false);

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
