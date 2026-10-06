/**
 * Einheitlicher Sticky-Header (wie Dashboard) auf allen App-Seiten mit msal-auth-ui.
 */
import {
    buildAppChromeHeaderHtml,
    detectAppHeaderMode,
    ensureAppChromeStylesheets,
    markToolSubheader,
    normalizeAppHeaderActionsOrder,
    resolveAppRootHref
} from './app-header-chrome.js';

export { resolveAppRootHref };

export function shouldMountAppGlobalHeader() {
    if (typeof document === 'undefined' || !document.body) return false;
    const p = String(window.location.pathname || '').replace(/\\/g, '/');
    if (/\/welcome\.html(?:\?|#|$)/i.test(p)) return false;
    if (/\/landing\//i.test(p)) return false;
    return !document.getElementById('dashCompactHeader');
}

/**
 * Stellt sicher, dass der App-Header (sticky) mit einheitlichen Styles vorhanden ist.
 * @returns {boolean}
 */
export function ensureAppHeaderChrome() {
    if (typeof document === 'undefined' || !document.body) return false;

    document.body.classList.add('app-shell-chrome');
    ensureAppChromeStylesheets();

    if (document.getElementById('dashCompactHeader')) {
        markToolSubheader();
        const slot = document.getElementById('adminAppTopActions');
        normalizeAppHeaderActionsOrder(slot);
        import('./app-tool-chrome.js')
            .then(function (m) {
                if (m && typeof m.bootToolPageChrome === 'function') m.bootToolPageChrome();
            })
            .catch(function () {
                /* ignore */
            });
        return true;
    }

    if (!shouldMountAppGlobalHeader()) return false;
    return mountAppGlobalHeader();
}

export function mountAppGlobalHeader() {
    if (typeof document === 'undefined' || !document.body) return false;

    document.body.classList.add('app-shell-chrome');
    ensureAppChromeStylesheets();

    if (document.getElementById('dashCompactHeader')) {
        markToolSubheader();
        return false;
    }

    if (!shouldMountAppGlobalHeader()) return false;

    const container = document.querySelector('.container.page-card, .container');
    if (!container) return false;

    const mode = detectAppHeaderMode();
    const wrap = document.createElement('div');
    wrap.innerHTML = buildAppChromeHeaderHtml(mode);
    const header = wrap.firstElementChild;
    if (!header) return false;

    const first = container.firstElementChild;
    if (first) container.insertBefore(header, first);
    else container.appendChild(header);

    markToolSubheader();

    import('./app-tool-chrome.js')
        .then(function (m) {
            if (m && typeof m.bootToolPageChrome === 'function') m.bootToolPageChrome();
        })
        .catch(function () {
            /* ignore */
        });

    import('./dashboard-compact-header.js')
        .then(function (m) {
            if (m && typeof m.mountDashboardCompactHeader === 'function') {
                m.mountDashboardCompactHeader();
            }
        })
        .catch(function () {
            /* ignore */
        });

    return true;
}

export { normalizeAppHeaderActionsOrder, ensureAppChromeStylesheets, buildAppChromeHeaderHtml };
