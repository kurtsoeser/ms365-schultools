(function () {
    'use strict';

    var script = document.currentScript;
    var SESSION_KEY = 'ms365-access-granted-v1';
    var ADMIN_SESSION_KEY = 'ms365-admin-access-granted-v1';
    var ACCESS_OVERRIDE_KEY = 'ms365-schooltool-access-override-v1';

    function safeLoadJson(key) {
        try {
            var raw = localStorage.getItem(key);
            if (!raw) return null;
            return JSON.parse(String(raw));
        } catch (e) {
            return null;
        }
    }

    function effectiveUserAccessConfig(config) {
        var override = safeLoadJson(ACCESS_OVERRIDE_KEY);
        /* Repo-Config enabled:false schlägt lokales Override (PIN bewusst abgeschaltet). */
        var enabled =
            config && config.enabled === false
                ? false
                : override && typeof override.enabled === 'boolean'
                  ? override.enabled
                  : !!(config && config.enabled !== false);

        var pinsFromOverride = override && Array.isArray(override.pins) ? override.pins : null;
        var pins =
            pinsFromOverride && pinsFromOverride.length
                ? pinsFromOverride
                : config && Array.isArray(config.pins)
                  ? config.pins
                  : [];

        return { enabled: enabled, pins: pins };
    }

    function adminPinsFromConfig(config) {
        if (!config) return [];
        if (Array.isArray(config.adminPins) && config.adminPins.length) return config.adminPins;
        if (typeof config.adminPin === 'string' && config.adminPin) return [config.adminPin];
        return [];
    }

    function injectScript(fileName, marker, asModule) {
        try {
            if (!script || !script.src) return;
            if (document.querySelector('script[' + marker + '="1"]')) return;
            var s = document.createElement('script');
            s.src = new URL(fileName, script.src).href;
            if (asModule) s.type = 'module';
            else s.defer = true;
            s.setAttribute(marker, '1');
            (document.head || document.documentElement).appendChild(s);
        } catch (e) {
            /* ignore */
        }
    }

    function injectContextBar() {
        if (/\/welcome\.html(?:\?|#|$)/i.test(location.pathname)) return;
        if (/\/ms365-schooltool\.html(?:\?|#|$)/i.test(location.pathname)) return;
        injectScript('context-bar.js', 'data-ms365-context-bar', false);
    }

    function injectPublishedStamp() {
        if (/\/welcome\.html(?:\?|#|$)/i.test(location.pathname)) return;
        injectScript('app-published-stamp.js', 'data-ms365-published-stamp', true);
    }

    if (/\/welcome\.html(?:\?|#|$)/i.test(location.pathname)) return;
    /* Hilfe/Datenschutz ohne erneute PIN – auch in einem neuen Tab lesbar. */
    var isHelpPage = /\/hilfe\.html(?:\?|#|$)/i.test(location.pathname);

    var config = typeof window !== 'undefined' ? window.MS365_ACCESS_CONFIG : null;

    var isAdminPage =
        /\/admin\.html(?:\?|#|$)/i.test(location.pathname) ||
        /\/tools\/license-backend-setup\.html(?:\?|#|$)/i.test(location.pathname);
    if (isAdminPage) {
        /* Admin-UI immer laden – Zugangsprüfung läuft über MSAL + operatorUpns (API).
           Kein Soft-Gate mit pointer-events:none (wirkt sonst „tot“). */
        injectScript('operator-access.js', 'data-ms365-operator-access', false);
        injectContextBar();
        injectPublishedStamp();
        injectScript('app-paths.js', 'data-ms365-app-paths', true);
        injectScript('app-paths-boot.js', 'data-ms365-app-paths-boot', true);
        /* weiter: kein Early-Return, damit unten nichts „hängt“ */
    }

    var userAccess = effectiveUserAccessConfig(config);
    var needsPin = !!(userAccess && userAccess.enabled !== false && Array.isArray(userAccess.pins) && userAccess.pins.length);
    var hasAdminSession = sessionStorage.getItem(ADMIN_SESSION_KEY) === '1';
    if (!isAdminPage && !isHelpPage && needsPin && sessionStorage.getItem(SESSION_KEY) !== '1' && !hasAdminSession) {
        var welcome = 'welcome.html';
        if (script && script.src) {
            try {
                welcome = new URL('../../welcome.html', script.src).href;
            } catch (e) {
                /* keep relative fallback */
            }
        }
        var ret = location.pathname + location.search + location.hash;
        var sep = welcome.indexOf('?') >= 0 ? '&' : '?';
        location.replace(welcome + sep + 'return=' + encodeURIComponent(ret));
        return;
    }

    injectContextBar();
    injectPublishedStamp();
    injectScript('app-paths.js', 'data-ms365-app-paths', true);
    injectScript('app-paths-boot.js', 'data-ms365-app-paths-boot', true);
    injectScript('operator-access.js', 'data-ms365-operator-access', false);
    /* Phase 3: Lizenz-Gate – nicht auf Admin/Setup (dort greift isExemptPath zusätzlich) */
    if (!isAdminPage && !isHelpPage) {
        injectScript('license-api-client.js', 'data-ms365-license-api', false);
        injectScript('license-gate-core.js', 'data-ms365-license-gate-core', false);
        injectScript('license-gate.js', 'data-ms365-license-gate', false);
    }
})();
