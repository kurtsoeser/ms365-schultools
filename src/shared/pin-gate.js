(function () {
    'use strict';

    var SESSION_KEY = 'ms365-access-granted-v1';

    if (/\/welcome\.html(?:\?|#|$)/i.test(location.pathname)) return;
    if (sessionStorage.getItem(SESSION_KEY) === '1') return;

    var config = typeof window !== 'undefined' ? window.MS365_ACCESS_CONFIG : null;
    if (!config || config.enabled === false) return;
    if (!Array.isArray(config.pins) || !config.pins.length) return;

    var script = document.currentScript;
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
})();
