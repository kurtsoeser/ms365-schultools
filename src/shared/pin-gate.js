(function () {
    'use strict';

    var script = document.currentScript;
    var SESSION_KEY = 'ms365-access-granted-v1';

    function injectContextBar() {
        try {
            if (!script || !script.src) return;
            if (/\/welcome\.html(?:\?|#|$)/i.test(location.pathname)) return;
            if (/\/ms365-schooltool\.html(?:\?|#|$)/i.test(location.pathname)) return;
            var src = new URL('context-bar.js', script.src).href;
            if (document.querySelector('script[data-ms365-context-bar="1"]')) return;
            var s = document.createElement('script');
            s.src = src;
            s.defer = true;
            s.setAttribute('data-ms365-context-bar', '1');
            (document.head || document.documentElement).appendChild(s);
        } catch (e) {
            /* ignore */
        }
    }

    if (/\/welcome\.html(?:\?|#|$)/i.test(location.pathname)) return;
    /* Hilfe/Datenschutz ohne erneute PIN – auch in einem neuen Tab lesbar. */
    var isHelpPage = /\/hilfe\.html(?:\?|#|$)/i.test(location.pathname);

    var config = typeof window !== 'undefined' ? window.MS365_ACCESS_CONFIG : null;
    var needsPin =
        !!(config && config.enabled !== false && Array.isArray(config.pins) && config.pins.length);
    if (!isHelpPage && needsPin && sessionStorage.getItem(SESSION_KEY) !== '1') {
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
})();
