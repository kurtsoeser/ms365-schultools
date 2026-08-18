(function () {
    'use strict';

    var script = document.currentScript;
    var SESSION_KEY = 'ms365-access-granted-v1';

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
    injectPublishedStamp();
})();
