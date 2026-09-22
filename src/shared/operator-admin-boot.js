/**
 * Admin-Seiten: Wenn keine Admin-PIN-Session, MS365-Betreiber abwarten und freischalten.
 * Sonst Redirect zu welcome.html?mode=admin.
 */
(function () {
    'use strict';

    var ADMIN_SESSION_KEY = 'ms365-admin-access-granted-v1';

    function hasAdminSession() {
        try {
            return sessionStorage.getItem(ADMIN_SESSION_KEY) === '1';
        } catch (e) {
            return false;
        }
    }

    function goWelcomeAdmin() {
        var ret = location.pathname + location.search + location.hash;
        var welcome = 'welcome.html';
        try {
            var scripts = document.getElementsByTagName('script');
            for (var i = scripts.length - 1; i >= 0; i--) {
                var src = scripts[i].src || '';
                if (/operator-admin-boot\.js(\?|$)/i.test(src)) {
                    welcome = new URL('../../welcome.html', src).href;
                    break;
                }
            }
        } catch (e) {
            /* keep relative */
        }
        var sep = welcome.indexOf('?') >= 0 ? '&' : '?';
        location.replace(welcome + sep + 'return=' + encodeURIComponent(ret) + '&mode=admin');
    }

    function tryOperator() {
        if (hasAdminSession()) return true;
        if (
            window.ms365OperatorAccess &&
            typeof window.ms365OperatorAccess.grantAdminSessionIfOperator === 'function'
        ) {
            return !!window.ms365OperatorAccess.grantAdminSessionIfOperator();
        }
        return false;
    }

    function finishOk() {
        document.documentElement.removeAttribute('data-ms365-admin-boot');
    }

    function boot() {
        if (hasAdminSession()) {
            finishOk();
            return;
        }

        document.documentElement.setAttribute('data-ms365-admin-boot', '1');

        var tries = 0;
        var max = 40;

        function tick() {
            tries++;
            if (tryOperator()) {
                finishOk();
                return;
            }
            var loggedIn =
                typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
            if (loggedIn && window.ms365OperatorAccess) {
                // Angemeldet, aber kein Betreiber → PIN-Welcome
                goWelcomeAdmin();
                return;
            }
            if (tries >= max) {
                goWelcomeAdmin();
                return;
            }
            setTimeout(tick, 150);
        }

        window.addEventListener('ms365-auth-state-changed', function () {
            if (tryOperator()) finishOk();
            else if (
                typeof window.ms365AuthIsLoggedIn === 'function' &&
                window.ms365AuthIsLoggedIn() &&
                window.ms365OperatorAccess &&
                !window.ms365OperatorAccess.isCurrentUserOperator()
            ) {
                goWelcomeAdmin();
            }
        });

        if (typeof window.ms365AuthEnsureInitialized === 'function') {
            Promise.resolve(window.ms365AuthEnsureInitialized())
                .then(function () {
                    tick();
                })
                .catch(function () {
                    tick();
                });
        } else {
            tick();
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
    } else {
        boot();
    }
})();
