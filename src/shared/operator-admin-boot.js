/**
 * Admin-Seiten: nur Betreiber-MS365 (operatorUpns). Kein Master-PIN-Fallback.
 */
(function () {
    'use strict';

    var ADMIN_SESSION_KEY = 'ms365-admin-access-granted-v1';
    var BLOCK_ID = 'ms365AdminOperatorBlock';

    function hasAdminSession() {
        try {
            return sessionStorage.getItem(ADMIN_SESSION_KEY) === '1';
        } catch (e) {
            return false;
        }
    }

    function finishOk() {
        document.documentElement.removeAttribute('data-ms365-admin-boot');
        var block = document.getElementById(BLOCK_ID);
        if (block) block.remove();
    }

    function ensureBlock(message, showLogin) {
        var el = document.getElementById(BLOCK_ID);
        if (!el) {
            el = document.createElement('div');
            el.id = BLOCK_ID;
            el.className = 'ms365-admin-op-block';
            el.setAttribute('role', 'dialog');
            el.setAttribute('aria-modal', 'true');
            el.innerHTML =
                '<div class="ms365-admin-op-block__card">' +
                '  <div class="ms365-admin-op-block__icon" aria-hidden="true"><i class="bi bi-shield-lock"></i></div>' +
                '  <h2 class="ms365-admin-op-block__title">Admin nur für Betreiber</h2>' +
                '  <p class="ms365-admin-op-block__text" id="ms365AdminOpBlockText"></p>' +
                '  <div class="ms365-admin-op-block__actions" id="ms365AdminOpBlockActions"></div>' +
                '</div>';
            document.body.appendChild(el);
        }
        var text = document.getElementById('ms365AdminOpBlockText');
        if (text) text.textContent = message || '';
        var actions = document.getElementById('ms365AdminOpBlockActions');
        if (actions) {
            actions.replaceChildren();
            if (showLogin) {
                var btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn';
                btn.innerHTML = '<i class="bi bi-box-arrow-in-right"></i> Mit MS365 anmelden';
                btn.addEventListener('click', function () {
                    if (typeof window.ms365AuthLogin === 'function') {
                        Promise.resolve(window.ms365AuthLogin()).catch(function () {
                            /* ignore */
                        });
                    } else if (typeof window.ms365AuthLoginPopup === 'function') {
                        Promise.resolve(window.ms365AuthLoginPopup()).catch(function () {
                            /* ignore */
                        });
                    } else {
                        window.alert('Bitte über das Konto-Menü oben rechts anmelden.');
                    }
                });
                actions.appendChild(btn);
            }
            var back = document.createElement('a');
            back.className = 'btn alt';
            back.href = 'welcome.html';
            try {
                var scripts = document.getElementsByTagName('script');
                for (var i = scripts.length - 1; i >= 0; i--) {
                    var src = scripts[i].src || '';
                    if (/operator-admin-boot\.js(\?|$)/i.test(src)) {
                        back.href = new URL('../../welcome.html', src).href;
                        break;
                    }
                }
            } catch (e) {
                /* keep */
            }
            back.innerHTML = '<i class="bi bi-arrow-left"></i> Zur App';
            actions.appendChild(back);
        }
        return el;
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

    function boot() {
        if (hasAdminSession()) {
            finishOk();
            return;
        }

        document.documentElement.setAttribute('data-ms365-admin-boot', '1');

        var tries = 0;
        var max = 50;

        function tick() {
            tries++;
            if (tryOperator()) {
                finishOk();
                return;
            }
            var loggedIn =
                typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
            if (loggedIn && window.ms365OperatorAccess) {
                ensureBlock(
                    'Angemeldet, aber kein Betreiber-Konto. Admin ist nur für hinterlegte UPNs (z. B. kurt@kurtsoeser.at).',
                    false
                );
                return;
            }
            if (tries >= max) {
                ensureBlock(
                    'Bitte mit dem Betreiber-Konto (kurt@kurtsoeser.at) anmelden.',
                    true
                );
                return;
            }
            setTimeout(tick, 150);
        }

        window.addEventListener('ms365-auth-state-changed', function () {
            if (tryOperator()) {
                finishOk();
                return;
            }
            if (
                typeof window.ms365AuthIsLoggedIn === 'function' &&
                window.ms365AuthIsLoggedIn() &&
                window.ms365OperatorAccess &&
                !window.ms365OperatorAccess.isCurrentUserOperator()
            ) {
                ensureBlock(
                    'Angemeldet, aber kein Betreiber-Konto. Admin ist nur für hinterlegte UPNs (z. B. kurt@kurtsoeser.at).',
                    false
                );
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
