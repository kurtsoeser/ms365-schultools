/**
 * Admin-Seiten: Zugang nur nach License-API /admin/me (Betreiber in Azure).
 */
(function () {
    'use strict';

    var BLOCK_ID = 'ms365AdminOperatorBlock';

    function finishOk() {
        document.documentElement.removeAttribute('data-ms365-admin-boot');
        var block = document.getElementById(BLOCK_ID);
        if (block) block.remove();
    }

    function welcomeHref() {
        var href = 'welcome.html';
        try {
            var scripts = document.getElementsByTagName('script');
            for (var i = scripts.length - 1; i >= 0; i--) {
                var src = scripts[i].src || '';
                if (/operator-admin-boot\.js(\?|$)/i.test(src)) {
                    href = new URL('../../welcome.html', src).href;
                    break;
                }
            }
        } catch (e) {
            /* keep */
        }
        return href;
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
            (document.body || document.documentElement).appendChild(el);
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
                    var loginFn =
                        typeof window.ms365AuthLogin === 'function'
                            ? window.ms365AuthLogin
                            : typeof window.ms365AuthSwitchAccount === 'function'
                              ? window.ms365AuthSwitchAccount
                              : null;
                    if (loginFn) {
                        Promise.resolve(loginFn()).catch(function () {
                            /* ignore */
                        });
                    } else {
                        window.alert('Anmeldung noch nicht bereit – bitte 1–2 Sekunden warten und erneut klicken.');
                    }
                });
                actions.appendChild(btn);
            }
            var back = document.createElement('a');
            back.className = 'btn alt';
            back.href = welcomeHref();
            back.innerHTML = '<i class="bi bi-arrow-left"></i> Zur App';
            actions.appendChild(back);
        }
        return el;
    }

    function showNeedLogin() {
        ensureBlock('Bitte mit dem Betreiber-Microsoft-365-Konto anmelden.', true);
    }

    function showNotOperator() {
        ensureBlock(
            'Angemeldet, aber kein Betreiber-Konto. Die Freigabe liegt nur auf dem Server (Azure App Settings).',
            false
        );
    }

    function showChecking() {
        ensureBlock('Betreiber-Zugang wird geprüft …', false);
    }

    function showApiMissing() {
        ensureBlock(
            'License-API ist nicht konfiguriert (MS365_LICENSE_API.baseUrl). Admin kann nicht geprüft werden.',
            false
        );
    }

    /** @returns {Promise<'ok'|'login'|'denied'|'wait'|'config'>} */
    async function evaluate() {
        var loggedIn =
            typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
        if (!loggedIn) {
            showNeedLogin();
            return 'login';
        }

        var cfg = window.MS365_LICENSE_API || {};
        if (!String(cfg.baseUrl || '').trim()) {
            showApiMissing();
            return 'config';
        }

        if (!window.ms365OperatorAccess || typeof window.ms365OperatorAccess.refreshOperatorStatus !== 'function') {
            showChecking();
            return 'wait';
        }

        showChecking();
        var ok = await window.ms365OperatorAccess.refreshOperatorStatus({ force: true });
        if (ok) {
            finishOk();
            return 'ok';
        }
        showNotOperator();
        return 'denied';
    }

    function boot() {
        document.documentElement.setAttribute('data-ms365-admin-boot', '1');
        showNeedLogin();

        var running = false;
        function run() {
            if (running) return;
            running = true;
            Promise.resolve(evaluate())
                .catch(function () {
                    showNotOperator();
                })
                .finally(function () {
                    running = false;
                });
        }

        window.addEventListener('ms365-auth-state-changed', run);
        window.addEventListener('ms365-auth-widget-ready', run);

        if (typeof window.ms365AuthEnsureInitialized === 'function') {
            Promise.resolve(window.ms365AuthEnsureInitialized())
                .then(run)
                .catch(run);
        } else {
            setTimeout(run, 400);
        }
        setTimeout(run, 800);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
    } else {
        boot();
    }
})();
