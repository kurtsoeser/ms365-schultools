(function () {
    'use strict';

    const MSAL_LOADER_IMPORT = (function () {
        // Vite-Bundle: import.meta.url zeigt auf assets/msal-auth-ui-*.js → msal-loader im selben Ordner.
        try {
            if (typeof import.meta !== 'undefined' && import.meta.url) {
                return new URL('./msal-loader.js', import.meta.url).href;
            }
        } catch (_) {
            // ignore
        }
        const needle = 'msal-auth-ui.js';
        const rel = './msal-loader.js';
        const scripts = document.getElementsByTagName('script');
        for (let i = scripts.length - 1; i >= 0; i--) {
            const src = scripts[i].src || '';
            if (src.indexOf(needle) !== -1) {
                try {
                    return new URL(rel, src).href;
                } catch (_) {}
            }
        }
        // Fallback: Site-Root (nicht document.baseURI – unter /tools/*.html wäre das falsch).
        try {
            const p = String(window.location.pathname || '/').split('?')[0].split('#')[0];
            const iTools = p.toLowerCase().indexOf('/tools/');
            const base =
                iTools !== -1
                    ? p.slice(0, iTools)
                    : p.lastIndexOf('/') > 0
                      ? p.slice(0, p.lastIndexOf('/'))
                      : '';
            const root = base.endsWith('/') ? base.slice(0, -1) : base;
            return new URL((root ? root + '/' : '/') + 'src/shared/msal-loader.js', window.location.origin).href;
        } catch (_) {
            return 'src/shared/msal-loader.js';
        }
    })();

    const DEFAULT_SCOPES = ['https://graph.microsoft.com/User.Read'];
    const POST_LOGIN_KEY = 'ms365-post-login-url';

    let msalMod = null;
    let pca = null;
    let initPromise = null;

    function $(sel, root) {
        return (root || document).querySelector(sel);
    }

    function resolveMsalConfig() {
        let cfg = window.MS365_MSAL_CONFIG;
        if (!cfg) cfg = {};
        let id = String(cfg.clientId || '').trim();
        if (!id) {
            const meta = document.querySelector('meta[name="ms365-graph-client-id"]');
            const fromMeta = meta && meta.getAttribute('content') ? meta.getAttribute('content').trim() : '';
            if (fromMeta) id = fromMeta;
        }
        if (!id) throw new Error('Keine clientId: ms365-config.js fehlt/leer oder blockiert.');
        return {
            clientId: id,
            authority: cfg.authority || 'https://login.microsoftonline.com/organizations',
            redirectUri: (cfg.redirectUri || window.location.href.split('#')[0]).trim()
        };
    }

    async function loadMsal() {
        if (msalMod) return msalMod;
        const loader = await import(/* @vite-ignore */ MSAL_LOADER_IMPORT);
        if (typeof loader.loadMsalBrowser !== 'function') {
            throw new Error('MSAL-Loader: loadMsalBrowser fehlt.');
        }
        msalMod = await loader.loadMsalBrowser();
        return msalMod;
    }

    function isInteractionRequired(e) {
        if (!e) return false;
        if (e.name === 'InteractionRequiredAuthError') return true;
        const code = String(e.errorCode || '').toLowerCase();
        if (
            code === 'interaction_required' ||
            code === 'consent_required' ||
            code === 'login_required' ||
            code === 'invalid_grant' ||
            code === 'no_account_in_silent_request' ||
            code === 'no_tokens_found' ||
            code === 'monitor_window_timeout' ||
            code === 'native_account_unavailable'
        ) {
            return true;
        }
        const msg = String((e && e.message) || '').toLowerCase();
        return (
            msg.indexOf('interaction_required') !== -1 ||
            msg.indexOf('consent_required') !== -1 ||
            msg.indexOf('login_required') !== -1 ||
            msg.indexOf('invalid_grant') !== -1 ||
            msg.indexOf('aadsts65001') !== -1 || // Consent fehlt
            msg.indexOf('aadsts50058') !== -1 || // Sitzung verloren
            msg.indexOf('aadsts70008') !== -1 || // Refresh-Token abgelaufen
            msg.indexOf('aadsts50173') !== -1 || // Refresh-Token widerrufen
            msg.indexOf('aadsts50076') !== -1 || // MFA nötig
            msg.indexOf('aadsts50079') !== -1 || // MFA registration nötig
            msg.indexOf('aadsts700084') !== -1 || // Cookie hash mismatch
            msg.indexOf('token contains an invalid signature') !== -1
        );
    }

    async function ensurePca() {
        if (initPromise) return initPromise;
        initPromise = (async () => {
            const m = await loadMsal();
            const PublicClientApplication = m.PublicClientApplication || (m.default && m.default.PublicClientApplication);
            if (!PublicClientApplication) throw new Error('MSAL: PublicClientApplication nicht gefunden.');
            const cfg = resolveMsalConfig();
            pca = new PublicClientApplication({
                auth: { clientId: cfg.clientId, authority: cfg.authority, redirectUri: cfg.redirectUri },
                // localStorage statt sessionStorage: ermöglicht Single-Sign-On zwischen Browser-Tabs
                // (Microsoft 365 Anmeldung wird übernommen, wenn der Benutzer bereits in einem
                // anderen Tab/Modul angemeldet ist).
                cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true }
            });
            await pca.initialize();
            await pca.handleRedirectPromise();

            const accounts = pca.getAllAccounts();
            if (accounts && accounts[0] && typeof pca.setActiveAccount === 'function') {
                pca.setActiveAccount(accounts[0]);
            }
            return pca;
        })();
        return initPromise;
    }

    /**
     * Versucht eine unsichtbare Single-Sign-On-Anmeldung über die bestehende
     * Microsoft-365-Browser-Sitzung (Hidden Iframe an login.microsoftonline.com).
     * Funktioniert, wenn der Benutzer in einem anderen Tab/Fenster bereits angemeldet ist
     * und Third-Party-Cookies für Microsoft erlaubt sind.
     * Wirft NICHT bei Fehlschlag (z. B. wenn kein Account vorhanden / Cookies blockiert).
     */
    async function trySsoSilent(scopes) {
        if (!pca) return null;
        try {
            const req = {
                scopes: Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES
            };
            const result = await pca.ssoSilent(req);
            if (result && result.account && typeof pca.setActiveAccount === 'function') {
                pca.setActiveAccount(result.account);
            }
            return result;
        } catch {
            return null;
        }
    }

    function getAccount() {
        if (!pca) return null;
        const a = typeof pca.getActiveAccount === 'function' ? pca.getActiveAccount() : null;
        if (a) return a;
        const all = pca.getAllAccounts();
        return all && all[0] ? all[0] : null;
    }

    function accountLabel(a) {
        if (!a) return '';
        const u = a.username ? String(a.username) : '';
        const n = a.name ? String(a.name) : '';
        if (n && u && n !== u) return n + ' (' + u + ')';
        return n || u || '';
    }

    function accountDisplayName(a) {
        if (!a) return '';
        const n = a.name ? String(a.name).trim() : '';
        const u = a.username ? String(a.username).trim() : '';
        return n || u || '';
    }

    function closeAuthMenu() {
        const menu = document.getElementById('ms365AuthMenu');
        const trigger = document.getElementById('ms365AuthBadge');
        const drop = document.getElementById('ms365AuthDropdown');
        if (menu) menu.classList.remove('is-open');
        if (trigger) trigger.setAttribute('aria-expanded', 'false');
        if (drop) drop.hidden = true;
    }

    function toggleAuthMenu() {
        const menu = document.getElementById('ms365AuthMenu');
        const trigger = document.getElementById('ms365AuthBadge');
        const drop = document.getElementById('ms365AuthDropdown');
        if (!menu || !trigger || !drop || menu.hidden) return;
        const open = !menu.classList.contains('is-open');
        menu.classList.toggle('is-open', open);
        trigger.setAttribute('aria-expanded', open ? 'true' : 'false');
        drop.hidden = !open;
    }

    function bindAuthMenuDismiss() {
        if (bindAuthMenuDismiss.bound) return;
        bindAuthMenuDismiss.bound = true;
        document.addEventListener('click', function (e) {
            const menu = document.getElementById('ms365AuthMenu');
            if (!menu || !menu.classList.contains('is-open')) return;
            if (menu.contains(e.target)) return;
            closeAuthMenu();
        });
        document.addEventListener('keydown', function (e) {
            if (e.key === 'Escape') closeAuthMenu();
        });
    }

    /**
     * Anmeldung per Redirect.
     * - Ohne opts.prompt: Microsoft entscheidet selbst (nutzt bestehende Browser-Session,
     *   zeigt Account-Auswahl nur falls nötig). So funktioniert SSO mit anderen MS-365-Tabs.
     * - Mit opts.prompt === 'select_account': erzwingt Account-Auswahl (z. B. zum Konto wechseln).
     */
    async function login(scopes, opts) {
        const instance = await ensurePca();
        try {
            sessionStorage.setItem(POST_LOGIN_KEY, window.location.href);
        } catch {
            // ignore
        }
        const req = {
            scopes: Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES,
            redirectStartPage: window.location.href
        };
        if (opts && typeof opts.prompt === 'string' && opts.prompt) {
            req.prompt = opts.prompt;
        }
        await instance.loginRedirect(req);
        // redirect -> no further code
    }

    async function switchAccount(scopes) {
        return login(scopes, { prompt: 'select_account' });
    }

    async function logout() {
        const instance = await ensurePca();
        const a = getAccount();
        try {
            sessionStorage.setItem(POST_LOGIN_KEY, window.location.href);
        } catch {
            // ignore
        }
        await instance.logoutRedirect({ account: a || undefined, postLogoutRedirectUri: window.location.href.split('#')[0] });
    }

    function looksLikeBrokenCache(e) {
        if (!e) return false;
        const msg = String((e && e.message) || '').toLowerCase();
        return (
            msg.indexOf('token contains an invalid signature') !== -1 ||
            msg.indexOf('invalid_grant') !== -1 ||
            msg.indexOf('aadsts70008') !== -1 ||
            msg.indexOf('aadsts50173') !== -1 ||
            msg.indexOf('aadsts700084') !== -1
        );
    }

    async function clearMsalCache(instance) {
        try {
            const accounts = instance && typeof instance.getAllAccounts === 'function' ? instance.getAllAccounts() : [];
            if (typeof instance.clearCache === 'function') {
                try {
                    await instance.clearCache();
                } catch {
                    // ignore – wir versuchen es danach noch manuell
                }
            }
            (accounts || []).forEach((acc) => {
                if (acc && typeof instance.logoutSilent === 'function') {
                    instance.logoutSilent({ account: acc }).catch(() => {});
                }
            });
        } catch {
            // ignore
        }
        try {
            const removeIf = (store, predicate) => {
                const keys = [];
                for (let i = 0; i < store.length; i++) {
                    const k = store.key(i);
                    if (k && predicate(k)) keys.push(k);
                }
                keys.forEach((k) => {
                    try {
                        store.removeItem(k);
                    } catch {
                        // ignore
                    }
                });
            };
            const isMsalKey = (k) =>
                k.indexOf('msal.') === 0 ||
                k.indexOf('msal-') === 0 ||
                k.indexOf('login.microsoftonline.com') !== -1 ||
                k.indexOf('login.windows.net') !== -1 ||
                /[-.]?msal[-.]/i.test(k);
            removeIf(localStorage, isMsalKey);
            removeIf(sessionStorage, isMsalKey);
        } catch {
            // ignore
        }
    }

    async function acquireToken(scopes) {
        const instance = await ensurePca();
        let accounts = instance.getAllAccounts();
        if (!accounts.length) {
            await login(scopes);
            throw new Error('Weiterleitung zur Anmeldung …');
        }
        const a = getAccount() || accounts[0];
        const req = { scopes: Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES, account: a };
        try {
            return (await instance.acquireTokenSilent(req)).accessToken;
        } catch (e) {
            // Bei kaputtem/abgelaufenem MSAL-Cache („Token contains an invalid signature",
            // invalid_grant, AADSTS70008/50173/700084 etc.) den lokalen Cache leeren,
            // damit der frische Login-Redirect tatsächlich frische Tokens holt.
            if (looksLikeBrokenCache(e)) {
                try {
                    await clearMsalCache(instance);
                } catch {
                    // ignore
                }
            }
            if (isInteractionRequired(e) || looksLikeBrokenCache(e)) {
                try {
                    sessionStorage.setItem(POST_LOGIN_KEY, window.location.href);
                } catch {
                    // ignore
                }
                const redirectReq = { ...req, redirectStartPage: window.location.href };
                // Beim "Cache broken" zusätzlich Consent erzwingen, damit der Tenant
                // den User korrekt neu authentifiziert.
                if (looksLikeBrokenCache(e)) {
                    redirectReq.prompt = 'select_account';
                }
                await instance.acquireTokenRedirect(redirectReq);
                throw new Error('Weiterleitung zur Anmeldung …');
            }
            throw e;
        }
    }

    /** Popup-Anmeldung (ohne Seiten-Redirect) – für Einrichtung und Werkzeuge mit lokalen Formularen. */
    async function acquireTokenPopup(scopes) {
        const instance = await ensurePca();
        const scopeList = Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES;
        let accounts = instance.getAllAccounts();
        if (!accounts.length) {
            await instance.loginPopup({ scopes: scopeList, prompt: 'select_account' });
            accounts = instance.getAllAccounts();
        }
        if (!accounts.length) {
            throw new Error('Anmeldung abgebrochen.');
        }
        const a = getAccount() || accounts[0];
        if (a && typeof instance.setActiveAccount === 'function') {
            instance.setActiveAccount(a);
        }
        const req = { scopes: scopeList, account: a };
        try {
            const token = (await instance.acquireTokenSilent(req)).accessToken;
            setWidgetState();
            return token;
        } catch (e) {
            if (looksLikeBrokenCache(e)) {
                try {
                    await clearMsalCache(instance);
                } catch {
                    // ignore
                }
            }
            if (isInteractionRequired(e) || looksLikeBrokenCache(e)) {
                const token = (await instance.acquireTokenPopup(req)).accessToken;
                setWidgetState();
                return token;
            }
            throw e;
        }
    }

    function createAuthWidget() {
        const wrap = document.createElement('div');
        wrap.id = 'ms365AuthWidget';
        wrap.className = 'ms365-auth-widget';

        const actions = document.createElement('div');
        actions.id = 'ms365AuthActions';
        actions.className = 'ms365-auth-actions';

        const btn = document.createElement('button');
        btn.id = 'ms365AuthBtn';
        btn.type = 'button';
        btn.className = 'btn';
        btn.innerHTML = '<i class="bi bi-box-arrow-in-right"></i>Anmelden';

        const menu = document.createElement('div');
        menu.id = 'ms365AuthMenu';
        menu.className = 'ms365-auth-menu';
        menu.hidden = true;

        const trigger = document.createElement('button');
        trigger.id = 'ms365AuthBadge';
        trigger.type = 'button';
        trigger.className = 'ms365-auth-menu__trigger';
        trigger.setAttribute('aria-haspopup', 'menu');
        trigger.setAttribute('aria-expanded', 'false');
        trigger.setAttribute('aria-controls', 'ms365AuthDropdown');
        trigger.setAttribute('aria-label', 'Konto');
        trigger.innerHTML =
            '<i class="bi bi-person-circle" aria-hidden="true"></i>' +
            '<span id="ms365AuthBadgeText">–</span>' +
            '<i class="bi bi-chevron-down ms365-auth-menu__chevron" aria-hidden="true"></i>';

        const drop = document.createElement('div');
        drop.id = 'ms365AuthDropdown';
        drop.className = 'ms365-auth-menu__panel';
        drop.setAttribute('role', 'menu');
        drop.hidden = true;
        drop.innerHTML =
            '<div class="ms365-auth-menu__meta">' +
            '<div class="ms365-auth-menu__meta-name" id="ms365AuthMenuName"></div>' +
            '<div class="ms365-auth-menu__meta-mail" id="ms365AuthMenuMail"></div>' +
            '</div>' +
            '<div class="ms365-auth-menu__ctx" aria-label="Schulkontext">' +
            '<div class="ms365-auth-menu__ctx-row"><span class="ms365-auth-menu__ctx-k">Schuljahr</span>' +
            '<span class="ms365-auth-menu__ctx-v" id="ms365AuthCtxYear">–</span></div>' +
            '<div class="ms365-auth-menu__ctx-row"><span class="ms365-auth-menu__ctx-k">Domain</span>' +
            '<span class="ms365-auth-menu__ctx-v" id="ms365AuthCtxDomain">–</span></div>' +
            '</div>' +
            '<a class="ms365-auth-menu__item" role="menuitem" id="ms365AuthActionLogLink" href="action-log.html">' +
            '<i class="bi bi-journal-text" aria-hidden="true"></i>Aktionsprotokoll</a>' +
            '<button type="button" class="ms365-auth-menu__item" role="menuitem" id="ms365AuthSwitchBtn" title="Konto wechseln / Anmeldung zurücksetzen">' +
            '<i class="bi bi-arrow-repeat" aria-hidden="true"></i>Konto wechseln</button>' +
            '<button type="button" class="ms365-auth-menu__item ms365-auth-menu__item--danger" role="menuitem" id="ms365AuthLogoutBtn">' +
            '<i class="bi bi-box-arrow-right" aria-hidden="true"></i>Abmelden</button>';

        menu.appendChild(trigger);
        menu.appendChild(drop);
        wrap.appendChild(actions);
        wrap.appendChild(btn);
        wrap.appendChild(menu);
        return wrap;
    }

    function ensureAuthMenuBindings() {
        const trigger = document.getElementById('ms365AuthBadge');
        const actionLogLink = document.getElementById('ms365AuthActionLogLink');
        const switchBtn = document.getElementById('ms365AuthSwitchBtn');
        const logoutBtn = document.getElementById('ms365AuthLogoutBtn');
        if (trigger && !trigger.dataset.bound) {
            trigger.dataset.bound = '1';
            trigger.addEventListener('click', function (e) {
                e.stopPropagation();
                toggleAuthMenu();
            });
        }
        if (switchBtn && !switchBtn.dataset.bound) {
            switchBtn.dataset.bound = '1';
            switchBtn.addEventListener('click', function () {
                closeAuthMenu();
                forceFreshLogin().catch(function () {});
            });
        }
        if (actionLogLink && !actionLogLink.dataset.bound) {
            actionLogLink.dataset.bound = '1';
            actionLogLink.addEventListener('click', function () {
                closeAuthMenu();
            });
        }
        if (logoutBtn && !logoutBtn.dataset.bound) {
            logoutBtn.dataset.bound = '1';
            logoutBtn.addEventListener('click', function () {
                closeAuthMenu();
                logout().catch(function () {});
            });
        }
        bindAuthMenuDismiss();
    }

    function placeAuthWidgetInMenuHeader() {
        const header = $('.header') || $('header');
        if (!header) return false;
        let wrap = $('#ms365AuthWidget');
        if (!wrap || !wrap.querySelector('#ms365AuthMenu')) {
            if (wrap && wrap.parentElement) wrap.parentElement.removeChild(wrap);
            wrap = createAuthWidget();
        }
        try {
            header.style.position = header.style.position || 'relative';
        } catch {
            /* ignore */
        }
        wrap.style.position = 'absolute';
        wrap.style.top = '16px';
        wrap.style.right = '16px';
        wrap.style.zIndex = '6';
        wrap.style.marginLeft = '0';
        wrap.style.flexWrap = 'nowrap';
        if (wrap.parentElement !== header) header.appendChild(wrap);
        ensureAuthMenuBindings();
        if (pca) setWidgetState();
        try {
            if (typeof window.ms365RefreshContextBar === 'function') window.ms365RefreshContextBar();
        } catch {
            /* ignore */
        }
        return true;
    }

    function ensureHeaderWidget() {
        const header = $('.header') || $('header');
        if (!header) return;
        placeAuthWidgetInMenuHeader();
        try {
            window.dispatchEvent(new CustomEvent('ms365-auth-widget-ready'));
        } catch {
            /* ignore */
        }
    }

    async function forceFreshLogin() {
        try {
            const instance = await ensurePca();
            await clearMsalCache(instance);
        } catch {
            // ignore – wir versuchen den Redirect trotzdem
        }
        try {
            return await login(DEFAULT_SCOPES, { prompt: 'select_account' });
        } catch {
            // bei Redirect ohnehin kein weiterer Code mehr
        }
    }

    function setWidgetState() {
        const badgeText = document.getElementById('ms365AuthBadgeText');
        const btn = document.getElementById('ms365AuthBtn');
        const menu = document.getElementById('ms365AuthMenu');
        const trigger = document.getElementById('ms365AuthBadge');
        const menuName = document.getElementById('ms365AuthMenuName');
        const menuMail = document.getElementById('ms365AuthMenuMail');
        const menuMeta = menu && menu.querySelector('.ms365-auth-menu__meta');
        const switchBtn = document.getElementById('ms365AuthSwitchBtn');
        const logoutBtn = document.getElementById('ms365AuthLogoutBtn');
        const a = getAccount();
        const name = accountDisplayName(a);
        const mail = a && a.username ? String(a.username) : '';
        closeAuthMenu();
        ensureAuthMenuBindings();
        if (badgeText) badgeText.textContent = a ? name : 'Konto';
        if (menuName) menuName.textContent = name;
        if (menuMail) {
            menuMail.textContent = mail && mail !== name ? mail : '';
            menuMail.hidden = !(mail && mail !== name);
        }
        if (menuMeta) menuMeta.hidden = !a;
        if (switchBtn) switchBtn.hidden = !a;
        if (logoutBtn) logoutBtn.hidden = !a;
        if (trigger) {
            trigger.setAttribute('aria-label', a ? 'Konto: ' + accountLabel(a) : 'Konto');
            trigger.title = a ? accountLabel(a) : 'Konto';
        }
        if (menu) menu.hidden = false;
        if (btn) {
            if (a) {
                btn.hidden = true;
                btn.onclick = null;
            } else {
                btn.hidden = false;
                btn.setAttribute('aria-label', 'Anmelden');
                btn.title = 'Anmelden';
                btn.innerHTML = '<i class="bi bi-box-arrow-in-right"></i>Anmelden';
                btn.onclick = function () {
                    login(DEFAULT_SCOPES).catch(function () {});
                };
            }
        }
    }

    async function init() {
        if (typeof document === 'undefined') return;
        ensureHeaderWidget();
        try {
            // In case widget existed already, still notify listeners.
            window.dispatchEvent(new CustomEvent('ms365-auth-widget-ready'));
        } catch {
            // ignore
        }
        try {
            await ensurePca();
        } catch {
            // ignore (widget still renders)
        }
        // Wenn lokal noch kein Account im Cache ist, einmalig SSO Silent versuchen.
        // Damit wird die Microsoft-365-Anmeldung übernommen, wenn der Benutzer
        // in einem anderen Tab/Fenster (z. B. Outlook, Teams Web, anderes Modul) bereits
        // angemeldet ist – ohne sichtbaren Redirect.
        try {
            if (pca && !getAccount()) {
                await trySsoSilent(DEFAULT_SCOPES);
            }
        } catch {
            // ignore
        }
        setWidgetState();
    }

    // Public API for tools
    window.ms365AuthEnsureInitialized = ensurePca;
    window.ms365AuthGetActionSlot = function () {
        try {
            return document.getElementById('ms365AuthActions');
        } catch {
            return null;
        }
    };
    window.ms365AuthGetAccountLabel = function () {
        try {
            return accountLabel(getAccount());
        } catch {
            return '';
        }
    };
    window.ms365AuthIsLoggedIn = function () {
        try {
            return !!getAccount();
        } catch {
            return false;
        }
    };
    window.ms365AuthLogin = login;
    window.ms365AuthSwitchAccount = switchAccount;
    window.ms365AuthLogout = logout;
    window.ms365AuthAcquireToken = acquireToken;
    window.ms365AuthAcquireTokenPopup = acquireTokenPopup;
    window.ms365AuthRefreshWidget = setWidgetState;
    window.ms365AuthGetTenantId = async function ms365AuthGetTenantId() {
        try {
            await ensurePca();
            const a = getAccount();
            if (!a) return '';
            return String(a.tenantId || (a.idTokenClaims && a.idTokenClaims.tid) || '').trim();
        } catch {
            return '';
        }
    };
    window.ms365AuthGetUserPrincipalName = function ms365AuthGetUserPrincipalName() {
        try {
            const a = getAccount();
            return a && a.username ? String(a.username).trim() : '';
        } catch {
            return '';
        }
    };

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
    try {
        window.addEventListener('ms365-menu-header-ready', function () {
            placeAuthWidgetInMenuHeader();
        });
    } catch {
        /* ignore */
    }
})();

