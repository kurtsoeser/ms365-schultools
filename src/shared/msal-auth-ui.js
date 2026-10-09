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

    function resolvePreferredLoginScopes() {
        try {
            const pageScopes = window.MS365_AUTH_LOGIN_SCOPES;
            if (Array.isArray(pageScopes) && pageScopes.length) {
                return pageScopes.slice();
            }
        } catch {
            /* ignore */
        }
        try {
            const g = window.ms365GraphUnifiedGroups;
            if (g && Array.isArray(g.GRAPH_SCOPES) && g.GRAPH_SCOPES.length) {
                return g.GRAPH_SCOPES.slice();
            }
        } catch {
            /* ignore */
        }
        return DEFAULT_SCOPES.slice();
    }
    const POST_LOGIN_KEY = 'ms365-post-login-url';

    function rememberPostLoginReturnUrl(url) {
        const target = String(url || (window.location && window.location.href) || '').trim();
        if (!target) return;
        try {
            sessionStorage.setItem(POST_LOGIN_KEY, target);
        } catch {
            // ignore
        }
    }

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
        if (pca) return pca;
        if (initPromise) return initPromise;
        initPromise = (async () => {
            try {
                const m = await loadMsal();
                const PublicClientApplication =
                    m.PublicClientApplication || (m.default && m.default.PublicClientApplication);
                if (!PublicClientApplication) throw new Error('MSAL: PublicClientApplication nicht gefunden.');
                const cfg = resolveMsalConfig();
                const instance = new PublicClientApplication({
                    auth: { clientId: cfg.clientId, authority: cfg.authority, redirectUri: cfg.redirectUri },
                    // localStorage statt sessionStorage: ermöglicht Single-Sign-On zwischen Browser-Tabs
                    // (Microsoft 365 Anmeldung wird übernommen, wenn der Benutzer bereits in einem
                    // anderen Tab/Modul angemeldet ist).
                    cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true }
                });
                await withTimeout(instance.initialize(), 8000);
                try {
                    await withTimeout(instance.handleRedirectPromise(), 8000);
                } catch {
                    // Redirect-Handling darf Login nicht dauerhaft blockieren
                }

                const accounts = instance.getAllAccounts();
                if (accounts && accounts[0] && typeof instance.setActiveAccount === 'function') {
                    instance.setActiveAccount(accounts[0]);
                }
                pca = instance;
                return pca;
            } catch (e) {
                initPromise = null;
                throw e;
            }
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
    function withTimeout(promise, ms) {
        return new Promise(function (resolve, reject) {
            var settled = false;
            var t = setTimeout(function () {
                if (settled) return;
                settled = true;
                reject(new Error('timeout'));
            }, ms);
            Promise.resolve(promise).then(
                function (v) {
                    if (settled) return;
                    settled = true;
                    clearTimeout(t);
                    resolve(v);
                },
                function (e) {
                    if (settled) return;
                    settled = true;
                    clearTimeout(t);
                    reject(e);
                }
            );
        });
    }

    async function trySsoSilent(scopes) {
        if (!pca) return null;
        try {
            const req = {
                scopes: Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES
            };
            // ssoSilent kann bei blockierten 3rd-Party-Cookies lange hängen –
            // dann wäre der Anmelden-Button ohne Handler „tot“.
            const result = await withTimeout(pca.ssoSilent(req), 4000);
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

    function closeBackupPanel() {
        const wrap = document.getElementById('ms365BackupHeader');
        const trigger = document.getElementById('ms365BackupHeaderBtn');
        const panel = document.getElementById('ms365BackupPanel');
        if (wrap) wrap.classList.remove('is-open');
        if (trigger) trigger.setAttribute('aria-expanded', 'false');
        if (panel) panel.hidden = true;
    }

    function toggleAuthMenu() {
        const menu = document.getElementById('ms365AuthMenu');
        const trigger = document.getElementById('ms365AuthBadge');
        const drop = document.getElementById('ms365AuthDropdown');
        if (!menu || !trigger || !drop || menu.hidden) return;
        const open = !menu.classList.contains('is-open');
        if (open) closeBackupPanel();
        menu.classList.toggle('is-open', open);
        trigger.setAttribute('aria-expanded', open ? 'true' : 'false');
        drop.hidden = !open;
        if (
            open &&
            window.ms365OperatorAccess &&
            typeof window.ms365OperatorAccess.refreshOperatorStatus === 'function'
        ) {
            window.ms365OperatorAccess.refreshOperatorStatus({ force: false }).then(function () {
                const adminLink = document.getElementById('ms365AuthAdminLink');
                if (!adminLink) return;
                const show =
                    typeof window.ms365OperatorAccess.shouldShowAdminMenuLink === 'function'
                        ? window.ms365OperatorAccess.shouldShowAdminMenuLink()
                        : window.ms365OperatorAccess.isCurrentUserOperator();
                adminLink.hidden = !show;
            });
        }
    }

    function toggleBackupPanel() {
        const wrap = document.getElementById('ms365BackupHeader');
        const trigger = document.getElementById('ms365BackupHeaderBtn');
        const panel = document.getElementById('ms365BackupPanel');
        if (!wrap || !trigger || !panel) return;
        const open = !wrap.classList.contains('is-open');
        if (open) closeAuthMenu();
        wrap.classList.toggle('is-open', open);
        trigger.setAttribute('aria-expanded', open ? 'true' : 'false');
        panel.hidden = !open;
        if (open) refreshBackupHeaderStatus();
    }

    const MENU_BRANDS = ['teal', 'classic', 'wine'];

    function normalizeMenuBrand(brand) {
        return MENU_BRANDS.indexOf(brand) !== -1 ? brand : 'teal';
    }

    function applyMenuBrandChoice(brand) {
        const next = normalizeMenuBrand(brand);
        document.documentElement.setAttribute('data-brand', next);
        try {
            localStorage.setItem('ms365-brand-v1', next);
        } catch (_) {
            /* ignore */
        }
        if (window.ms365Theme && typeof window.ms365Theme.setBrand === 'function') {
            window.ms365Theme.setBrand(next);
        }
        if (document.documentElement.getAttribute('data-brand') !== next) {
            document.documentElement.setAttribute('data-brand', next);
            try {
                localStorage.setItem('ms365-brand-v1', next);
            } catch (_) {
                /* ignore */
            }
        }
        if (window.ms365Theme && typeof window.ms365Theme.syncBrandUi === 'function') {
            window.ms365Theme.syncBrandUi();
        } else {
            document.querySelectorAll('[data-ms365-brand]').forEach(function (el) {
                const on = el.getAttribute('data-ms365-brand') === next;
                el.setAttribute('aria-checked', on ? 'true' : 'false');
                el.classList.toggle('is-active', on);
            });
        }
    }

    function bindAuthMenuDismiss() {
        if (bindAuthMenuDismiss.bound) return;
        bindAuthMenuDismiss.bound = true;
        document.addEventListener('click', function (e) {
            const menu = document.getElementById('ms365AuthMenu');
            if (menu && menu.classList.contains('is-open') && !menu.contains(e.target)) {
                closeAuthMenu();
            }
            const backup = document.getElementById('ms365BackupHeader');
            if (backup && backup.classList.contains('is-open') && !backup.contains(e.target)) {
                closeBackupPanel();
            }
        });
        document.addEventListener('keydown', function (e) {
            if (e.key === 'Escape') {
                closeAuthMenu();
                closeBackupPanel();
            }
        });
    }

    /**
     * Anmeldung. Lokal bevorzugt Popup (zuverlässiger mit Vite-Ports),
     * sonst Redirect. Fehler werden sichtbar gemeldet.
     */
    async function login(scopes, opts) {
        const instance = await ensurePca();
        rememberPostLoginReturnUrl();
        const req = {
            scopes: Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES
        };
        if (opts && typeof opts.prompt === 'string' && opts.prompt) {
            req.prompt = opts.prompt;
        }
        const host = String((window.location && window.location.hostname) || '').toLowerCase();
        const isLocal =
            host === 'localhost' || host === '127.0.0.1' || host === '::1' || host.endsWith('.localhost');
        const forcePopup = !!(opts && opts.popup) || isLocal;
        if (forcePopup && typeof instance.loginPopup === 'function') {
            try {
                const result = await instance.loginPopup(req);
                if (result && result.account && typeof instance.setActiveAccount === 'function') {
                    instance.setActiveAccount(result.account);
                }
                setWidgetState();
                return result;
            } catch (e) {
                if (isPopupWindowError(e) && !isUserCancelledAuth(e)) {
                    req.redirectStartPage = window.location.href;
                    await instance.loginRedirect(req);
                    throw new Error('Weiterleitung zur Anmeldung …');
                }
                throw e;
            }
        }
        req.redirectStartPage = window.location.href;
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
        try {
            if (window.ms365LicenseGateCore && typeof window.ms365LicenseGateCore.clearCache === 'function') {
                window.ms365LicenseGateCore.clearCache();
            } else {
                sessionStorage.removeItem('ms365-license-me-v1');
            }
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

    function isUserCancelledAuth(e) {
        if (!e) return false;
        const msg = String((e && e.message) || e || '');
        const code = String((e && (e.errorCode || e.code)) || '');
        return (
            /abgebrochen|cancelled|canceled|user_cancelled/i.test(msg) ||
            /user_cancelled/i.test(code)
        );
    }

    function isPopupWindowError(e) {
        if (!e) return false;
        const code = String((e.errorCode || e.code) || '');
        const msg = String((e && e.message) || e || '');
        return (
            code === 'popup_window_error' ||
            /popup_window_error/i.test(msg) ||
            /error opening popup/i.test(msg)
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
                rememberPostLoginReturnUrl();
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

    /**
     * Nur Silent – kein Redirect/Popup. Für Hintergrundprüfungen (z. B. Admin-Menü).
     * @param {string[]} [scopes]
     * @returns {Promise<string>}
     */
    async function acquireTokenSilentOnly(scopes) {
        const instance = await ensurePca();
        const accounts = instance.getAllAccounts();
        if (!accounts.length) {
            throw new Error('Nicht angemeldet.');
        }
        const a = getAccount() || accounts[0];
        const req = {
            scopes: Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES,
            account: a
        };
        return (await instance.acquireTokenSilent(req)).accessToken;
    }

    /** Popup-Anmeldung (ohne Seiten-Redirect) – für Einrichtung und Werkzeuge mit lokalen Formularen. */
    async function acquireTokenPopup(scopes) {
        const instance = await ensurePca();
        const scopeList = Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES;
        let accounts = instance.getAllAccounts();
        if (!accounts.length) {
            try {
                await instance.loginPopup({ scopes: scopeList, prompt: 'select_account' });
            } catch (e) {
                if (isPopupWindowError(e) && !isUserCancelledAuth(e)) {
                    return acquireToken(scopeList);
                }
                throw e;
            }
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
            const result = await instance.acquireTokenSilent(req);
            setWidgetState({ silent: true });
            return result.accessToken;
        } catch (e) {
            if (looksLikeBrokenCache(e)) {
                try {
                    await clearMsalCache(instance);
                } catch {
                    // ignore
                }
            }
            if (isInteractionRequired(e) || looksLikeBrokenCache(e)) {
                try {
                    const result = await instance.acquireTokenPopup(req);
                    setWidgetState({ silent: true });
                    return result.accessToken;
                } catch (pe) {
                    if (isPopupWindowError(pe) && !isUserCancelledAuth(pe)) {
                        return acquireToken(scopeList);
                    }
                    throw pe;
                }
            }
            throw e;
        }
    }

    /**
     * ID-Token still (für License-API). Fallback: Access Token.
     * @param {string[]} [scopes]
     */
    async function acquireIdToken(scopes) {
        const instance = await ensurePca();
        let accounts = instance.getAllAccounts();
        if (!accounts.length) {
            await login(scopes);
            throw new Error('Weiterleitung zur Anmeldung …');
        }
        const a = getAccount() || accounts[0];
        const req = {
            scopes: Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES,
            account: a
        };
        try {
            const result = await instance.acquireTokenSilent(req);
            return result.idToken || result.accessToken;
        } catch (e) {
            if (looksLikeBrokenCache(e)) {
                try {
                    await clearMsalCache(instance);
                } catch {
                    // ignore
                }
            }
            if (isInteractionRequired(e) || looksLikeBrokenCache(e)) {
                rememberPostLoginReturnUrl();
                const redirectReq = { ...req, redirectStartPage: window.location.href };
                if (looksLikeBrokenCache(e)) {
                    redirectReq.prompt = 'select_account';
                }
                await instance.acquireTokenRedirect(redirectReq);
                throw new Error('Weiterleitung zur Anmeldung …');
            }
            throw e;
        }
    }

    /**
     * ID-Token für eigene Backends (z. B. License-API). Fallback: Access Token.
     * @param {string[]} [scopes]
     */
    async function acquireIdTokenPopup(scopes) {
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
            const result = await instance.acquireTokenSilent(req);
            setWidgetState({ silent: true });
            return result.idToken || result.accessToken;
        } catch (e) {
            if (looksLikeBrokenCache(e)) {
                try {
                    await clearMsalCache(instance);
                } catch {
                    // ignore
                }
            }
            if (isInteractionRequired(e) || looksLikeBrokenCache(e)) {
                const result = await instance.acquireTokenPopup(req);
                setWidgetState({ silent: true });
                return result.idToken || result.accessToken;
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
        drop.setAttribute('data-ms365-auth-menu-v', '2');
        drop.innerHTML =
            '<div class="ms365-auth-menu__meta">' +
            '<div class="ms365-auth-menu__meta-name" id="ms365AuthMenuName"></div>' +
            '<div class="ms365-auth-menu__meta-mail" id="ms365AuthMenuMail"></div>' +
            '</div>' +
            '<div class="ms365-auth-menu__section ms365-auth-menu__section--ctx" aria-label="Kontext">' +
            '<div class="ms365-auth-menu__section-label">Kontext</div>' +
            '<div class="ms365-auth-menu__ctx">' +
            '<div class="ms365-auth-menu__ctx-row ms365-auth-menu__ctx-row--year">' +
            '<span class="ms365-auth-menu__ctx-k">Schuljahr</span>' +
            '<div class="ms365-auth-menu__year-wrap">' +
            '<select class="ms365-auth-menu__year-select" id="schoolYearSelect" aria-label="Aktives Schuljahr"></select>' +
            '<button type="button" class="ms365-auth-menu__year-add" id="schoolYearAddBtn" title="Weiteres Schuljahr anlegen" aria-label="Neues Schuljahr anlegen">' +
            '<i class="bi bi-plus-lg" aria-hidden="true"></i></button>' +
            '</div></div>' +
            '<div class="ms365-auth-menu__ctx-row"><span class="ms365-auth-menu__ctx-k">Domain</span>' +
            '<span class="ms365-auth-menu__ctx-v" id="ms365AuthCtxDomain">–</span></div>' +
            '</div></div>' +
            '<div class="ms365-auth-menu__section" id="ms365AuthDashViewSection" role="group" aria-label="Ansicht (Rollen-Vorschau)" hidden>' +
            '<div class="ms365-auth-menu__section-label" id="ms365AuthDashViewLabel">Ansicht (Rollen-Vorschau)</div>' +
            '<div class="ms365-auth-menu__dash-seg" role="group" aria-labelledby="ms365AuthDashViewLabel">' +
            '<button type="button" class="ms365-auth-menu__dash-seg-btn" role="menuitemradio" data-dash-layout="full" aria-checked="false" title="Vollzugriff (Schul-IT)" aria-label="Vollzugriff Schul-IT">' +
            '<i class="bi bi-grid-3x3-gap-fill" aria-hidden="true"></i><span class="ms365-auth-menu__dash-seg-k">IT</span></button>' +
            '<button type="button" class="ms365-auth-menu__dash-seg-btn" role="menuitemradio" data-dash-layout="preview-lehrer" aria-checked="false" title="Vorschau Lehrkraft" aria-label="Vorschau Lehrkraft">' +
            '<i class="bi bi-person-badge" aria-hidden="true"></i><span class="ms365-auth-menu__dash-seg-k">Lehrer</span></button>' +
            '<button type="button" class="ms365-auth-menu__dash-seg-btn" role="menuitemradio" data-dash-layout="preview-schueler" aria-checked="false" title="Vorschau Schüler/in" aria-label="Vorschau Schüler/in">' +
            '<i class="bi bi-mortarboard" aria-hidden="true"></i><span class="ms365-auth-menu__dash-seg-k">Schüler</span></button>' +
            '</div>' +
            '<p class="ms365-auth-menu__dash-hint" id="ms365AuthDashViewHint" hidden></p>' +
            '</div>' +
            '<div class="ms365-auth-menu__section" role="group" aria-label="Design">' +
            '<div class="ms365-auth-menu__section-label">Design</div>' +
            '<div class="ms365-auth-menu__brand-row">' +
            '<button type="button" class="ms365-auth-menu__brand" role="menuitemradio" data-ms365-brand="teal" aria-checked="true">' +
            '<span class="ms365-auth-menu__brand-swatch ms365-auth-menu__brand-swatch--teal" aria-hidden="true"></span>Blau-Grün</button>' +
            '<button type="button" class="ms365-auth-menu__brand" role="menuitemradio" data-ms365-brand="classic" aria-checked="false">' +
            '<span class="ms365-auth-menu__brand-swatch ms365-auth-menu__brand-swatch--classic" aria-hidden="true"></span>Klassisch</button>' +
            '<button type="button" class="ms365-auth-menu__brand ms365-auth-menu__theme" data-ms365-auth-theme-toggle="1" aria-label="Hell- und Dunkelmodus umschalten">' +
            '<span class="ms365-auth-menu__brand-swatch ms365-auth-menu__brand-swatch--dark" aria-hidden="true"></span>Dunkel</button>' +
            '</div></div>' +
            '<div class="ms365-auth-menu__section ms365-auth-menu__section--links" role="group" aria-label="Verwaltung">' +
            '<a class="ms365-auth-menu__item" role="menuitem" id="ms365AuthAdminLink" href="admin.html" hidden>' +
            '<i class="bi bi-shield-lock" aria-hidden="true"></i>Admin-Bereich</a>' +
            '<a class="ms365-auth-menu__item" role="menuitem" id="ms365AuthActionLogLink" href="action-log.html">' +
            '<i class="bi bi-journal-text" aria-hidden="true"></i>Aktionsprotokoll</a>' +
            '</div>' +
            '<div class="ms365-auth-menu__section ms365-auth-menu__section--account" role="group" aria-label="Konto">' +
            '<button type="button" class="ms365-auth-menu__item" role="menuitem" id="ms365AuthSwitchBtn" title="Konto wechseln / Anmeldung zurücksetzen">' +
            '<i class="bi bi-arrow-repeat" aria-hidden="true"></i>Konto wechseln</button>' +
            '<button type="button" class="ms365-auth-menu__item ms365-auth-menu__item--danger" role="menuitem" id="ms365AuthLogoutBtn">' +
            '<i class="bi bi-box-arrow-right" aria-hidden="true"></i>Abmelden</button>' +
            '</div>';

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
        const adminLink = document.getElementById('ms365AuthAdminLink');
        const switchBtn = document.getElementById('ms365AuthSwitchBtn');
        const logoutBtn = document.getElementById('ms365AuthLogoutBtn');
        if (trigger && !trigger.dataset.bound) {
            trigger.dataset.bound = '1';
            trigger.addEventListener('click', function (e) {
                e.stopPropagation();
                toggleAuthMenu();
            });
        }
        if (adminLink && !adminLink.dataset.bound) {
            adminLink.dataset.bound = '1';
            adminLink.addEventListener('click', function (e) {
                e.preventDefault();
                closeAuthMenu();
                if (window.ms365OperatorAccess && typeof window.ms365OperatorAccess.openAdminArea === 'function') {
                    window.ms365OperatorAccess.openAdminArea();
                } else {
                    location.href = 'admin.html';
                }
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
        const themeBtn = document.querySelector('#ms365AuthDropdown [data-ms365-auth-theme-toggle]');
        if (themeBtn && !themeBtn.dataset.bound) {
            themeBtn.dataset.bound = '1';
            themeBtn.addEventListener('click', function (e) {
                e.preventDefault();
                e.stopPropagation();
                if (window.ms365Theme && typeof window.ms365Theme.toggle === 'function') {
                    window.ms365Theme.toggle();
                }
            });
        }
        const brandBtns = document.querySelectorAll('#ms365AuthDropdown [data-ms365-brand]');
        brandBtns.forEach(function (btn) {
            if (btn.dataset.bound) return;
            btn.dataset.bound = '1';
            btn.addEventListener('click', function (e) {
                e.preventDefault();
                e.stopPropagation();
                applyMenuBrandChoice(btn.getAttribute('data-ms365-brand'));
            });
        });
        if (window.ms365Theme && typeof window.ms365Theme.getBrand === 'function') {
            const cur = window.ms365Theme.getBrand();
            document.querySelectorAll('[data-ms365-brand]').forEach(function (el) {
                const on = el.getAttribute('data-ms365-brand') === cur;
                el.setAttribute('aria-checked', on ? 'true' : 'false');
                el.classList.toggle('is-active', on);
            });
        }

        bindAuthMenuDismiss();
    }

    function resolveAppRootHref(file) {
        try {
            if (
                window.ms365OperatorAccess &&
                typeof window.ms365OperatorAccess.resolveAppRootHref === 'function'
            ) {
                return window.ms365OperatorAccess.resolveAppRootHref(file);
            }
        } catch {
            /* ignore */
        }
        try {
            const p = String(window.location.pathname || '/').split('?')[0].split('#')[0];
            const norm = p.replace(/\\/g, '/');
            const iTools = norm.toLowerCase().indexOf('/tools/');
            if (iTools !== -1) return norm.slice(0, iTools) + '/' + String(file || '').replace(/^\//, '');
            const slash = norm.lastIndexOf('/');
            const dir = slash >= 0 ? norm.slice(0, slash + 1) : '/';
            return dir + String(file || '').replace(/^\//, '');
        } catch {
            return String(file || '');
        }
    }

    function formatBackupWhenDe(iso) {
        const ts = Date.parse(iso);
        if (!iso || isNaN(ts)) return '';
        const d = new Date(ts);
        const now = new Date();
        const sameDay =
            d.getFullYear() === now.getFullYear() &&
            d.getMonth() === now.getMonth() &&
            d.getDate() === now.getDate();
        const hh = String(d.getHours()).padStart(2, '0');
        const mm = String(d.getMinutes()).padStart(2, '0');
        const time = hh + ':' + mm;
        if (sameDay) return 'heute ' + time;
        const yday = new Date(now);
        yday.setDate(yday.getDate() - 1);
        const isYday =
            d.getFullYear() === yday.getFullYear() &&
            d.getMonth() === yday.getMonth() &&
            d.getDate() === yday.getDate();
        if (isYday) return 'gestern ' + time;
        const dd = String(d.getDate()).padStart(2, '0');
        const mo = String(d.getMonth() + 1).padStart(2, '0');
        return dd + '.' + mo + '.' + d.getFullYear() + ' ' + time;
    }

    function readSpoSyncAtTenantAware() {
        try {
            if (
                window.ms365StammdatenSpoAutoSync &&
                typeof window.ms365StammdatenSpoAutoSync.getStatus === 'function'
            ) {
                const st = window.ms365StammdatenSpoAutoSync.getStatus();
                if (st && st.lastAt) return String(st.lastAt);
            }
        } catch {
            /* ignore */
        }
        try {
            let tid = '';
            if (typeof window.ms365AuthGetAccountInfo === 'function') {
                const info = window.ms365AuthGetAccountInfo();
                tid = info && info.tenantId ? String(info.tenantId).trim() : '';
            }
            if (tid) {
                const map = JSON.parse(
                    localStorage.getItem('ms365-stammdaten-spo-sync-by-tenant-v2') || '{}'
                );
                if (map && map[tid] && map[tid].at) return String(map[tid].at);
            }
        } catch {
            /* ignore */
        }
        try {
            const m = JSON.parse(localStorage.getItem('ms365-stammdaten-spo-sync-v1') || '{}') || {};
            return m.at ? String(m.at) : '';
        } catch {
            return '';
        }
    }

    function readLastBackupAt() {
        let browserAt = '';
        let spoAt = '';
        try {
            browserAt = localStorage.getItem('ms365-last-backup-export-at') || '';
        } catch {
            /* ignore */
        }
        spoAt = readSpoSyncAtTenantAware();
        const bTs = Date.parse(browserAt);
        const sTs = Date.parse(spoAt);
        if (!isNaN(bTs) && !isNaN(sTs)) return bTs >= sTs ? browserAt : spoAt;
        if (!isNaN(bTs)) return browserAt;
        if (!isNaN(sTs)) return spoAt;
        return '';
    }

    function refreshBackupHeaderStatus() {
        const statusEl = document.getElementById('ms365BackupPanelStatus');
        const trigger = document.getElementById('ms365BackupHeaderBtn');
        const last = readLastBackupAt();
        const ts = Date.parse(last);
        const ageHours = !isNaN(ts) ? (Date.now() - ts) / (1000 * 3600) : Infinity;
        const fresh = ageHours < 24;
        const aging = ageHours >= 24 && ageHours < 72;
        if (statusEl) {
            if (!last || isNaN(ts)) {
                statusEl.textContent = 'Noch nicht gesichert';
                statusEl.className = 'ms365-backup-panel__status ms365-backup-panel__status--warn';
            } else {
                const when = formatBackupWhenDe(last);
                statusEl.textContent =
                    'Zuletzt gesichert: ' + when + (fresh ? ' ✅' : ' – bitte erneuern');
                statusEl.className =
                    'ms365-backup-panel__status' +
                    (fresh ? ' ms365-backup-panel__status--ok' : ' ms365-backup-panel__status--warn');
            }
        }
        const dot = trigger ? trigger.querySelector('.ms365-backup-header__dot') : null;
        if (dot) {
            dot.classList.remove('ms365-backup-header__dot--ok', 'ms365-backup-header__dot--warn', 'ms365-backup-header__dot--stale');
            if (fresh) dot.classList.add('ms365-backup-header__dot--ok');
            else if (aging) dot.classList.add('ms365-backup-header__dot--warn');
            else dot.classList.add('ms365-backup-header__dot--stale');
        }
        if (trigger) {
            trigger.classList.toggle('ms365-backup-header__btn--ok', !!fresh);
            trigger.classList.toggle('ms365-backup-header__btn--stale', !fresh && !aging);
            trigger.classList.toggle('ms365-backup-header__btn--aging', !!aging);
            trigger.title = fresh
                ? 'Datensicherung – zuletzt: ' + formatBackupWhenDe(last)
                : 'Datensicherung – bitte sichern';
            trigger.setAttribute(
                'aria-label',
                fresh ? 'Datensicherung (aktuell)' : 'Datensicherung (nicht aktuell)'
            );
        }
    }

    function bindBackupPanelActions() {
        try {
            if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.bindUi === 'function') {
                window.ms365BrowserBackup.bindUi();
            }
        } catch {
            /* ignore */
        }
        try {
            if (window.ms365StammdatenSpoSyncUi && typeof window.ms365StammdatenSpoSyncUi.bind === 'function') {
                window.ms365StammdatenSpoSyncUi.bind();
            }
        } catch {
            /* ignore */
        }
        refreshBackupHeaderStatus();
    }

    function ensureBackupHeaderControl(container) {
        if (!container) return null;
        let wrap = document.getElementById('ms365BackupHeader');
        if (!wrap) {
            wrap = document.createElement('div');
            wrap.id = 'ms365BackupHeader';
            wrap.className = 'ms365-backup-header';
            wrap.innerHTML =
                '<button type="button" class="ms365-backup-header__btn" id="ms365BackupHeaderBtn" ' +
                'aria-haspopup="dialog" aria-expanded="false" aria-controls="ms365BackupPanel" ' +
                'title="Datensicherung" aria-label="Datensicherung">' +
                '<i class="bi bi-floppy" aria-hidden="true"></i>' +
                '<span class="ms365-backup-header__dot" aria-hidden="true"></span>' +
                '</button>' +
                '<div class="ms365-backup-panel" id="ms365BackupPanel" role="dialog" aria-label="Datensicherung" hidden>' +
                '<div class="ms365-backup-panel__head">' +
                '<i class="bi bi-floppy" aria-hidden="true"></i>' +
                '<strong>Datensicherung</strong>' +
                '</div>' +
                '<div class="ms365-backup-panel__block">' +
                '<div class="ms365-backup-panel__label">Browser-Backup</div>' +
                '<div class="ms365-backup-panel__actions">' +
                '<button type="button" class="ms365-backup-panel__btn" data-ms365-backup="export">' +
                '<i class="bi bi-upload" aria-hidden="true"></i>Exportieren</button>' +
                '<label class="ms365-backup-panel__btn" for="ms365BackupPanelImportFile">' +
                '<i class="bi bi-download" aria-hidden="true"></i>Importieren</label>' +
                '<input type="file" id="ms365BackupPanelImportFile" class="ms365-backup-panel__file" ' +
                'data-ms365-backup="import-file" accept="application/json,.json" hidden>' +
                '</div></div>' +
                '<div class="ms365-backup-panel__block">' +
                '<div class="ms365-backup-panel__label">SharePoint (IT)</div>' +
                '<div class="ms365-backup-panel__actions">' +
                '<button type="button" class="ms365-backup-panel__btn" data-ms365-spo-sync="upload">' +
                '<i class="bi bi-cloud-arrow-up" aria-hidden="true"></i>Sichern</button>' +
                '<button type="button" class="ms365-backup-panel__btn" data-ms365-spo-sync="load">' +
                '<i class="bi bi-cloud-arrow-down" aria-hidden="true"></i>Laden</button>' +
                '</div></div>' +
                '<div class="ms365-backup-panel__block">' +
                '<div class="ms365-backup-panel__label">Abgleich</div>' +
                '<div class="ms365-backup-panel__actions">' +
                '<a class="ms365-backup-panel__btn" href="' +
                resolveAppRootHref('tools/stammdaten-backup-abgleich.html?run=1') +
                '">' +
                '<i class="bi bi-columns-gap" aria-hidden="true"></i>Lokal vs. IT prüfen</a>' +
                '</div></div>' +
                '<div class="ms365-backup-panel__foot">' +
                '<a class="ms365-backup-panel__setup" href="' +
                resolveAppRootHref('tools/stammdaten-uebergabe.html#setup') +
                '" data-ms365-spo-sync="setup">' +
                '<i class="bi bi-gear" aria-hidden="true"></i>IT-Sicherungsbibliothek einrichten</a>' +
                '<p class="ms365-backup-panel__status" id="ms365BackupPanelStatus" role="status"></p>' +
                '</div></div>';
        } else {
            const setup = wrap.querySelector('[data-ms365-spo-sync="setup"]');
            if (setup) setup.setAttribute('href', resolveAppRootHref('tools/stammdaten-uebergabe.html#setup'));
            const compare = wrap.querySelector('a[href*="stammdaten-backup-abgleich"]');
            if (compare) {
                compare.setAttribute('href', resolveAppRootHref('tools/stammdaten-backup-abgleich.html?run=1'));
            }
        }

        if (
            document.body &&
            (document.body.classList.contains('page-dashboard') ||
                document.body.classList.contains('app-shell-chrome'))
        ) {
            wrap.classList.add('ms365-backup-header--dash');
        }

        const schulregister = document.getElementById('ms365HeaderSchulregister');
        const widget = document.getElementById('ms365AuthWidget');
        const before =
            schulregister && schulregister.parentElement === container
                ? schulregister
                : widget && widget.parentElement === container
                  ? widget
                  : null;
        if (before) {
            if (wrap.parentElement !== container || wrap.nextElementSibling !== before) {
                container.insertBefore(wrap, before);
            }
        } else if (wrap.parentElement !== container) {
            container.appendChild(wrap);
        }

        const btn = document.getElementById('ms365BackupHeaderBtn');
        if (btn && !btn.dataset.bound) {
            btn.dataset.bound = '1';
            btn.addEventListener('click', function (e) {
                e.stopPropagation();
                toggleBackupPanel();
            });
        }
        bindAuthMenuDismiss();
        bindBackupPanelActions();
        return wrap;
    }

    function resolveTenantPageHref() {
        try {
            const p = String(window.location.pathname || '/').split('?')[0].split('#')[0];
            const norm = p.replace(/\\/g, '/');
            const iTools = norm.toLowerCase().indexOf('/tools/');
            if (iTools !== -1) return norm.slice(0, iTools) + '/tenant.html';
            const slash = norm.lastIndexOf('/');
            const dir = slash >= 0 ? norm.slice(0, slash + 1) : '/';
            return dir + 'tenant.html';
        } catch {
            return 'tenant.html';
        }
    }

    function isOnTenantRegisterPage() {
        return /\/tenant\.html(?:\?|#|$)/i.test(String(window.location.pathname || '').replace(/\\/g, '/'));
    }

    function ensureSchulregisterHeaderLink(container) {
        if (!container) return null;
        let link = document.getElementById('ms365HeaderSchulregister');
        const hide = isOnTenantRegisterPage();
        if (!link) {
            link = document.createElement('a');
            link.id = 'ms365HeaderSchulregister';
            link.className = 'ms365-header-schulregister';
            link.href = resolveTenantPageHref();
            link.title = 'Stammdaten pflegen';
            link.innerHTML =
                '<i class="bi bi-journal-bookmark" aria-hidden="true"></i>' +
                '<span class="ms365-header-schulregister__label">Stammdaten</span>';
        } else {
            link.href = resolveTenantPageHref();
        }
        link.hidden = hide;
        if (hide) return null;
        if (
            document.body &&
            (document.body.classList.contains('page-dashboard') ||
                document.body.classList.contains('app-shell-chrome'))
        ) {
            link.classList.add('ms365-header-schulregister--dash');
        }
        if (typeof window.ms365SyncFrontendPlannerHeaderChrome === 'function') {
            window.ms365SyncFrontendPlannerHeaderChrome();
        }
        const widget = document.getElementById('ms365AuthWidget');
        const before = widget && widget.parentElement === container ? widget : null;
        if (before) {
            if (link.parentElement !== container || link.nextElementSibling !== before) {
                container.insertBefore(link, before);
            }
        } else if (link.parentElement !== container) {
            container.appendChild(link);
        }
        return link;
    }

    function placeAuthWidgetInMenuHeader() {
        const adminSlot = document.getElementById('adminAppTopActions');
        const header = adminSlot || $('.header') || $('header');
        if (!header) return false;
        let wrap = $('#ms365AuthWidget');
        const menuV =
            wrap &&
            wrap.querySelector('#ms365AuthDropdown') &&
            wrap.querySelector('#ms365AuthDropdown').getAttribute('data-ms365-auth-menu-v');
        if (!wrap || !wrap.querySelector('#ms365AuthMenu') || menuV !== '2') {
            if (wrap && wrap.parentElement) wrap.parentElement.removeChild(wrap);
            wrap = createAuthWidget();
        }
        if (adminSlot) {
            adminSlot.hidden = false;
            wrap.style.position = '';
            wrap.style.top = '';
            wrap.style.right = '';
            wrap.style.zIndex = '';
            wrap.style.marginLeft = '0';
            wrap.style.flexWrap = 'nowrap';
            if (wrap.parentElement !== adminSlot) adminSlot.appendChild(wrap);
            ensureSchulregisterHeaderLink(adminSlot);
            ensureBackupHeaderControl(adminSlot);
        } else {
            try {
                header.style.position = header.style.position || 'relative';
            } catch {
                /* ignore */
            }
            let cluster = document.getElementById('ms365HeaderRightCluster');
            if (!cluster) {
                cluster = document.createElement('div');
                cluster.id = 'ms365HeaderRightCluster';
                cluster.className = 'ms365-header-right-cluster';
                header.appendChild(cluster);
            }
            wrap.style.position = '';
            wrap.style.top = '';
            wrap.style.right = '';
            wrap.style.zIndex = '';
            wrap.style.marginLeft = '0';
            wrap.style.flexWrap = 'nowrap';
            if (wrap.parentElement !== cluster) cluster.appendChild(wrap);
            ensureSchulregisterHeaderLink(cluster);
            ensureBackupHeaderControl(cluster);
        }
        ensureAuthMenuBindings();
        import('./app-header-chrome.js')
            .then(function (m) {
                if (m && typeof m.normalizeAppHeaderActionsOrder === 'function') {
                    m.normalizeAppHeaderActionsOrder(document.getElementById('adminAppTopActions'));
                }
            })
            .catch(function () {
                /* ignore */
            });
        // UI sofort, ohne Auth-Event-Sturm beim Platzieren
        setWidgetState({ silent: true });
        try {
            if (typeof window.ms365RefreshContextBar === 'function') window.ms365RefreshContextBar();
        } catch {
            /* ignore */
        }
        if (typeof window.ms365SyncFrontendPlannerHeaderChrome === 'function') {
            window.ms365SyncFrontendPlannerHeaderChrome();
        }
        return true;
    }

    function ensureHeaderWidget() {
        const header = document.getElementById('adminAppTopActions') || $('.header') || $('header');
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

    function setWidgetState(opts) {
        if (setWidgetState._busy) return;
        setWidgetState._busy = true;
        try {
            setWidgetStateImpl(opts || {});
        } finally {
            setWidgetState._busy = false;
        }
    }

    function setWidgetStateImpl(opts) {
        const silent = !!(opts && opts.silent);
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
        const loggedIn = !!a;
        const prevLoggedIn = setWidgetState._lastLoggedIn;
        const prevLabel = setWidgetState._lastLabel || '';
        const label = a ? accountLabel(a) : '';
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
        const adminLink = document.getElementById('ms365AuthAdminLink');
        const applyAdminLink = function () {
            if (!adminLink) return;
            var show = false;
            if (
                a &&
                window.ms365OperatorAccess &&
                typeof window.ms365OperatorAccess.shouldShowAdminMenuLink === 'function'
            ) {
                show = !!window.ms365OperatorAccess.shouldShowAdminMenuLink();
            } else if (
                a &&
                window.ms365OperatorAccess &&
                typeof window.ms365OperatorAccess.isCurrentUserOperator === 'function'
            ) {
                show = !!window.ms365OperatorAccess.isCurrentUserOperator();
            }
            adminLink.hidden = !show;
            if (
                show &&
                window.ms365OperatorAccess &&
                window.ms365OperatorAccess.resolveAppRootHref
            ) {
                adminLink.href = window.ms365OperatorAccess.resolveAppRootHref('admin.html');
            }
        };
        const operatorReady =
            !!(
                window.ms365OperatorAccess &&
                typeof window.ms365OperatorAccess.refreshOperatorStatus === 'function' &&
                window.ms365LicenseApi &&
                typeof window.ms365LicenseApi.fetchAdminMe === 'function'
            );
        applyAdminLink();
        var skipOperator =
            !!(window.MS365_LICENSE_API && window.MS365_LICENSE_API.skipOperatorCheck === true);
        if (a && operatorReady && !skipOperator) {
            var gen = (setWidgetState._operatorRefreshGen || 0) + 1;
            setWidgetState._operatorRefreshGen = gen;
            window.ms365OperatorAccess.refreshOperatorStatus().then(function () {
                if (setWidgetState._operatorRefreshGen !== gen) return;
                applyAdminLink();
            });
        } else if (a && skipOperator) {
            applyAdminLink();
            if (adminLink) adminLink.hidden = true;
        } else if (a && !operatorReady) {
            /* pin-gate lädt operator-access/license-api deferred – kurz nachziehen */
            if (!setWidgetState._operatorRetryTimers) setWidgetState._operatorRetryTimers = 0;
            if (setWidgetState._operatorRetryTimers < 12) {
                setWidgetState._operatorRetryTimers += 1;
                setTimeout(function () {
                    setWidgetState({ silent: true });
                }, 250);
            }
        } else if (!a && window.ms365OperatorAccess && window.ms365OperatorAccess.clearOperatorCache) {
            window.ms365OperatorAccess.clearOperatorCache();
            applyAdminLink();
            setWidgetState._operatorRetryTimers = 0;
        }
        const actionLogLink = document.getElementById('ms365AuthActionLogLink');
        if (
            actionLogLink &&
            window.ms365OperatorAccess &&
            typeof window.ms365OperatorAccess.resolveAppRootHref === 'function'
        ) {
            actionLogLink.href = window.ms365OperatorAccess.resolveAppRootHref('action-log.html');
        }
        if (window.ms365SchoolYearUi && typeof window.ms365SchoolYearUi.bindSchoolYearControls === 'function') {
            window.ms365SchoolYearUi.bindSchoolYearControls();
        }
        refreshBackupHeaderStatus();
        if (trigger) {
            trigger.setAttribute('aria-label', a ? 'Konto: ' + accountLabel(a) : 'Konto');
            trigger.title = a ? accountLabel(a) : 'Konto';
        }
        if (menu) menu.hidden = !a;
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
                    btn.disabled = true;
                    const prev = btn.innerHTML;
                    btn.innerHTML = '<i class="bi bi-hourglass-split"></i>Anmelden …';
                    login(resolvePreferredLoginScopes())
                        .catch(function (e) {
                            const msg = (e && e.message) || String(e || 'Anmeldung fehlgeschlagen');
                            if (typeof window.ms365ToastOrAlert === 'function') {
                                window.ms365ToastOrAlert(msg);
                            } else {
                                window.alert(msg);
                            }
                        })
                        .finally(function () {
                            btn.disabled = false;
                            if (!getAccount()) {
                                btn.innerHTML = prev || '<i class="bi bi-box-arrow-in-right"></i>Anmelden';
                            }
                        });
                };
            }
        }
        setWidgetState._lastLoggedIn = loggedIn;
        setWidgetState._lastLabel = label;
        if (typeof window.ms365SyncFrontendPlannerHeaderChrome === 'function') {
            window.ms365SyncFrontendPlannerHeaderChrome();
        }
        // Kein Event-Sturm: nur bei echtem Login-/Logout-Wechsel benachrichtigen
        if (silent) return;
        if (prevLoggedIn === loggedIn && prevLabel === label && prevLoggedIn !== undefined) return;
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-auth-state-changed', {
                    detail: { loggedIn: loggedIn, accountLabel: label }
                })
            );
        } catch {
            /* ignore */
        }
    }

    async function init() {
        if (typeof document === 'undefined') return;
        try {
            const headerMod = await import('./app-global-header.js');
            if (headerMod && typeof headerMod.ensureAppHeaderChrome === 'function') {
                headerMod.ensureAppHeaderChrome();
            } else if (headerMod && typeof headerMod.mountAppGlobalHeader === 'function') {
                headerMod.mountAppGlobalHeader();
            }
        } catch {
            /* ignore */
        }
        try {
            const chromePol = await import('./frontend-planner-chrome-policy.js');
            if (chromePol && typeof chromePol.bootFrontendPlannerChromePolicy === 'function') {
                chromePol.bootFrontendPlannerChromePolicy();
            }
        } catch {
            /* ignore */
        }
        ensureHeaderWidget();
        try {
            window.dispatchEvent(new CustomEvent('ms365-auth-widget-ready'));
        } catch {
            // ignore
        }
        try {
            await withTimeout(ensurePca(), 8000);
        } catch {
            // ignore (widget still renders; Anmelden bleibt nutzbar)
        }
        setWidgetState();
        import('./dashboard-auth-menu-policy.js')
            .then(function (m) {
                if (m && typeof m.bootAuthMenuAudiencePolicy === 'function') {
                    m.bootAuthMenuAudiencePolicy();
                }
            })
            .catch(function () {
                /* ignore */
            });
        // Admin: kein SSO-Silent (vermeidet Hänger/Freezes auf localhost)
        const path = String((window.location && window.location.pathname) || '');
        const isAdmin = /\/admin\.html(?:\?|#|$)/i.test(path);
        if (!isAdmin) {
            try {
                if (pca && !getAccount()) {
                    await trySsoSilent(DEFAULT_SCOPES);
                }
            } catch {
                // ignore
            }
            setWidgetState();
        }
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
    window.ms365AuthRememberReturnUrl = rememberPostLoginReturnUrl;
    window.ms365AuthAcquireToken = acquireToken;
    window.ms365AuthAcquireTokenSilent = acquireTokenSilentOnly;
    window.ms365AuthAcquireTokenPopup = acquireTokenPopup;
    window.ms365AuthAcquireIdToken = acquireIdToken;
    window.ms365AuthAcquireIdTokenPopup = acquireIdTokenPopup;
    window.ms365AuthGetAccountInfo = function () {
        try {
            const a = getAccount();
            if (!a) return null;
            const claims = a.idTokenClaims || {};
            return {
                username: a.username ? String(a.username) : '',
                name: a.name ? String(a.name) : '',
                tenantId: String(a.tenantId || claims.tid || '').trim(),
                oid: String(claims.oid || '').trim(),
                upn: String(claims.preferred_username || claims.upn || a.username || '').trim()
            };
        } catch {
            return null;
        }
    };
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
            // Nur platzieren, nicht erneut voll initialisieren (vermeidet Event-Schleifen)
            const header = document.getElementById('adminAppTopActions') || document.querySelector('.header') || document.querySelector('header');
            if (!header) return;
            if (document.getElementById('ms365AuthWidget')) return;
            placeAuthWidgetInMenuHeader();
        });
    } catch {
        /* ignore */
    }
    try {
        window.addEventListener('ms365-spo-sync-status', function () {
            refreshBackupHeaderStatus();
        });
        window.addEventListener('ms365-auth-state-changed', function () {
            refreshBackupHeaderStatus();
        });
    } catch {
        /* ignore */
    }
})();

