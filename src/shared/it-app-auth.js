/**
 * Zweite Entra-App (ms365-schultools-it) für Schreib- und andere
 * hochprivilegierte Aufrufe.
 *
 * Eigene PublicClientApplication. Access- und Refresh-Tokens legt MSAL
 * unter der IT-Client-ID ab (Schlüssel enthalten die Client-ID), nicht im
 * Cache der Portal-App. Dieselbe Redirect-URI und dieselbe Authority
 * (organizations, mehrere Mandanten) wie die Portal-Anmeldung.
 *
 * Fehlt die IT-Client-ID, gibt es einen deutschen Fehler und keinen
 * stillen Rückfall auf die Portal-App.
 */

import { loadMsalBrowser } from './msal-loader.js';
import {
    prepareItAppScopes,
    scopesRequireItApp
} from './auth-login-scopes.js';

export {
    clientKindForScopes,
    prepareItAppScopes,
    rewriteItSharePointScope,
    scopeRequiresItApp,
    scopesRequireItApp
} from './auth-login-scopes.js';

export const IT_CLIENT_MISSING_ERROR =
    'Die IT-App ms365-schultools-it ist nicht konfiguriert (keine Client-ID). Schreib- und Administrationszugriffe laufen nicht über die Portal-App. Bitte itClientId in ms365-config.js oder das Meta-Tag ms365-graph-it-client-id setzen.';

/** @type {import('@azure/msal-browser').PublicClientApplication | null} */
let itPca = null;
/** @type {Promise<import('@azure/msal-browser').PublicClientApplication> | null} */
let itInit = null;

/**
 * @param {{ MS365_MSAL_CONFIG?: { itClientId?: string }, document?: Document }} [env]
 * @returns {string}
 */
export function resolveItClientId(env) {
    const e = env || (typeof window !== 'undefined' ? window : {});
    const cfg = e.MS365_MSAL_CONFIG || {};
    let id = String(cfg.itClientId || '').trim();
    if (!id && e.document && typeof e.document.querySelector === 'function') {
        const meta = e.document.querySelector('meta[name="ms365-graph-it-client-id"]');
        const fromMeta = meta && meta.getAttribute('content') ? meta.getAttribute('content').trim() : '';
        if (fromMeta) id = fromMeta;
    }
    return id;
}

/**
 * @param {{ MS365_MSAL_CONFIG?: object, document?: Document, location?: Location }} [env]
 */
export function resolveItMsalConfig(env) {
    const e = env || (typeof window !== 'undefined' ? window : {});
    const cfg = e.MS365_MSAL_CONFIG || {};
    const id = resolveItClientId(e);
    if (!id) {
        const err = new Error(IT_CLIENT_MISSING_ERROR);
        err.code = 'ms365_it_client_missing';
        throw err;
    }
    let redirectUri = String(cfg.redirectUri || '').trim();
    if (!redirectUri && e.location) {
        redirectUri = String(e.location.href || '').split('#')[0];
    }
    return {
        clientId: id,
        authority: cfg.authority || 'https://login.microsoftonline.com/organizations',
        redirectUri: redirectUri
    };
}

/**
 * Hash oder Query enthält eine OAuth-Antwort und der state gehört zur IT-Client-ID.
 * @param {string} [clientId]
 * @returns {boolean}
 */
export function redirectStateMatchesClient(clientId) {
    const id = String(clientId || '').trim();
    if (!id || typeof window === 'undefined') return false;
    const params = authResponseParams();
    if (!params) return false;
    const state = params.get('state') || '';
    if (!state) return false;
    const stores = [];
    try {
        if (window.sessionStorage) stores.push(window.sessionStorage);
    } catch {
        /* ignore */
    }
    try {
        if (window.localStorage) stores.push(window.localStorage);
    } catch {
        /* ignore */
    }
    for (let s = 0; s < stores.length; s++) {
        const store = stores[s];
        for (let i = 0; i < store.length; i++) {
            const key = store.key(i) || '';
            if (key.indexOf(id) === -1) continue;
            if (key.indexOf(state) !== -1) return true;
            let value = '';
            try {
                value = store.getItem(key) || '';
            } catch {
                value = '';
            }
            if (value && value.indexOf(state) !== -1) return true;
        }
    }
    return false;
}

function authResponseParams() {
    if (typeof window === 'undefined' || !window.location) return null;
    const hash = String(window.location.hash || '').replace(/^#/, '');
    const search = String(window.location.search || '').replace(/^\?/, '');
    const fromHash = new URLSearchParams(hash);
    if (fromHash.get('code') || fromHash.get('error') || fromHash.get('id_token')) return fromHash;
    const fromSearch = new URLSearchParams(search);
    if (fromSearch.get('code') || fromSearch.get('error') || fromSearch.get('id_token')) return fromSearch;
    return null;
}

export function isInteractionRequired(e) {
    if (!e) return false;
    if (e.name === 'InteractionRequiredAuthError') return true;
    const code = String(e.errorCode || e.code || '').toLowerCase();
    if (
        code === 'interaction_required' ||
        code === 'consent_required' ||
        code === 'login_required' ||
        code === 'invalid_grant' ||
        code === 'no_account_in_silent_request' ||
        code === 'no_tokens_found' ||
        code === 'monitor_window_timeout' ||
        code === 'timed_out' ||
        code === 'block_iframe_reload'
    ) {
        return true;
    }
    const msg = String((e && e.message) || '').toLowerCase();
    return (
        msg.indexOf('interaction_required') !== -1 ||
        msg.indexOf('consent_required') !== -1 ||
        msg.indexOf('login_required') !== -1 ||
        msg.indexOf('monitor_window_timeout') !== -1 ||
        msg.indexOf('token acquisition in iframe failed') !== -1 ||
        msg.indexOf('aadsts65001') !== -1
    );
}

function isUserCancelled(e) {
    if (!e) return false;
    const msg = String((e && e.message) || e || '');
    const code = String((e && (e.errorCode || e.code)) || '');
    return /abgebrochen|cancelled|canceled|user_cancelled/i.test(msg) || /user_cancelled/i.test(code);
}

function isPopupWindowError(e) {
    if (!e) return false;
    const code = String((e && (e.errorCode || e.code)) || '');
    const msg = String((e && e.message) || e || '');
    return code === 'popup_window_error' || /popup_window_error/i.test(msg) || /error opening popup/i.test(msg);
}

function portalLoginHint() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            if (info && info.username) return String(info.username).trim();
            if (info && info.upn) return String(info.upn).trim();
        }
    } catch {
        /* ignore */
    }
    return '';
}

function pickAccount(instance, hint) {
    const all = typeof instance.getAllAccounts === 'function' ? instance.getAllAccounts() || [] : [];
    if (hint) {
        const h = hint.toLowerCase();
        for (let i = 0; i < all.length; i++) {
            if (String(all[i] && all[i].username ? all[i].username : '').toLowerCase() === h) return all[i];
        }
    }
    const active = typeof instance.getActiveAccount === 'function' ? instance.getActiveAccount() : null;
    if (active) return active;
    return all[0] || null;
}

const SKIP_SILENT_KEY = 'ms365-it-skip-silent';

function shouldSkipSilentIframe() {
    try {
        return sessionStorage.getItem(SKIP_SILENT_KEY) === '1';
    } catch {
        return false;
    }
}

function markSkipSilentIframe() {
    try {
        sessionStorage.setItem(SKIP_SILENT_KEY, '1');
    } catch {
        /* ignore */
    }
}

function clearSkipSilentIframe() {
    try {
        sessionStorage.removeItem(SKIP_SILENT_KEY);
    } catch {
        /* ignore */
    }
}

function rememberReturnUrl() {
    try {
        if (typeof window.ms365AuthRememberReturnUrl === 'function') {
            window.ms365AuthRememberReturnUrl(window.location.href);
            return;
        }
        sessionStorage.setItem('ms365-post-login-url', window.location.href);
    } catch {
        /* ignore */
    }
}

export async function ensureItPca() {
    if (itPca) return itPca;
    if (itInit) return itInit;
    itInit = (async () => {
        try {
            const cfg = resolveItMsalConfig();
            const m = await loadMsalBrowser();
            const PublicClientApplication =
                m.PublicClientApplication || (m.default && m.default.PublicClientApplication);
            if (!PublicClientApplication) throw new Error('MSAL: PublicClientApplication nicht gefunden.');
            const onSchooltool = /\/ms365-schooltool\.html$/i.test(
                String((typeof location !== 'undefined' && location.pathname) || '')
            );
            const instance = new PublicClientApplication({
                auth: {
                    clientId: cfg.clientId,
                    authority: cfg.authority,
                    redirectUri: cfg.redirectUri,
                    // Redirect-Seite: Rückkehr macht ms365-schooltool-redirect.js.
                    navigateToLoginRequestUrl: !onSchooltool
                },
                cache: {
                    cacheLocation: 'localStorage',
                    storeAuthStateInCookie: true
                }
            });
            await instance.initialize();
            try {
                await instance.handleRedirectPromise();
            } catch {
                /* Redirect der Portal-App gehört nicht hierher */
            }
            const hint = portalLoginHint();
            const account = pickAccount(instance, hint);
            if (account && typeof instance.setActiveAccount === 'function') {
                instance.setActiveAccount(account);
            }
            itPca = instance;
            return itPca;
        } catch (e) {
            itInit = null;
            throw e;
        }
    })();
    return itInit;
}

/**
 * Token der IT-App. Erster Aufruf interaktiv (Zustimmung der Schule, dasselbe
 * Konto wie die Portal-Anmeldung, kein Kontowechsel). Danach still.
 *
 * @param {string[]} scopes
 * @param {{ mode?: 'interactive' | 'popup' | 'silent' | 'fresh', forceRefresh?: boolean }} [opts]
 * @returns {Promise<string>}
 */
export async function acquireItAppToken(scopes, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    const prepared = prepareItAppScopes(Array.isArray(scopes) ? scopes : []);
    if (!scopesRequireItApp(prepared)) {
        throw new Error('Intern: acquireItAppToken ohne IT-Scope aufgerufen.');
    }
    const instance = await ensureItPca();
    const hint = portalLoginHint();
    const account = pickAccount(instance, hint);
    const mode = o.mode || 'interactive';
    const req = { scopes: prepared };
    if (account) req.account = account;
    if (hint) req.loginHint = hint;
    if (o.forceRefresh || mode === 'fresh') req.forceRefresh = true;

    // Nach einem Iframe-Timeout den nächsten Klick sofort als Fenster öffnen.
    // Sonst vergeht die Klick-Geste und der Browser blockt das Fenster.
    const skipSilent = mode === 'popup' && shouldSkipSilentIframe();
    const trySilent = account && mode !== 'interactive-only' && !skipSilent;
    if (trySilent) {
        try {
            const result = await instance.acquireTokenSilent(req);
            clearSkipSilentIframe();
            return result.accessToken;
        } catch (e) {
            if (isInteractionRequired(e)) markSkipSilentIframe();
            if (mode === 'silent' || mode === 'fresh') throw e;
            if (!isInteractionRequired(e)) throw e;
        }
    } else if (!account && (mode === 'silent' || mode === 'fresh')) {
        const err = new Error('Nicht an der IT-App angemeldet.');
        err.code = 'ms365_it_interaction_required';
        throw err;
    }

    rememberReturnUrl();
    const interactiveReq = { scopes: prepared };
    if (account) interactiveReq.account = account;
    if (hint) interactiveReq.loginHint = hint;

    if (typeof instance.acquireTokenPopup === 'function') {
        try {
            const result = await instance.acquireTokenPopup(interactiveReq);
            if (result && result.account && typeof instance.setActiveAccount === 'function') {
                instance.setActiveAccount(result.account);
            }
            clearSkipSilentIframe();
            return result.accessToken;
        } catch (e) {
            if (isUserCancelled(e)) throw e;
            if (mode === 'popup') {
                if (isPopupWindowError(e)) {
                    throw new Error(
                        'Das Microsoft-Fenster wurde blockiert. Pop-ups für diese Seite erlauben und den Button noch einmal drücken.'
                    );
                }
                throw e;
            }
            if (!isPopupWindowError(e)) throw e;
        }
    }

    if (mode === 'popup') {
        throw new Error(
            'Anmeldung braucht ein Microsoft-Fenster. Pop-ups erlauben und den Button noch einmal drücken.'
        );
    }

    interactiveReq.redirectStartPage = window.location.href;
    await instance.acquireTokenRedirect(interactiveReq);
    throw new Error('Weiterleitung zur Anmeldung an der IT-App …');
}

/**
 * Token der IT-App, wenn die Scopes das verlangen. Sonst `null` (Aufrufer bleibt
 * auf der Portal-App). Wirft, wenn die IT-Client-ID fehlt — kein Portal-Fallback.
 *
 * @param {string[] | null | undefined} scopes
 * @param {{ mode?: 'interactive' | 'popup' | 'silent' | 'fresh', forceRefresh?: boolean }} [opts]
 * @returns {Promise<string | null>}
 */
export async function acquireItTokenIfRequired(scopes, opts) {
    const list = Array.isArray(scopes) ? scopes : [];
    const prepared = prepareItAppScopes(list);
    if (!scopesRequireItApp(list) && !scopesRequireItApp(prepared)) return null;
    return acquireItAppToken(prepared, opts);
}
