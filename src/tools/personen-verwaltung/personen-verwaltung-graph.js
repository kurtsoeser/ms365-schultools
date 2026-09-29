/**
 * Graph/MSAL für Personen-Verwaltung (Analyse 02 Phase B).
 * Move-first – Verhalten unverändert.
 */

export const GRAPH_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/User.ReadWrite.All',
    'https://graph.microsoft.com/Group.ReadWrite.All',
    'https://graph.microsoft.com/Organization.Read.All'
];

let msalMod = null;
let pca = null;

export async function loadMsal() {
    if (msalMod) return msalMod;
    try {
        msalMod = await import('https://esm.sh/@azure/msal-browser@3.26.1');
    } catch {
        msalMod = await import('https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.26.1/+esm');
    }
    return msalMod;
}

export function isInteractionRequired(e) {
    return (
        e &&
        (e.name === 'InteractionRequiredAuthError' ||
            e.errorCode === 'interaction_required' ||
            (typeof e.message === 'string' && e.message.indexOf('interaction_required') !== -1))
    );
}

export function resolveMsalConfig() {
    let cfg = window.MS365_MSAL_CONFIG;
    if (!cfg) cfg = {};
    let id = String(cfg.clientId || '').trim();
    if (!id) {
        const meta = document.querySelector('meta[name="ms365-graph-client-id"]');
        const fromMeta = meta && meta.getAttribute('content') ? meta.getAttribute('content').trim() : '';
        if (fromMeta) id = fromMeta;
    }
    if (!id) {
        throw new Error(
            'Keine clientId: ms365-config.js fehlt/leer oder blockiert. Seite mit Strg+F5 neu laden.'
        );
    }
    return {
        clientId: id,
        authority: cfg.authority || 'https://login.microsoftonline.com/organizations',
        redirectUri: (cfg.redirectUri || window.location.href.split('#')[0]).trim()
    };
}

export async function getPca() {
    const m = await loadMsal();
    const PublicClientApplication = m.PublicClientApplication || (m.default && m.default.PublicClientApplication);
    if (!PublicClientApplication) {
        throw new Error('MSAL: PublicClientApplication nicht gefunden (Import).');
    }
    const cfg = resolveMsalConfig();
    if (!pca) {
        pca = new PublicClientApplication({
            auth: {
                clientId: cfg.clientId,
                authority: cfg.authority,
                redirectUri: cfg.redirectUri
            },
            cache: {
                cacheLocation: 'sessionStorage',
                storeAuthStateInCookie: true
            }
        });
        await pca.initialize();
        await pca.handleRedirectPromise();
    }
    return pca;
}

export async function getGraphToken() {
    // Pilot M7: gemeinsame Auth-UI bevorzugen (eine PCA-Session)
    if (typeof window.ms365AuthAcquireToken === 'function') {
        try {
            if (typeof window.ms365AuthEnsureInitialized === 'function') {
                await window.ms365AuthEnsureInitialized();
            }
            return await window.ms365AuthAcquireToken(GRAPH_SCOPES);
        } catch (sharedErr) {
            /* Fallback auf lokale PCA, wenn Shared-Auth fehlt/scheitert */
            if (!isInteractionRequired(sharedErr) && window.ms365AuthIsLoggedIn && window.ms365AuthIsLoggedIn()) {
                throw sharedErr;
            }
        }
    }
    const instance = await getPca();
    let accounts = instance.getAllAccounts();
    if (!accounts.length) {
        await instance.loginPopup({ scopes: GRAPH_SCOPES, prompt: 'select_account' });
        accounts = instance.getAllAccounts();
    }
    if (!accounts.length) {
        throw new Error('Anmeldung abgebrochen.');
    }
    const req = { scopes: GRAPH_SCOPES, account: accounts[0] };
    try {
        return (await instance.acquireTokenSilent(req)).accessToken;
    } catch (e) {
        if (isInteractionRequired(e)) {
            return (await instance.acquireTokenPopup(req)).accessToken;
        }
        throw e;
    }
}

export function sleep(ms) {
    return new Promise(function (r) {
        setTimeout(r, ms);
    });
}

export async function graphRequest(method, pathOrUrl, token, body, extraHeaders) {
    const url =
        pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
    let attempt = 0;
    while (true) {
        const headers = { Authorization: 'Bearer ' + token };
        if (extraHeaders && typeof extraHeaders === 'object') {
            Object.assign(headers, extraHeaders);
        }
        if (body !== undefined && method !== 'GET' && method !== 'DELETE') {
            headers['Content-Type'] = 'application/json';
        }
        const res = await fetch(url, {
            method: method,
            headers: headers,
            body: body !== undefined ? JSON.stringify(body) : undefined
        });
        if (res.status === 429 && attempt < 8) {
            const ra = parseInt(res.headers.get('Retry-After') || '5', 10);
            await sleep((isNaN(ra) ? 5 : ra) * 1000);
            attempt++;
            continue;
        }
        return res;
    }
}

export async function graphJson(method, pathOrUrl, token, body, extraHeaders) {
    const res = await graphRequest(method, pathOrUrl, token, body, extraHeaders);
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = text;
        }
    }
    if (!res.ok) {
        const msg =
            typeof data === 'object' && data && data.error
                ? JSON.stringify(data.error)
                : text || String(res.status);
        throw new Error(method + ' ' + pathOrUrl + ': ' + msg);
    }
    return data || {};
}

export async function graphDelete(path, token) {
    const res = await graphRequest('DELETE', path, token, undefined);
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = text;
        }
    }
    if (!res.ok) {
        const msg =
            typeof data === 'object' && data && data.error
                ? JSON.stringify(data.error)
                : text || String(res.status);
        throw new Error('DELETE ' + path + ': ' + msg);
    }
}

export function appendLog(msg, kind) {
    const el = document.getElementById('pvLog');
    if (!el) return;
    const line = document.createElement('div');
    line.textContent = new Date().toLocaleTimeString() + '  ' + msg;
    if (kind === 'err') line.style.color = '#b00020';
    else if (kind === 'ok') line.style.color = '#0d8050';
    else if (kind === 'warn') line.style.color = '#856404';
    else line.style.color = '#212529';
    el.appendChild(line);
    el.scrollTop = el.scrollHeight;
}

export function clearLog() {
    const el = document.getElementById('pvLog');
    if (el) el.replaceChildren();
}

export async function fetchAllPages(token, initialPath, onProgress) {
    const out = [];
    let next = initialPath;
    let page = 0;
    while (next) {
        page++;
        const data = await graphJson('GET', next, token, undefined);
        const vals = data.value;
        if (Array.isArray(vals)) {
            for (let i = 0; i < vals.length; i++) out.push(vals[i]);
        }
        next = data['@odata.nextLink'] || null;
        if (onProgress) onProgress(out.length, page, !!next);
    }
    return out;
}

export function odataEscape(s) {
    return String(s || '').replace(/'/g, "''");
}
