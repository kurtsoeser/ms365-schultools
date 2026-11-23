/**
 * Kanonischer Graph-Client – R22 / b11 (180-Tage-Plan, Analyse 02 Phase A).
 *
 * Einheitlicher Transport für Microsoft Graph:
 * - Retry bei HTTP 429 und 5xx (Retry-After, sonst Exponential Backoff mit Cap)
 * - Paging über `@odata.nextLink` inkl. Truncation-Flag
 * - Journal-Hook → `ms365ActionLog` (Aktionsprotokoll)
 *
 * Token: bevorzugt `window.ms365GraphUnifiedGroups` (MSAL/PCA), sonst Header-Auth.
 * Neue und migrierte Tools sollen **nur** dieses Modul für Fetch/Paging nutzen –
 * kein lokales `fetch('https://graph.microsoft.com/…')` und kein dupliziertes Retry.
 *
 * @example
 *   import { getGraphToken, graphJson, fetchAllPages, noteGraphAction } from '../../shared/graph-client.js';
 *   const token = await getGraphToken(READ_SCOPES);
 *   const { items, truncated } = await fetchAllPages(token, '/groups?$top=100');
 *   noteGraphAction({ tool: 'mein-tool', action: 'group-patch', target: id, summary: '…' });
 */

import { accessTokenIncludesScopes, scopesForGraphCall, scopesRequireItApp } from './auth-login-scopes.js';

/** @type {typeof fetch | null} */
let fetchImpl = null;

/** @type {null | ((entry: object) => void)} */
let journalHook = null;

/**
 * @param {typeof fetch} [fn]
 */
export function setFetchImpl(fn) {
    fetchImpl = typeof fn === 'function' ? fn : null;
}

export function getFetchImpl() {
    if (fetchImpl) return fetchImpl;
    if (typeof fetch === 'function') return fetch.bind(globalThis);
    throw new Error('Graph-Client: fetch ist nicht verfügbar.');
}

/**
 * Journal-Hook setzen (z. B. Tests oder SharePoint-Journal).
 * `null` stellt den Default (`ms365ActionLog`) wieder her.
 * @param {null | ((entry: object) => void)} fn
 */
export function setJournalHook(fn) {
    journalHook = typeof fn === 'function' ? fn : null;
}

export function getJournalHook() {
    return journalHook;
}

/**
 * @returns {typeof window.ms365GraphUnifiedGroups}
 */
export function getGraphApi() {
    const g = typeof window !== 'undefined' ? window.ms365GraphUnifiedGroups : null;
    if (!g || typeof g.getGraphToken !== 'function') {
        throw new Error(
            'Graph-Client: ms365GraphUnifiedGroups nicht geladen. Bitte graph-unified-groups.js vor dem Tool-Script einbinden.'
        );
    }
    return g;
}

/**
 * Ohne Argument: Portal-Default User.Read.
 * Nicht Group.ReadWrite.All — Schreibscopes kommen über scopesForGraphCall.
 */
export function getDefaultGraphScopes() {
    const g = typeof window !== 'undefined' ? window.ms365GraphUnifiedGroups : null;
    const fromShared = g && Array.isArray(g.GRAPH_SCOPES) ? g.GRAPH_SCOPES.filter(Boolean) : [];
    if (fromShared.length) return fromShared.slice();
    return ['https://graph.microsoft.com/User.Read'];
}

/**
 * @param {string} method
 * @param {string} pathOrUrl
 * @param {string} token
 * @returns {Promise<string>}
 */
async function ensureDelegatedGraphToken(method, pathOrUrl, token) {
    const needed = scopesForGraphCall(method, pathOrUrl);
    if (!needed || !needed.length) return token;
    // IT-Scopes immer neu holen. Ein Portal-Token mit noch vorhandenem
    // Write.All darf den Schreibaufruf nicht weiter bedienen.
    if (scopesRequireItApp(needed)) return getGraphToken(needed);
    if (token && accessTokenIncludesScopes(token, needed)) return token;
    return getGraphToken(needed);
}

/**
 * @param {string[]} [scopes]
 * @returns {Promise<string>}
 */
export async function getGraphToken(scopes) {
    const list = scopes && scopes.length ? scopes : null;
    try {
        const g = getGraphApi();
        return g.getGraphToken(list || g.GRAPH_SCOPES);
    } catch (e) {
        if (typeof window !== 'undefined' && typeof window.ms365AuthAcquireToken === 'function') {
            return window.ms365AuthAcquireToken(list || getDefaultGraphScopes());
        }
        throw e;
    }
}

export function sleep(ms) {
    const g = typeof window !== 'undefined' ? window.ms365GraphUnifiedGroups : null;
    if (g && typeof g.sleep === 'function') return g.sleep(ms);
    return new Promise(function (r) {
        setTimeout(r, ms);
    });
}

export function odataEscape(s) {
    const g = typeof window !== 'undefined' ? window.ms365GraphUnifiedGroups : null;
    if (g && typeof g.odataEscape === 'function') return g.odataEscape(s);
    return String(s || '').replace(/'/g, "''");
}

/**
 * Eintrag ins Aktionsprotokoll (Journal-Hook).
 * Blockiert Writes nie: Fehler werden geschluckt.
 * @param {object} entry
 */
export function noteGraphAction(entry) {
    try {
        if (journalHook) {
            journalHook(entry);
            return;
        }
        if (typeof window !== 'undefined' && window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append(entry);
            return;
        }
        const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
        if (!api || typeof api.getSetup !== 'function' || typeof api.patchSetup !== 'function') return;
        const cur = api.getSetup() || {};
        const list = Array.isArray(cur.actionLog) ? cur.actionLog.slice() : [];
        list.push({
            at: new Date().toISOString(),
            tool: String((entry && entry.tool) || 'graph'),
            action: String((entry && entry.action) || 'write'),
            target: String((entry && entry.target) || ''),
            summary: String((entry && entry.summary) || ''),
            result: entry && entry.result === 'error' ? 'error' : entry && entry.result === 'skip' ? 'skip' : 'ok'
        });
        while (list.length > 200) list.shift();
        api.patchSetup({ actionLog: list });
    } catch {
        /* Protokoll darf Aufrufe nicht blockieren */
    }
}

/**
 * @param {number} status
 * @returns {boolean}
 */
export function shouldRetryGraphStatus(status) {
    const n = Number(status) || 0;
    return n === 429 || (n >= 500 && n <= 599);
}

/**
 * Wartezeit vor dem nächsten Versuch (ms).
 * Bevorzugt Retry-After (Sekunden); sonst Exponential Backoff mit Cap.
 *
 * @param {Response | null} res
 * @param {number} attempt 0-basiert (nach erstem Fehlschlag = 0)
 * @param {{
 *   url?: string,
 *   backoffBaseMs?: number,
 *   backoffCapMs?: number,
 *   defaultRetryAfterSec?: number
 * }} [opts]
 * @returns {number}
 */
export function resolveRetryWaitMs(res, attempt, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    const url = String(o.url || '');
    const isOnenote = /\/onenote\//i.test(url);
    const base = typeof o.backoffBaseMs === 'number' && o.backoffBaseMs > 0 ? o.backoffBaseMs : 500;
    const cap = typeof o.backoffCapMs === 'number' && o.backoffCapMs > 0 ? o.backoffCapMs : 30000;
    const defaultSec =
        typeof o.defaultRetryAfterSec === 'number' && o.defaultRetryAfterSec > 0 ? o.defaultRetryAfterSec : 5;

    let fromHeader = NaN;
    if (res && typeof res.headers?.get === 'function') {
        const raw = res.headers.get('Retry-After');
        if (raw != null && String(raw).trim()) {
            const asInt = parseInt(String(raw).trim(), 10);
            if (!Number.isNaN(asInt)) fromHeader = asInt * 1000;
        }
    }

    let waitMs;
    if (!Number.isNaN(fromHeader) && fromHeader >= 0) {
        waitMs = fromHeader;
    } else {
        const exp = Math.min(cap, base * Math.pow(2, Math.max(0, attempt)));
        waitMs = exp;
    }

    if (Number.isNaN(waitMs) || waitMs < 0) waitMs = defaultSec * 1000;
    if (isOnenote) waitMs = Math.min(waitMs, 8000);
    return Math.min(cap, waitMs);
}

/**
 * @param {string} pathOrUrl
 * @param {'v1.0'|'beta'|string} [version]
 */
export function resolveGraphUrl(pathOrUrl, version) {
    const path = String(pathOrUrl || '');
    if (path.indexOf('http') === 0) return path;
    const v = version === 'beta' ? 'beta' : 'v1.0';
    const base = 'https://graph.microsoft.com/' + v;
    return base + (path.indexOf('/') === 0 ? path : '/' + path);
}

/**
 * @param {string} method
 * @param {string} pathOrUrl
 * @param {string} token
 * @param {unknown} [body]
 * @param {Record<string, string>} [extraHeaders]
 * @param {{
 *   version?: string,
 *   rawBody?: boolean,
 *   maxRetries?: number,
 *   backoffBaseMs?: number,
 *   backoffCapMs?: number,
 *   signal?: AbortSignal,
 *   fetch?: typeof fetch
 * }} [opts]
 * @returns {Promise<Response>}
 */
export async function graphRequest(method, pathOrUrl, token, body, extraHeaders, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    token = await ensureDelegatedGraphToken(method, pathOrUrl, token);
    const url = resolveGraphUrl(pathOrUrl, o.version);
    const isOnenote = /\/onenote\//i.test(url);
    const maxRetries =
        typeof o.maxRetries === 'number' && o.maxRetries >= 0 ? o.maxRetries : isOnenote ? 3 : 8;
    const doFetch = typeof o.fetch === 'function' ? o.fetch : getFetchImpl();

    let attempt = 0;
    while (true) {
        /** @type {Record<string, string>} */
        const headers = { Authorization: 'Bearer ' + token };
        if (extraHeaders && typeof extraHeaders === 'object') {
            Object.assign(headers, extraHeaders);
        }

        let payload = undefined;
        if (body !== undefined) {
            if (o.rawBody) {
                payload = /** @type {BodyInit} */ (body);
            } else {
                headers['Content-Type'] = headers['Content-Type'] || 'application/json';
                payload = JSON.stringify(body);
            }
        }

        const res = await doFetch(url, {
            method: method,
            headers: headers,
            body: payload,
            signal: o.signal
        });

        if (shouldRetryGraphStatus(res.status) && attempt < maxRetries) {
            const waitMs = resolveRetryWaitMs(res, attempt, {
                url: url,
                backoffBaseMs: o.backoffBaseMs,
                backoffCapMs: o.backoffCapMs
            });
            await sleep(waitMs);
            attempt++;
            continue;
        }
        return res;
    }
}

/**
 * @param {string} method
 * @param {string} pathOrUrl
 * @param {string} token
 * @param {unknown} [body]
 * @param {Record<string, string>} [extraHeaders]
 * @param {{
 *   version?: string,
 *   rawBody?: boolean,
 *   maxRetries?: number,
 *   journal?: boolean | object,
 *   tool?: string,
 *   summary?: string,
 *   signal?: AbortSignal,
 *   fetch?: typeof fetch
 * }} [opts]
 */
export async function graphJson(method, pathOrUrl, token, body, extraHeaders, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    const res = await graphRequest(method, pathOrUrl, token, body, extraHeaders, o);
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
                ? data.error.message || JSON.stringify(data.error)
                : text || String(res.status);
        const err = new Error(method + ' ' + pathOrUrl + ': ' + msg);
        /** @type {any} */ (err).status = res.status;
        /** @type {any} */ (err).payload = data;
        if (o.journal) {
            const base =
                typeof o.journal === 'object' && o.journal
                    ? o.journal
                    : {
                          tool: o.tool || 'graph',
                          action: String(method || 'write').toLowerCase(),
                          target: String(pathOrUrl || ''),
                          summary: o.summary || method + ' ' + pathOrUrl
                      };
            noteGraphAction(Object.assign({}, base, { result: 'error' }));
        }
        throw err;
    }
    if (o.journal) {
        const base =
            typeof o.journal === 'object' && o.journal
                ? o.journal
                : {
                      tool: o.tool || 'graph',
                      action: String(method || 'write').toLowerCase(),
                      target: String(pathOrUrl || ''),
                      summary: o.summary || method + ' ' + pathOrUrl
                  };
        noteGraphAction(Object.assign({}, base, { result: 'ok' }));
    }
    return data || {};
}

/**
 * Paging mit Caps und Truncation-Flag (Analyse 01 K3 / R22).
 *
 * @param {string} token
 * @param {string} initialPath
 * @param {{
 *   maxItems?: number,
 *   maxPages?: number,
 *   onProgress?: (info: { page: number, loaded: number, hasMore: boolean }) => void,
 *   extraHeaders?: Record<string, string>,
 *   version?: string,
 *   maxRetries?: number,
 *   fetch?: typeof fetch
 * }} [opts]
 * @returns {Promise<{ items: object[], truncated: boolean, pages: number }>}
 */
export async function fetchAllPages(token, initialPath, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    const maxItems = typeof o.maxItems === 'number' && o.maxItems > 0 ? o.maxItems : 4000;
    const maxPages = typeof o.maxPages === 'number' && o.maxPages > 0 ? o.maxPages : 40;
    const out = [];
    let next = initialPath;
    let pages = 0;
    while (next && pages < maxPages && out.length < maxItems) {
        pages++;
        const data = await graphJson('GET', next, token, undefined, o.extraHeaders, {
            version: o.version,
            maxRetries: o.maxRetries,
            fetch: o.fetch
        });
        const vals = data.value;
        if (Array.isArray(vals)) {
            for (let i = 0; i < vals.length; i++) out.push(vals[i]);
        }
        next = data['@odata.nextLink'] || null;
        if (typeof o.onProgress === 'function') {
            o.onProgress({ page: pages, loaded: out.length, hasMore: !!next });
        }
    }
    return {
        items: out,
        truncated: !!next || out.length >= maxItems,
        pages: pages
    };
}

/**
 * Convenience: Gruppenmitglieder über Shared-API (inkl. truncated), falls geladen.
 * @param {string} token
 * @param {string} groupId
 */
export async function fetchGroupMembers(token, groupId) {
    return getGraphApi().fetchGroupMembers(token, groupId);
}

if (typeof window !== 'undefined') {
    window.ms365GraphClient = {
        getGraphApi: getGraphApi,
        getGraphToken: getGraphToken,
        graphRequest: graphRequest,
        graphJson: graphJson,
        fetchAllPages: fetchAllPages,
        fetchGroupMembers: fetchGroupMembers,
        sleep: sleep,
        odataEscape: odataEscape,
        getDefaultGraphScopes: getDefaultGraphScopes,
        noteGraphAction: noteGraphAction,
        setJournalHook: setJournalHook,
        shouldRetryGraphStatus: shouldRetryGraphStatus,
        resolveRetryWaitMs: resolveRetryWaitMs,
        resolveGraphUrl: resolveGraphUrl
    };
}
