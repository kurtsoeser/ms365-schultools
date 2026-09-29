/**
 * Kanonischer Graph-Client (ESM-Fassade) – Analyse 02 Phase A.
 *
 * Neue und migrierte Tools sollen **nur** dieses Modul für Token/Fetch/Paging
 * nutzen – keine lokale PCA und kein dupliziertes `graphRequest`.
 *
 * Implementierung: Delegate an `window.ms365GraphUnifiedGroups`
 * (`graph-unified-groups.js`, IIFE). Truncation/Identity: siehe
 * `membership-reconcile.js` (`membershipFetchGuard`, `memberEmailsFromGraph`).
 *
 * @example
 *   import { getGraphToken, graphJson, fetchAllPages } from '../../shared/graph-client.js';
 *   const token = await getGraphToken(READ_SCOPES);
 *   const { items, truncated } = await fetchAllPages(token, '/groups?$top=100');
 */

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

/** Standard-Scopes aus dem Shared-Modul (falls vorhanden). */
export function getDefaultGraphScopes() {
    const g = typeof window !== 'undefined' ? window.ms365GraphUnifiedGroups : null;
    return (g && g.GRAPH_SCOPES) || [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.ReadWrite.All'
    ];
}

/**
 * @param {string[]} [scopes]
 * @returns {Promise<string>}
 */
export async function getGraphToken(scopes) {
    const g = getGraphApi();
    return g.getGraphToken(scopes && scopes.length ? scopes : g.GRAPH_SCOPES);
}

/**
 * @param {string} method
 * @param {string} pathOrUrl
 * @param {string} token
 * @param {unknown} [body]
 * @param {Record<string, string>} [extraHeaders]
 */
export async function graphRequest(method, pathOrUrl, token, body, extraHeaders) {
    return getGraphApi().graphRequest(method, pathOrUrl, token, body, extraHeaders);
}

/**
 * @param {string} method
 * @param {string} pathOrUrl
 * @param {string} token
 * @param {unknown} [body]
 * @param {Record<string, string>} [extraHeaders]
 */
export async function graphJson(method, pathOrUrl, token, body, extraHeaders) {
    return getGraphApi().graphJson(method, pathOrUrl, token, body, extraHeaders);
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
 * Paging mit Caps und Truncation-Flag (Analyse 01 K3).
 *
 * @param {string} token
 * @param {string} initialPath
 * @param {{
 *   maxItems?: number,
 *   maxPages?: number,
 *   onProgress?: (info: { page: number, loaded: number, hasMore: boolean }) => void,
 *   extraHeaders?: Record<string, string>
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
        const data = await graphJson('GET', next, token, undefined, o.extraHeaders);
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
 * Convenience: Gruppenmitglieder über Shared-API (inkl. truncated).
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
        getDefaultGraphScopes: getDefaultGraphScopes
    };
}
