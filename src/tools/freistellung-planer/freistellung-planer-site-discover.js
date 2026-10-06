/**
 * Freistellungsliste finden, wenn lokale Site-URL falsch ist (z. B. nur Mandanten-Stammweb).
 */
import { LIST_TITLE_DEFAULT } from './freistellung-planer-schema.js';
import { findFreistellungListOnSite } from './freistellung-planer-graph.js';
import { isLikelySharePointTenantRoot } from './freistellung-planer-state.js';

const SEARCH_TERMS = ['MS365-Schultools', 'Schultools', 'MS365 Schule', 'MS365'];

const DEFAULT_SERVER_PATHS = ['/sites/MS365-Schultools'];

function spoApi() {
    const G = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!G || typeof G.getGraphToken !== 'function' || typeof G.graphJson !== 'function') {
        return null;
    }
    return G;
}

function readScopes() {
    return [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.Read.All'
    ];
}

/** @returns {string[]} volle URLs oder server-relative Pfade */
export function freistellungSiteDiscoveryHints() {
    const out = [];
    try {
        const p = new URLSearchParams(typeof window !== 'undefined' ? window.location.search || '' : '');
        const fr = String(p.get('frSite') || p.get('siteUrl') || '').trim();
        if (fr) out.push(fr.replace(/\/$/, ''));
    } catch {
        /* ignore */
    }
    try {
        const cfg = typeof window !== 'undefined' ? window.MS365_FREISTELLUNG_PLANER : null;
        if (cfg && Array.isArray(cfg.sitePaths)) {
            cfg.sitePaths.forEach(function (path) {
                const s = String(path || '').trim();
                if (s) out.push(s);
            });
        }
        if (cfg && cfg.siteUrl) {
            out.push(String(cfg.siteUrl).trim().replace(/\/$/, ''));
        }
    } catch {
        /* ignore */
    }
    DEFAULT_SERVER_PATHS.forEach(function (path) {
        out.push(path);
    });
    const seen = new Set();
    return out.filter(function (entry) {
        const key = entry.toLowerCase();
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}

async function sharePointHostname(tok, G) {
    if (typeof G.getSharePointHostname === 'function') {
        const h = await G.getSharePointHostname(tok);
        if (h) return h;
    }
    try {
        const root = await G.graphJson('GET', '/sites/root?$select=webUrl', tok, undefined, 'v1.0');
        const w = root && root.webUrl ? String(root.webUrl) : '';
        if (w) return new URL(w).hostname;
    } catch {
        /* ignore */
    }
    return '';
}

/**
 * @param {string} tok
 * @param {string} query
 * @returns {Promise<{ id: string, displayName: string, webUrl: string }[]>}
 */
async function searchSites(tok, query) {
    const G = spoApi();
    if (!G) return [];
    const q = String(query || '').trim();
    if (!q) return [];
    try {
        const path =
            '/sites?search=' +
            encodeURIComponent(q) +
            '&$select=id,displayName,webUrl&$top=15';
        const data = await G.graphJson('GET', path, tok, undefined, 'v1.0');
        return ((data && data.value) || [])
            .map(function (s) {
                return {
                    id: String(s.id || ''),
                    displayName: String(s.displayName || ''),
                    webUrl: String(s.webUrl || '').replace(/\/$/, '')
                };
            })
            .filter(function (s) {
                return s.id && s.webUrl;
            });
    } catch {
        return [];
    }
}

/**
 * @param {string} siteWebUrl
 * @param {{ listName?: string, listId?: string }} opts
 * @returns {Promise<string>}
 */
async function tryListOnSite(siteWebUrl, opts) {
    const url = String(siteWebUrl || '').trim().replace(/\/$/, '');
    if (!url) return '';
    try {
        const list = await findFreistellungListOnSite(url, opts || {});
        return list && list.id ? String(list.id) : '';
    } catch {
        return '';
    }
}

async function resolveSiteWebUrlFromServerPath(tok, G, hostname, serverPath) {
    const rel = String(serverPath || '').trim();
    if (!hostname || !rel || /^https?:/i.test(rel)) return '';
    const normalized = rel.startsWith('/') ? rel : '/' + rel;
    try {
        const seg = encodeURIComponent(hostname + ':' + normalized);
        const site = await G.graphJson('GET', '/sites/' + seg, tok, undefined, 'v1.0');
        const w = site && site.webUrl ? String(site.webUrl).trim().replace(/\/$/, '') : '';
        return w;
    } catch {
        return '';
    }
}

/**
 * @param {{ listName?: string, listId?: string, preferSiteUrl?: string }} [opts]
 * @returns {Promise<{ siteUrl: string, listId: string, listName: string, via: string }|null>}
 */
export async function discoverFreistellungPlanerContext(opts) {
    const options = opts || {};
    const listName = String(options.listName || LIST_TITLE_DEFAULT).trim() || LIST_TITLE_DEFAULT;
    const hintListId = String(options.listId || '').trim();
    const prefer = String(options.preferSiteUrl || '').trim().replace(/\/$/, '');

    const G = spoApi();
    if (!G) return null;

    let tok;
    try {
        tok = await G.getGraphToken(readScopes());
    } catch {
        return null;
    }

    const trySite = async function (siteUrl, via) {
        const listId = await tryListOnSite(siteUrl, { listName, listId: hintListId });
        if (!listId) return null;
        return { siteUrl: siteUrl.replace(/\/$/, ''), listId, listName, via };
    };

    if (prefer && !isLikelySharePointTenantRoot(prefer)) {
        const hit = await trySite(prefer, 'prefer-site');
        if (hit) return hit;
    }

    const hints = freistellungSiteDiscoveryHints();
    const host = await sharePointHostname(tok, G);

    for (let i = 0; i < hints.length; i++) {
        const entry = hints[i];
        if (/^https?:\/\//i.test(entry)) {
            const hit = await trySite(entry, 'config-url');
            if (hit) return hit;
            continue;
        }
        if (!host) continue;
        const webUrl = await resolveSiteWebUrlFromServerPath(tok, G, host, entry);
        if (!webUrl) continue;
        const hit = await trySite(webUrl, 'getByPath');
        if (hit) return hit;
    }

    const seen = new Set();
    const candidates = [];

    for (let s = 0; s < SEARCH_TERMS.length; s++) {
        const sites = await searchSites(tok, SEARCH_TERMS[s]);
        sites.forEach(function (site) {
            const key = site.webUrl.toLowerCase();
            if (!seen.has(key)) {
                seen.add(key);
                candidates.push(site);
            }
        });
    }

    candidates.sort(function (a, b) {
        const score = function (name, url) {
            const n = String(name || '').toLowerCase();
            const u = String(url || '').toLowerCase();
            if (n.indexOf('schultools') >= 0 || u.indexOf('schultools') >= 0) return 0;
            if (n.indexOf('ms365') >= 0) return 1;
            return 2;
        };
        return score(a.displayName, a.webUrl) - score(b.displayName, b.webUrl);
    });

    for (let j = 0; j < candidates.length; j++) {
        const hit = await trySite(candidates[j].webUrl, 'site-search');
        if (hit) return hit;
    }

    if (prefer && isLikelySharePointTenantRoot(prefer)) {
        const rootHit = await trySite(prefer, 'tenant-root');
        if (rootHit) return rootHit;
    }

    return null;
}
