/**
 * Tenant-weite Suche nach der IT-Sicherungsbibliothek (MS365-IT-Stammdaten).
 * Wenn lokal keine Site-URL bekannt ist (neuer Browser/PC).
 */
import {
    IT_LIBRARY_TITLE,
    IT_LIBRARY_SITE_SEARCH_TERMS,
    scoreItLibraryDiscoverySite
} from './stammdaten-sharepoint-sync-logic.js';

const SCOPES_GRAPH = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/Group.Read.All'
];

function getG() {
    const G = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!G) throw new Error('SharePoint-Graph-Helfer nicht geladen (spo-graph-shared.js).');
    return G;
}

async function searchSites(token, query) {
    const G = getG();
    const q = String(query || '').trim();
    if (!q) return [];
    try {
        const path =
            '/sites?search=' +
            encodeURIComponent(q) +
            '&$select=id,displayName,webUrl&$top=15';
        const data = await G.graphJson('GET', path, token, undefined, 'v1.0');
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

async function findListByTitle(token, siteId, listTitle) {
    const G = getG();
    const title = String(listTitle || '').trim();
    if (!title || !siteId) return null;
    const path =
        G.graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await G.graphJson('GET', path, token, undefined, 'v1.0');
    const list = (data && data.value) || [];
    return list[0] || null;
}

async function getListDrive(token, siteId, listId) {
    const G = getG();
    return G.graphJson(
        'GET',
        G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/drive',
        token,
        undefined,
        'v1.0'
    );
}

/**
 * @param {{ siteUrl: string, listTitle?: string, via: string }} hit
 * @param {string} token
 * @param {string} listTitle
 */
async function resolveLibraryOnSite(token, siteUrl, listTitle, via) {
    const G = getG();
    const url = String(siteUrl || '').trim().replace(/\/$/, '');
    if (!url) return null;
    let site;
    try {
        site = await G.resolveSiteFromWebUrl(token, url);
    } catch {
        return null;
    }
    if (!site || !site.id) return null;

    let list;
    try {
        list = await findListByTitle(token, site.id, listTitle);
    } catch {
        return null;
    }
    if (!list || !(list.id || list.Id)) return null;

    const listId = String(list.id || list.Id);
    let drive;
    try {
        drive = await getListDrive(token, site.id, listId);
    } catch {
        return null;
    }
    if (!drive || !drive.id) return null;

    return {
        siteUrl: url,
        listTitle: listTitle,
        listId: listId,
        driveId: String(drive.id),
        webUrl: String(list.webUrl || drive.webUrl || ''),
        via: via || 'site'
    };
}

/**
 * @param {{ listTitle?: string, preferSiteUrl?: string, token?: string, searchTerms?: string[] }} [opts]
 * @returns {Promise<{ siteUrl: string, listTitle: string, listId: string, driveId: string, webUrl: string, via: string }|null>}
 */
export async function discoverItLibraryInTenant(opts) {
    const options = opts || {};
    const listTitle = String(options.listTitle || IT_LIBRARY_TITLE).trim() || IT_LIBRARY_TITLE;
    const prefer = String(options.preferSiteUrl || '').trim().replace(/\/$/, '');
    const terms = Array.isArray(options.searchTerms) && options.searchTerms.length
        ? options.searchTerms
        : IT_LIBRARY_SITE_SEARCH_TERMS;

    const G = getG();
    let token = options.token;
    if (!token) {
        try {
            token = await G.getGraphToken(SCOPES_GRAPH);
        } catch {
            return null;
        }
    }

    if (prefer) {
        const hit = await resolveLibraryOnSite(token, prefer, listTitle, 'prefer-site');
        if (hit) return hit;
    }

    const seen = new Set();
    const candidates = [];
    for (let i = 0; i < terms.length; i++) {
        const sites = await searchSites(token, terms[i]);
        sites.forEach(function (site) {
            const key = site.webUrl.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            candidates.push(site);
        });
    }

    candidates.sort(function (a, b) {
        return (
            scoreItLibraryDiscoverySite(a.displayName, a.webUrl) -
            scoreItLibraryDiscoverySite(b.displayName, b.webUrl)
        );
    });

    for (let j = 0; j < candidates.length; j++) {
        const hit = await resolveLibraryOnSite(
            token,
            candidates[j].webUrl,
            listTitle,
            'site-search'
        );
        if (hit) return hit;
    }

    return null;
}

export default { discoverItLibraryInTenant };
