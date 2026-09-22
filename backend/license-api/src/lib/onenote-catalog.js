'use strict';

/**
 * Zentrale OneNote-Vorlagen auf der Betreiber-Site.
 * Live-Graph (App-Only) ist seit 2025-03 von Microsoft blockiert → Fallback: SharePoint-Snapshot.
 */
const { getConfig } = require('./config');
const { getOperatorGraphToken } = require('./msal-app-only');
const { graphJson, resolveSiteId } = require('./sharepoint-license');
const {
    listSnapshotNotebooks,
    loadSnapshotNotebookTree,
    listSnapshotSectionPages,
    getSnapshotPageById,
    getSnapshotSectionExport,
    isOneNoteAppOnlyBlocked,
    appOnlyBlockedError,
    preferredNotebookName: snapshotPreferredName
} = require('./onenote-snapshot');

const GRAPH = 'https://graph.microsoft.com/v1.0';
const MAX_PAGES = 40;
const MAX_HTML_CHARS = 5_000_000;

/** @type {{ id: string, displayName: string, webUrl: string }|null} */
let siteCache = null;

function preferredNotebookName() {
    return snapshotPreferredName();
}

/**
 * @param {string} method
 * @param {string} url
 * @param {string} token
 * @param {string} [accept]
 */
async function graphText(method, url, token, accept) {
    const res = await fetch(url, {
        method,
        headers: {
            Authorization: 'Bearer ' + token,
            Accept: accept || 'text/html'
        }
    });
    const text = await res.text();
    if (!res.ok) {
        let msg = text || 'HTTP ' + res.status;
        let payload = null;
        try {
            payload = JSON.parse(text);
            if (payload && payload.error && payload.error.message) msg = payload.error.message;
        } catch {
            /* ignore */
        }
        const err = new Error(msg);
        err.status = res.status;
        err.payload = payload;
        throw err;
    }
    return text;
}

async function getCatalogSite() {
    if (siteCache && siteCache.id) return siteCache;
    const cfg = getConfig();
    const token = await getOperatorGraphToken();
    const site = await resolveSiteId(cfg.siteWebUrl, token);
    siteCache = {
        id: String(site.id),
        displayName: String(site.displayName || ''),
        webUrl: String(site.webUrl || cfg.siteWebUrl || '')
    };
    return siteCache;
}

function mapNotebook(row) {
    const lastBy =
        (row.lastModifiedBy && row.lastModifiedBy.user && row.lastModifiedBy.user.displayName) ||
        (row.lastModifiedBy && row.lastModifiedBy.application && row.lastModifiedBy.application.displayName) ||
        row.lastModifiedByName ||
        '';
    return {
        id: String(row.id || ''),
        displayName: String(row.displayName || ''),
        isDefault: !!row.isDefault,
        userRole: String(row.userRole || ''),
        lastModifiedDateTime: String(row.lastModifiedDateTime || row.updatedAt || '').trim(),
        lastModifiedByName: String(lastBy || '').trim()
    };
}

function mapSection(row) {
    return {
        id: String(row.id || ''),
        displayName: String(row.displayName || '')
    };
}

function mapSectionGroup(row) {
    const name = String(row.displayName || '');
    const lower = name.toLowerCase();
    let kind = 'other';
    if (lower.includes('content library') || lower.includes('inhaltsbibliothek')) kind = 'contentLibrary';
    else if (lower.includes('collaboration') || lower.includes('zusammenarbeit')) kind = 'collaboration';
    else if (lower.includes('teacher only') || lower.includes('nur lehrer') || lower.includes('lehrerbereich'))
        kind = 'teacherOnly';
    return {
        id: String(row.id || ''),
        displayName: name,
        kind
    };
}

/**
 * @param {Error} e
 * @param {() => Promise<unknown>} snapshotFn
 */
async function fallbackOrThrow(e, snapshotFn) {
    if (isOneNoteAppOnlyBlocked(e)) {
        try {
            return await snapshotFn();
        } catch (snapErr) {
            if (snapErr && (snapErr.code === 'snapshot_missing' || snapErr.status === 404)) {
                throw appOnlyBlockedError();
            }
            throw snapErr;
        }
    }
    throw e;
}

/**
 * @returns {Promise<{ site: object, notebooks: object[], preferredName: string, libraryHint: string }>}
 */
async function listCatalogNotebooks() {
    // Live-OneNote App-Only ist tot → Snapshot zuerst (SharePoint Sites.* funktioniert).
    try {
        return await listSnapshotNotebooks();
    } catch (snapErr) {
        if (!(snapErr && (snapErr.code === 'snapshot_missing' || snapErr.status === 404))) {
            // Corrupt/other: trotzdem Live versuchen (falls MS es wieder erlaubt)
            try {
                return await listLiveCatalogNotebooks();
            } catch (liveErr) {
                if (isOneNoteAppOnlyBlocked(liveErr)) throw snapErr;
                throw liveErr;
            }
        }
    }
    try {
        return await listLiveCatalogNotebooks();
    } catch (e) {
        return fallbackOrThrow(e, listSnapshotNotebooks);
    }
}

async function listLiveCatalogNotebooks() {
    const cfg = getConfig();
    const site = await getCatalogSite();
    const token = await getOperatorGraphToken();
    const url =
        GRAPH +
        '/sites/' +
        encodeURIComponent(site.id) +
        '/onenote/notebooks?$select=id,displayName,isDefault,userRole,lastModifiedDateTime,lastModifiedBy&$top=50';
    const data = await graphJson('GET', url, token);
    const notebooks = ((data && data.value) || []).map(mapNotebook).filter((n) => n.id);
    const prefer = preferredNotebookName().toLowerCase();
    notebooks.sort((a, b) => {
        const ap = a.displayName.toLowerCase() === prefer ? 0 : 1;
        const bp = b.displayName.toLowerCase() === prefer ? 0 : 1;
        if (ap !== bp) return ap - bp;
        return a.displayName.localeCompare(b.displayName, 'de');
    });
    return {
        site,
        notebooks,
        preferredName: preferredNotebookName(),
        libraryHint: (cfg.catalogLibraryName || 'MS365-Katalog') + '/notebooks',
        siteWebUrl: site.webUrl,
        via: 'catalog-api'
    };
}

/**
 * @param {string} notebookId
 */
async function loadCatalogNotebookTree(notebookId) {
    const nid = String(notebookId || '').trim();
    if (!nid) {
        const err = new Error('notebookId fehlt.');
        err.status = 400;
        throw err;
    }
    try {
        const site = await getCatalogSite();
        const token = await getOperatorGraphToken();
        const base =
            GRAPH + '/sites/' + encodeURIComponent(site.id) + '/onenote/notebooks/' + encodeURIComponent(nid);

        const [secData, grpData] = await Promise.all([
            graphJson('GET', base + '/sections?$select=id,displayName&$top=100', token),
            graphJson('GET', base + '/sectionGroups?$select=id,displayName&$top=100', token)
        ]);

        const sections = ((secData && secData.value) || []).map(mapSection).filter((s) => s.id);
        const groupsRaw = ((grpData && grpData.value) || []).map(mapSectionGroup).filter((g) => g.id);

        const groups = [];
        for (const g of groupsRaw) {
            let childSections = [];
            try {
                const child = await graphJson(
                    'GET',
                    GRAPH +
                        '/sites/' +
                        encodeURIComponent(site.id) +
                        '/onenote/sectionGroups/' +
                        encodeURIComponent(g.id) +
                        '/sections?$select=id,displayName&$top=100',
                    token
                );
                childSections = ((child && child.value) || []).map(mapSection).filter((s) => s.id);
            } catch {
                childSections = [];
            }
            groups.push(Object.assign({}, g, { sections: childSections }));
        }

        return {
            notebookId: nid,
            siteId: site.id,
            sections,
            groups,
            via: 'catalog-api'
        };
    } catch (e) {
        return fallbackOrThrow(e, () => loadSnapshotNotebookTree(nid));
    }
}

/**
 * @param {string} sectionId
 */
async function listCatalogSectionPages(sectionId) {
    const sid = String(sectionId || '').trim();
    if (!sid) {
        const err = new Error('sectionId fehlt.');
        err.status = 400;
        throw err;
    }
    try {
        const site = await getCatalogSite();
        const token = await getOperatorGraphToken();
        const url =
            GRAPH +
            '/sites/' +
            encodeURIComponent(site.id) +
            '/onenote/sections/' +
            encodeURIComponent(sid) +
            '/pages?$top=' +
            MAX_PAGES +
            '&$select=id,title,createdDateTime,lastModifiedDateTime,links&$orderby=lastModifiedDateTime desc';
        const data = await graphJson('GET', url, token);
        const pages = ((data && data.value) || [])
            .map((p) => ({
                id: String(p.id || ''),
                title: String(p.title || 'Ohne Titel'),
                webUrl: (p.links && p.links.oneNoteWebUrl && p.links.oneNoteWebUrl.href) || ''
            }))
            .filter((p) => p.id);
        return { sectionId: sid, pages, via: 'catalog-api' };
    } catch (e) {
        return fallbackOrThrow(e, () => listSnapshotSectionPages(sid));
    }
}

/**
 * @param {string} pageId
 */
async function getCatalogPagePreview(pageId) {
    const pid = String(pageId || '').trim();
    if (!pid) {
        const err = new Error('pageId fehlt.');
        err.status = 400;
        throw err;
    }
    try {
        const site = await getCatalogSite();
        const token = await getOperatorGraphToken();
        const url =
            GRAPH +
            '/sites/' +
            encodeURIComponent(site.id) +
            '/onenote/pages/' +
            encodeURIComponent(pid) +
            '/preview';
        const data = await graphJson('GET', url, token);
        return {
            pageId: pid,
            previewText: String((data && data.previewText) || '').trim(),
            previewImageUrl:
                (data &&
                    data.links &&
                    data.links.previewImageUrl &&
                    (data.links.previewImageUrl.href || data.links.previewImageUrl)) ||
                '',
            via: 'catalog-api'
        };
    } catch (e) {
        return fallbackOrThrow(e, async () => {
            const page = await getSnapshotPageById(pid);
            return {
                pageId: pid,
                previewText: page.previewText || '',
                previewImageUrl: '',
                via: 'catalog-snapshot'
            };
        });
    }
}

/**
 * HTML-Inhalt einer Seite (für Neuanlage im Schul-Tenant).
 * @param {string} pageId
 */
async function getCatalogPageContent(pageId) {
    const pid = String(pageId || '').trim();
    if (!pid) {
        const err = new Error('pageId fehlt.');
        err.status = 400;
        throw err;
    }
    try {
        const site = await getCatalogSite();
        const token = await getOperatorGraphToken();
        const url =
            GRAPH +
            '/sites/' +
            encodeURIComponent(site.id) +
            '/onenote/pages/' +
            encodeURIComponent(pid) +
            '/content';
        const html = await graphText('GET', url, token, 'text/html');
        if (html.length > MAX_HTML_CHARS) {
            const err = new Error('Seite ist zu groß für den Katalog-Transfer (Limit ~1,5 MB HTML).');
            err.status = 413;
            throw err;
        }
        return { pageId: pid, html, contentType: 'text/html', via: 'catalog-api' };
    } catch (e) {
        if (e && e.status === 413) throw e;
        return fallbackOrThrow(e, () => getSnapshotPageById(pid));
    }
}

/**
 * Alle Seiten-HTML eines Abschnitts (für Verteilen ohne Site-Zugriff).
 * @param {string} sectionId
 */
async function getCatalogSectionExport(sectionId) {
    const sid = String(sectionId || '').trim();
    try {
        const listed = await listCatalogSectionPages(sid);
        if (listed.via === 'catalog-snapshot') {
            return getSnapshotSectionExport(sid);
        }
        const pages = [];
        for (const p of listed.pages) {
            try {
                const content = await getCatalogPageContent(p.id);
                pages.push({
                    id: p.id,
                    title: p.title,
                    html: content.html
                });
            } catch (e) {
                pages.push({
                    id: p.id,
                    title: p.title,
                    html: '',
                    error: (e && e.message) || String(e)
                });
            }
        }
        return {
            sectionId: sid,
            pageCount: pages.length,
            pages,
            via: 'catalog-api'
        };
    } catch (e) {
        return fallbackOrThrow(e, () => getSnapshotSectionExport(sid));
    }
}

module.exports = {
    listCatalogNotebooks,
    loadCatalogNotebookTree,
    listCatalogSectionPages,
    getCatalogPagePreview,
    getCatalogPageContent,
    getCatalogSectionExport,
    preferredNotebookName
};
