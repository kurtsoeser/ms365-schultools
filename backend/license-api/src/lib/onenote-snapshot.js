'use strict';

/**
 * OneNote-Vorlagen als SharePoint-JSON-Snapshot (App-Only Sites.* funktioniert;
 * Live-OneNote Graph App-Only ist seit 2025-03 von Microsoft abgeschaltet).
 *
 * Struktur in MS365-Katalog:
 *   onenote-snapshot/index.json
 *   onenote-snapshot/notebooks/{notebookId}/tree.json
 *   onenote-snapshot/sections/{sectionId}.json
 */
const { getConfig } = require('./config');
const { getOperatorGraphToken } = require('./msal-app-only');
const { graphJson } = require('./sharepoint-license');
const {
    encodeDrivePath,
    clearCatalogCache
} = require('./sharepoint-catalog');

// resolveCatalogDrive/contentUrl are not exported – duplicate minimal helpers via require of private APIs
// Instead: use writeMaterialFile style by exporting new helpers below on the catalog module.
// We inline drive access by requiring the same pattern.

const GRAPH = 'https://graph.microsoft.com/v1.0';
const SNAPSHOT_ROOT = 'onenote-snapshot';
const MAX_SNAPSHOT_BYTES = 12 * 1024 * 1024;
/** Simple PUT oft unzuverlässig ab ~4 MB → Upload-Session. */
const SIMPLE_UPLOAD_MAX = 3.5 * 1024 * 1024;
/** Ab dieser HTML-Größe pro Seite → eigene Datei (Bilder/Karten). */
const INLINE_HTML_MAX = 80_000;

/** @type {{ siteId: string, listId: string, driveId: string, webUrl: string, library: string, at: number } | null} */
let driveCache = null;
const DRIVE_CACHE_MS = 10 * 60 * 1000;

function preferredNotebookName() {
    return String(process.env.CATALOG_ONENOTE_NOTEBOOK || 'MS365-Vorlagen-Notizbuch').trim();
}

/** SharePoint-sichere Segment-ID (OneNote-IDs können Sonderzeichen haben). */
function fileId(id) {
    return String(id || '')
        .trim()
        .replace(/[^a-zA-Z0-9._=-]+/g, '_')
        .slice(0, 180);
}

function isSafeSnapshotRel(rel) {
    const p = String(rel || '')
        .replace(/\\/g, '/')
        .replace(/^\/+/, '');
    if (!p.startsWith(SNAPSHOT_ROOT + '/') && p !== SNAPSHOT_ROOT) return false;
    if (p.includes('..')) return false;
    return true;
}

/**
 * @param {string} token
 */
async function resolveDrive(token) {
    const cfg = getConfig();
    const now = Date.now();
    if (
        driveCache &&
        now - driveCache.at < DRIVE_CACHE_MS &&
        driveCache.library === cfg.catalogLibraryName
    ) {
        return driveCache;
    }
    const { resolveSiteId, resolveListId } = require('./sharepoint-license');
    const site = await resolveSiteId(cfg.siteWebUrl, token);
    const list = await resolveListId(site.id, cfg.catalogLibraryName, token);
    const drive = await graphJson(
        'GET',
        GRAPH +
            '/sites/' +
            encodeURIComponent(site.id) +
            '/lists/' +
            encodeURIComponent(list.id) +
            '/drive?$select=id,webUrl',
        token
    );
    if (!drive.id) {
        const err = new Error('Katalog-Bibliothek ohne Drive.');
        err.status = 500;
        throw err;
    }
    driveCache = {
        siteId: site.id,
        listId: list.id,
        driveId: drive.id,
        webUrl: drive.webUrl || '',
        library: cfg.catalogLibraryName,
        at: now
    };
    return driveCache;
}

function contentUrl(driveId, relPath) {
    return (
        GRAPH + '/drives/' + encodeURIComponent(driveId) + '/root:/' + encodeDrivePath(relPath) + ':/content'
    );
}

/**
 * @param {string} relPath
 */
async function readJsonFile(relPath) {
    if (!isSafeSnapshotRel(relPath)) {
        const err = new Error('Ungültiger Snapshot-Pfad.');
        err.status = 400;
        throw err;
    }
    const token = await getOperatorGraphToken();
    const drive = await resolveDrive(token);
    const res = await fetch(contentUrl(drive.driveId, relPath), {
        method: 'GET',
        headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' }
    });
    const text = await res.text();
    if (!res.ok) {
        const err = new Error(
            res.status === 404
                ? 'OneNote-Snapshot fehlt. Betreiber muss Vorlagen einmal veröffentlichen.'
                : text || 'HTTP ' + res.status
        );
        err.status = res.status === 404 ? 404 : res.status;
        err.code = res.status === 404 ? 'snapshot_missing' : 'snapshot_read_failed';
        throw err;
    }
    try {
        return JSON.parse(text);
    } catch {
        const err = new Error('Snapshot-Datei ist kein gültiges JSON.');
        err.status = 500;
        err.code = 'snapshot_corrupt';
        throw err;
    }
}

/**
 * @param {string} relPath
 * @param {unknown} data
 */
async function writeJsonFile(relPath, data) {
    if (!isSafeSnapshotRel(relPath)) {
        const err = new Error('Ungültiger Snapshot-Pfad.');
        err.status = 400;
        throw err;
    }
    const body = Buffer.from(JSON.stringify(data, null, 0), 'utf8');
    if (body.length > MAX_SNAPSHOT_BYTES) {
        const err = new Error(
            'Snapshot-Datei zu groß (max. 12 MB): ' + relPath + ' (' + Math.round(body.length / 1024) + ' KB).'
        );
        err.status = 413;
        throw err;
    }
    const token = await getOperatorGraphToken();
    const drive = await resolveDrive(token);

    // Ensure parent folders exist
    const parts = relPath.split('/').filter(Boolean);
    parts.pop();
    let current = '';
    for (const part of parts) {
        current = current ? current + '/' + part : part;
        try {
            await graphJson(
                'GET',
                GRAPH +
                    '/drives/' +
                    encodeURIComponent(drive.driveId) +
                    '/root:/' +
                    encodeDrivePath(current),
                token
            );
        } catch (e) {
            if (!(e && (e.status === 404 || /itemNotFound|nicht gefunden/i.test(String(e.message))))) {
                throw e;
            }
            const parent = current.includes('/') ? current.slice(0, current.lastIndexOf('/')) : '';
            const url = parent
                ? GRAPH +
                  '/drives/' +
                  encodeURIComponent(drive.driveId) +
                  '/root:/' +
                  encodeDrivePath(parent) +
                  ':/children'
                : GRAPH + '/drives/' + encodeURIComponent(drive.driveId) + '/root/children';
            try {
                await graphJson('POST', url, token, {
                    name: part,
                    folder: {},
                    '@microsoft.graph.conflictBehavior': 'fail'
                });
            } catch (createErr) {
                const msg = String(createErr && createErr.message ? createErr.message : createErr);
                if (!/nameAlreadyExists|already exists|conflict/i.test(msg)) throw createErr;
            }
        }
    }

    if (body.length <= SIMPLE_UPLOAD_MAX) {
        const res = await fetch(contentUrl(drive.driveId, relPath), {
            method: 'PUT',
            headers: {
                Authorization: 'Bearer ' + token,
                'Content-Type': 'application/json; charset=utf-8'
            },
            body
        });
        if (!res.ok) {
            const text = await res.text();
            const err = new Error(text || 'Snapshot schreiben fehlgeschlagen (HTTP ' + res.status + ').');
            err.status = res.status;
            throw err;
        }
        return { path: relPath, size: body.length };
    }

    // Große Dateien (Bilder im HTML): Upload-Session
    const session = await graphJson(
        'POST',
        GRAPH +
            '/drives/' +
            encodeURIComponent(drive.driveId) +
            '/root:/' +
            encodeDrivePath(relPath) +
            ':/createUploadSession',
        token,
        {
            item: {
                '@microsoft.graph.conflictBehavior': 'replace',
                name: parts.length ? relPath.split('/').pop() : relPath
            }
        }
    );
    const uploadUrl = session && session.uploadUrl;
    if (!uploadUrl) {
        const err = new Error('Upload-Session ohne URL.');
        err.status = 500;
        throw err;
    }
    const chunkSize = 320 * 1024 * 8; // 2.5 MiB, Vielfaches von 320 KiB
    let offset = 0;
    while (offset < body.length) {
        const end = Math.min(offset + chunkSize, body.length) - 1;
        const chunk = body.subarray(offset, end + 1);
        const putRes = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                'Content-Length': String(chunk.length),
                'Content-Range': 'bytes ' + offset + '-' + end + '/' + body.length
            },
            body: chunk
        });
        if (!putRes.ok && putRes.status !== 202) {
            const text = await putRes.text();
            const err = new Error(
                text || 'Snapshot-Upload fehlgeschlagen (HTTP ' + putRes.status + ') für ' + relPath
            );
            err.status = putRes.status;
            throw err;
        }
        offset = end + 1;
    }
    return { path: relPath, size: body.length };
}

function sectionMetaPath(sectionId) {
    return SNAPSHOT_ROOT + '/sections/' + fileId(sectionId) + '.json';
}

function sectionPagePath(sectionId, pageId) {
    return (
        SNAPSHOT_ROOT +
        '/sections/' +
        fileId(sectionId) +
        '/pages/' +
        fileId(pageId) +
        '.json'
    );
}

/**
 * @param {string} sectionId
 * @param {{ id: string, title?: string }} pageMeta
 */
async function loadPageHtml(sectionId, pageMeta) {
    const pid = String((pageMeta && pageMeta.id) || '').trim();
    if (!pid) return { html: '', error: 'Seiten-ID fehlt' };
    // 1) Inline im Abschnitts-Meta (alte Snapshots)
    if (pageMeta && pageMeta.html) {
        return { html: String(pageMeta.html), error: undefined };
    }
    // 2) Einzeldatei
    try {
        const data = await readJsonFile(sectionPagePath(sectionId, pid));
        const html = String((data && data.html) || '');
        return {
            html,
            error: html ? undefined : 'Kein HTML im Snapshot (leere Seitendatei)'
        };
    } catch (e) {
        if (e && (e.status === 404 || e.code === 'snapshot_missing')) {
            return {
                html: '',
                error:
                    'Kein HTML im Snapshot – Abschnitt in kurtrocks erneut veröffentlichen (Seiteninhalt fehlte oder Upload zu groß).'
            };
        }
        return { html: '', error: (e && e.message) || String(e) };
    }
}

/**
 * @returns {Promise<{ site: object, notebooks: object[], preferredName: string, via: string, updatedAt?: string }>}
 */
async function listSnapshotNotebooks() {
    const cfg = getConfig();
    const index = await readJsonFile(SNAPSHOT_ROOT + '/index.json');
    const notebooks = (Array.isArray(index.notebooks) ? index.notebooks : [])
        .map((n) => ({
            id: String(n.id || ''),
            displayName: String(n.displayName || ''),
            isDefault: !!n.isDefault,
            userRole: 'Reader',
            lastModifiedDateTime: String(n.lastModifiedDateTime || index.updatedAt || '').trim(),
            lastModifiedByName: String(n.lastModifiedByName || index.publishedBy || '').trim(),
            publishedAt: String(n.publishedAt || index.updatedAt || '').trim(),
            publishedBy: String(n.publishedBy || index.publishedBy || '').trim()
        }))
        .filter((n) => n.id);
    const prefer = preferredNotebookName().toLowerCase();
    notebooks.sort((a, b) => {
        const ap = a.displayName.toLowerCase() === prefer ? 0 : 1;
        const bp = b.displayName.toLowerCase() === prefer ? 0 : 1;
        if (ap !== bp) return ap - bp;
        return a.displayName.localeCompare(b.displayName, 'de');
    });
    return {
        site: {
            id: '',
            displayName: 'MS365-Katalog (Snapshot)',
            webUrl: cfg.siteWebUrl
        },
        notebooks,
        preferredName: preferredNotebookName(),
        libraryHint: cfg.catalogLibraryName + '/' + SNAPSHOT_ROOT,
        siteWebUrl: cfg.siteWebUrl,
        updatedAt: index.updatedAt || null,
        via: 'catalog-snapshot'
    };
}

/**
 * @param {string} notebookId
 */
async function loadSnapshotNotebookTree(notebookId) {
    const nid = String(notebookId || '').trim();
    if (!nid) {
        const err = new Error('notebookId fehlt.');
        err.status = 400;
        throw err;
    }
    const tree = await readJsonFile(SNAPSHOT_ROOT + '/notebooks/' + fileId(nid) + '/tree.json');
    return {
        notebookId: nid,
        siteId: '',
        sections: Array.isArray(tree.sections) ? tree.sections : [],
        groups: Array.isArray(tree.groups) ? tree.groups : [],
        via: 'catalog-snapshot'
    };
}

/**
 * @param {string} sectionId
 */
async function listSnapshotSectionPages(sectionId) {
    const sid = String(sectionId || '').trim();
    if (!sid) {
        const err = new Error('sectionId fehlt.');
        err.status = 400;
        throw err;
    }
    const data = await readJsonFile(sectionMetaPath(sid));
    const pages = (Array.isArray(data.pages) ? data.pages : []).map((p) => ({
        id: String(p.id || ''),
        title: String(p.title || 'Ohne Titel'),
        webUrl: String(p.webUrl || '')
    }));
    return { sectionId: sid, pages, via: 'catalog-snapshot' };
}

/**
 * @param {string} pageId
 */
async function getSnapshotPageContent(pageId) {
    const pid = String(pageId || '').trim();
    if (!pid) {
        const err = new Error('pageId fehlt.');
        err.status = 400;
        throw err;
    }
    // pageId format in snapshot: sectionId::pageId or we search – store pages under section files only.
    // Callers use section export for distribute; for preview we need page lookup.
    const err = new Error('Seitenvorschau aus Snapshot: bitte Abschnitt wählen und Seite in der Liste öffnen.');
    err.status = 400;
    err.code = 'snapshot_page_lookup';
    throw err;
}

/**
 * Find page HTML inside section files via index map or scan.
 * Snapshot section JSON contains full page HTML.
 * @param {string} pageId
 * @param {string} [sectionId]
 */
async function getSnapshotPageById(pageId, sectionId) {
    const pid = String(pageId || '').trim();
    let sid = String(sectionId || '').trim();
    if (!sid) {
        try {
            const index = await readJsonFile(SNAPSHOT_ROOT + '/index.json');
            const map =
                index.pageToSection && typeof index.pageToSection === 'object' ? index.pageToSection : {};
            sid = map[pid] ? String(map[pid]) : '';
        } catch {
            sid = '';
        }
    }
    if (!sid) {
        const err = new Error('Seite im Snapshot nicht gefunden.');
        err.status = 404;
        err.code = 'snapshot_page_missing';
        throw err;
    }
    const data = await readJsonFile(sectionMetaPath(sid));
    const page = (Array.isArray(data.pages) ? data.pages : []).find((p) => String(p.id) === pid);
    if (!page) {
        const err = new Error('Seite im Snapshot nicht gefunden.');
        err.status = 404;
        err.code = 'snapshot_page_missing';
        throw err;
    }
    const loaded = await loadPageHtml(sid, page);
    if (!loaded.html) {
        const err = new Error(loaded.error || 'Kein HTML im Snapshot');
        err.status = 404;
        err.code = 'snapshot_page_empty';
        throw err;
    }
    return {
        pageId: pid,
        html: loaded.html,
        contentType: 'text/html',
        previewText: String(page.previewText || '').slice(0, 400),
        via: 'catalog-snapshot'
    };
}

/**
 * @param {string} sectionId
 */
async function getSnapshotSectionExport(sectionId) {
    const sid = String(sectionId || '').trim();
    if (!sid) {
        const err = new Error('sectionId fehlt.');
        err.status = 400;
        throw err;
    }
    const data = await readJsonFile(sectionMetaPath(sid));
    const metaPages = Array.isArray(data.pages) ? data.pages : [];
    const pages = [];
    for (const p of metaPages) {
        const loaded = await loadPageHtml(sid, p);
        pages.push({
            id: String(p.id || ''),
            title: String(p.title || 'Ohne Titel'),
            html: loaded.html,
            error: loaded.error
        });
    }
    return {
        sectionId: sid,
        displayName: String(data.displayName || ''),
        pageCount: pages.length,
        pages,
        via: 'catalog-snapshot'
    };
}

/**
 * Snapshot schreiben. notebooks/trees optional bei Abschnitts-Nachzug.
 * @param {{
 *   notebooks?: Array<{ id: string, displayName: string }>,
 *   trees?: Record<string, { sections: Array, groups: Array }>,
 *   sections?: Record<string, { displayName: string, pages: Array }>,
 *   publishedBy?: string
 * }} payload
 */
async function publishSnapshot(payload) {
    const notebooks = Array.isArray(payload && payload.notebooks) ? payload.notebooks : [];
    const trees = payload && payload.trees && typeof payload.trees === 'object' ? payload.trees : {};
    const sections =
        payload && payload.sections && typeof payload.sections === 'object' ? payload.sections : {};
    if (!notebooks.length && !Object.keys(sections).length) {
        const err = new Error('Snapshot ohne Notizbücher und Abschnitte.');
        err.status = 400;
        throw err;
    }

    const written = [];
    let pagesWithHtml = 0;
    let pagesWithoutHtml = 0;

    let existing = { notebooks: [], pageToSection: {}, publishedBy: '' };
    try {
        existing = await readJsonFile(SNAPSHOT_ROOT + '/index.json');
    } catch (e) {
        if (!(e && (e.status === 404 || e.code === 'snapshot_missing'))) throw e;
    }
    const mergedById = new Map();
    (Array.isArray(existing.notebooks) ? existing.notebooks : []).forEach((n) => {
        const id = String(n.id || '').trim();
        if (id) mergedById.set(id, n);
    });
    const mergedPageMap =
        existing.pageToSection && typeof existing.pageToSection === 'object'
            ? Object.assign({}, existing.pageToSection)
            : {};

    const publishedAt = new Date().toISOString();
    const publishedBy = String((payload && payload.publishedBy) || '').trim();

    for (const nb of notebooks) {
        const nid = String(nb.id || '').trim();
        if (!nid) continue;
        if (Object.prototype.hasOwnProperty.call(trees, nid)) {
            const tree = trees[nid] || { sections: [], groups: [] };
            written.push(
                await writeJsonFile(SNAPSHOT_ROOT + '/notebooks/' + fileId(nid) + '/tree.json', {
                    notebookId: nid,
                    displayName: String(nb.displayName || ''),
                    sections: tree.sections || [],
                    groups: tree.groups || []
                })
            );
        }
        const prev = mergedById.get(nid) || {};
        mergedById.set(nid, {
            id: nid,
            displayName: String(nb.displayName || prev.displayName || ''),
            lastModifiedDateTime: String(
                nb.lastModifiedDateTime || prev.lastModifiedDateTime || ''
            ).trim(),
            lastModifiedByName: String(
                nb.lastModifiedByName || prev.lastModifiedByName || ''
            ).trim(),
            publishedAt,
            publishedBy: publishedBy || String(prev.publishedBy || '').trim()
        });
    }

    for (const [sid, sec] of Object.entries(sections)) {
        const sectionId = String(sid || '').trim();
        if (!sectionId) continue;
        const pages = Array.isArray(sec.pages) ? sec.pages : [];
        const metaPages = [];

        for (const p of pages) {
            if (!p || !p.id) continue;
            const pageId = String(p.id);
            const html = String(p.html || '');
            mergedPageMap[pageId] = sectionId;

            const storeExternal = html.length > INLINE_HTML_MAX;
            if (html) {
                pagesWithHtml++;
                if (storeExternal) {
                    written.push(
                        await writeJsonFile(sectionPagePath(sectionId, pageId), {
                            id: pageId,
                            title: String(p.title || 'Ohne Titel'),
                            html
                        })
                    );
                }
            } else {
                pagesWithoutHtml++;
            }

            metaPages.push({
                id: pageId,
                title: String(p.title || 'Ohne Titel'),
                html: storeExternal ? '' : html,
                htmlExternal: !!storeExternal,
                previewText: String(p.previewText || '').slice(0, 500),
                webUrl: String(p.webUrl || ''),
                media: p.media || undefined
            });
        }

        written.push(
            await writeJsonFile(sectionMetaPath(sectionId), {
                sectionId,
                displayName: String(sec.displayName || ''),
                pages: metaPages
            })
        );
    }

    const mergedNotebooks = Array.from(mergedById.values());
    const index = {
        updatedAt: publishedAt,
        preferredName: preferredNotebookName(),
        publishedBy: publishedBy || String(existing.publishedBy || '').trim(),
        notebooks: mergedNotebooks.length
            ? mergedNotebooks
            : Array.isArray(existing.notebooks)
              ? existing.notebooks
              : [],
        pageToSection: mergedPageMap
    };
    written.push(await writeJsonFile(SNAPSHOT_ROOT + '/index.json', index));
    clearCatalogCache();
    driveCache = null;

    return {
        ok: true,
        via: 'catalog-snapshot',
        updatedAt: index.updatedAt,
        notebookCount: index.notebooks.length,
        sectionCount: Object.keys(sections).length,
        pagesWithHtml,
        pagesWithoutHtml,
        files: written.length,
        merged: true
    };
}

function isOneNoteAppOnlyBlocked(err) {
    const msg = String((err && err.message) || err || '');
    const code = err && err.payload && err.payload.error && err.payload.error.code;
    return (
        String(code) === '40001' ||
        /no longer support app-only|app-only tokens starting from March 31/i.test(msg)
    );
}

function appOnlyBlockedError() {
    const err = new Error(
        'Kein OneNote-Snapshot im Katalog gefunden. In kurtrocks: Zentrale Vorlagen laden → „Für Schulen veröffentlichen“ abwarten (Toast „Snapshot veröffentlicht“), dann im Schul-Tenant erneut laden.'
    );
    err.status = 503;
    err.code = 'onenote_app_only_blocked';
    return err;
}

module.exports = {
    SNAPSHOT_ROOT,
    listSnapshotNotebooks,
    loadSnapshotNotebookTree,
    listSnapshotSectionPages,
    getSnapshotPageById,
    getSnapshotSectionExport,
    publishSnapshot,
    isOneNoteAppOnlyBlocked,
    appOnlyBlockedError,
    preferredNotebookName
};
