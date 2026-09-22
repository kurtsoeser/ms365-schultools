'use strict';

const { getConfig } = require('./config');
const { getOperatorGraphToken } = require('./msal-app-only');
const { graphJson, resolveSiteId, resolveListId } = require('./sharepoint-license');
const { buildStoredCatalog, emptyCatalog } = require('./catalog-payload');

const GRAPH = 'https://graph.microsoft.com/v1.0';

/** @type {{ siteId: string, listId: string, driveId: string, webUrl: string, library: string, at: number } | null} */
let driveCache = null;
const DRIVE_CACHE_MS = 10 * 60 * 1000;

/**
 * @param {string} rel
 */
function encodeDrivePath(rel) {
    return String(rel || '')
        .replace(/\\/g, '/')
        .split('/')
        .filter(Boolean)
        .map((seg) => encodeURIComponent(seg))
        .join('/');
}

/**
 * @param {string} method
 * @param {string} url
 * @param {string} token
 * @param {string} [body]
 * @param {string} [contentType]
 */
async function graphText(method, url, token, body, contentType) {
    const headers = {
        Authorization: 'Bearer ' + token,
        Accept: 'application/json'
    };
    if (body !== undefined) {
        headers['Content-Type'] = contentType || 'application/octet-stream';
    }
    const res = await fetch(url, {
        method,
        headers,
        body
    });
    const text = Buffer.from(await res.arrayBuffer()).toString('utf8');
    if (!res.ok) {
        let msg = text || 'HTTP ' + res.status;
        try {
            const data = JSON.parse(text);
            if (data && data.error && data.error.message) msg = data.error.message;
        } catch {
            /* Rohtext behalten */
        }
        const err = new Error(msg);
        err.status = res.status;
        throw err;
    }
    return text;
}

/**
 * @param {string} token
 * @param {{ create?: boolean }} opts
 */
async function resolveCatalogDrive(token, opts) {
    const cfg = getConfig();
    const now = Date.now();
    if (
        driveCache &&
        now - driveCache.at < DRIVE_CACHE_MS &&
        driveCache.library === cfg.catalogLibraryName
    ) {
        return driveCache;
    }
    const site = await resolveSiteId(cfg.siteWebUrl, token);
    let list = null;
    try {
        list = await resolveListId(site.id, cfg.catalogLibraryName, token);
    } catch (e) {
        const missing = /nicht gefunden/i.test(String(e && e.message));
        if (!missing || !opts || !opts.create) throw e;
        list = await graphJson(
            'POST',
            GRAPH + '/sites/' + encodeURIComponent(site.id) + '/lists',
            token,
            {
                displayName: cfg.catalogLibraryName,
                description:
                    'Zentraler Katalog der MS365-Schultools (Vorlagen, später Materialien) für alle Schulen.',
                list: { template: 'documentLibrary' }
            }
        );
    }
    if (!list || !list.id) {
        throw new Error('Dokumentbibliothek „' + cfg.catalogLibraryName + '“ konnte nicht angelegt werden.');
    }
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
        throw new Error('„' + cfg.catalogLibraryName + '“ ist keine Dokumentbibliothek.');
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

/**
 * @param {string} driveId
 * @param {string} relPath
 */
function contentUrl(driveId, relPath) {
    return (
        GRAPH + '/drives/' + encodeURIComponent(driveId) + '/root:/' + encodeDrivePath(relPath) + ':/content'
    );
}

function childrenUrl(driveId, relPath) {
    return (
        GRAPH + '/drives/' + encodeURIComponent(driveId) + '/root:/' + encodeDrivePath(relPath) + ':/children'
    );
}

function itemMetaUrl(driveId, relPath) {
    return (
        GRAPH +
        '/drives/' +
        encodeURIComponent(driveId) +
        '/root:/' +
        encodeDrivePath(relPath) +
        '?$select=id,name,size,file,folder,webUrl'
    );
}

/**
 * Relativpfad nur unter der Materialien-Wurzel.
 * @param {unknown} raw
 * @param {string} root
 * @returns {string}
 */
function normalizeMaterialsRelPath(raw, root) {
    const base = String(root || 'materialien')
        .replace(/\\/g, '/')
        .replace(/^\/+|\/+$/g, '');
    if (!base) throw Object.assign(new Error('Materialien-Wurzel fehlt.'), { status: 500 });
    let path = String(raw == null ? '' : raw)
        .replace(/\\/g, '/')
        .replace(/^\/+|\/+$/g, '');
    if (!path) path = base;
    if (path === '.' || path.includes('..')) {
        const err = new Error('Ungültiger Material-Pfad.');
        err.status = 400;
        throw err;
    }
    const parts = path.split('/').filter(Boolean);
    if (!parts.length || parts[0].toLowerCase() !== base.toLowerCase()) {
        const err = new Error('Pfad muss unter „' + base + '/“ liegen.');
        err.status = 400;
        throw err;
    }
    if (parts.length > 12) {
        const err = new Error('Pfad ist zu tief.');
        err.status = 400;
        throw err;
    }
    for (const p of parts) {
        if (!p || p === '.' || p === '..') {
            const err = new Error('Ungültiger Material-Pfad.');
            err.status = 400;
            throw err;
        }
    }
    return parts.join('/');
}

/**
 * @param {string} [relPath]
 */
async function listMaterials(relPath) {
    const cfg = getConfig();
    const path = normalizeMaterialsRelPath(relPath, cfg.catalogMaterialsRoot);
    const token = await getOperatorGraphToken();
    let drive;
    try {
        drive = await resolveCatalogDrive(token, { create: false });
    } catch (e) {
        if (isMissingError(e)) {
            return {
                path,
                missing: true,
                message: 'Bibliothek „' + cfg.catalogLibraryName + '“ fehlt.',
                items: [],
                webUrl: ''
            };
        }
        throw e;
    }
    try {
        const page = await graphJson(
            'GET',
            childrenUrl(drive.driveId, path) +
                '?$select=id,name,size,file,folder,webUrl&$orderby=name&$top=200',
            token
        );
        const rows = page.value || [];
        const items = rows.map((row) => {
            const name = String(row.name || '').trim();
            const childPath = path + '/' + name;
            const isFolder = !!(row.folder && typeof row.folder === 'object');
            return {
                id: String(row.id || ''),
                name,
                path: childPath,
                isFolder,
                size: Number(row.size) || 0,
                mimeType:
                    row.file && row.file.mimeType ? String(row.file.mimeType) : isFolder ? '' : 'application/octet-stream',
                webUrl: row.webUrl || ''
            };
        });
        return {
            path,
            missing: false,
            message: '',
            items,
            webUrl: drive.webUrl || '',
            library: cfg.catalogLibraryName
        };
    } catch (e) {
        if (isMissingError(e)) {
            return {
                path,
                missing: true,
                message:
                    'Ordner „' +
                    path +
                    '“ fehlt. In der Bibliothek „' +
                    cfg.catalogLibraryName +
                    '“ anlegen und Test-Dateien ablegen.',
                items: [],
                webUrl: drive.webUrl || '',
                library: cfg.catalogLibraryName
            };
        }
        throw e;
    }
}

const MAX_MATERIAL_BYTES = 8 * 1024 * 1024;

/**
 * @param {string} relPath
 * @returns {Promise<{ name: string, path: string, contentType: string, bytes: Buffer }>}
 */
async function readMaterialFile(relPath) {
    const cfg = getConfig();
    const path = normalizeMaterialsRelPath(relPath, cfg.catalogMaterialsRoot);
    if (path.toLowerCase() === String(cfg.catalogMaterialsRoot).toLowerCase()) {
        const err = new Error('Bitte eine Datei angeben, nicht den Wurzelordner.');
        err.status = 400;
        throw err;
    }
    const token = await getOperatorGraphToken();
    const drive = await resolveCatalogDrive(token, { create: false });
    const meta = await graphJson('GET', itemMetaUrl(drive.driveId, path), token);
    if (meta.folder) {
        const err = new Error('„' + path + '“ ist ein Ordner, keine Datei.');
        err.status = 400;
        throw err;
    }
    const size = Number(meta.size) || 0;
    if (size > MAX_MATERIAL_BYTES) {
        const err = new Error('Datei ist größer als 8 MB (Test-Limit).');
        err.status = 413;
        throw err;
    }
    const headers = {
        Authorization: 'Bearer ' + token
    };
    const res = await fetch(contentUrl(drive.driveId, path), { method: 'GET', headers });
    if (!res.ok) {
        const text = await res.text();
        let msg = text || 'HTTP ' + res.status;
        try {
            const data = JSON.parse(text);
            if (data && data.error && data.error.message) msg = data.error.message;
        } catch {
            /* Rohtext */
        }
        const err = new Error(msg);
        err.status = res.status;
        throw err;
    }
    const bytes = Buffer.from(await res.arrayBuffer());
    const contentType =
        (meta.file && meta.file.mimeType) ||
        res.headers.get('content-type') ||
        'application/octet-stream';
    return {
        name: String(meta.name || path.split('/').pop() || 'datei'),
        path,
        contentType: String(contentType).split(';')[0].trim() || 'application/octet-stream',
        bytes
    };
}

function clearCatalogCache() {
    driveCache = null;
}

module.exports = {
    readKursteamCatalog,
    writeKursteamCatalog,
    listMaterials,
    readMaterialFile,
    clearCatalogCache,
    encodeDrivePath,
    normalizeMaterialsRelPath,
    MAX_MATERIAL_BYTES
};

/**
 * Zentrale Kursteam-Vorlagen lesen. Fehlende Bibliothek/Datei ist kein Fehler.
 */
async function readKursteamCatalog() {
    const cfg = getConfig();
    const token = await getOperatorGraphToken();
    let drive;
    try {
        drive = await resolveCatalogDrive(token, { create: false });
    } catch (e) {
        if (isMissingError(e)) {
            return emptyCatalog(cfg, {
                missing: true,
                message: 'Bibliothek „' + cfg.catalogLibraryName + '“ ist noch nicht angelegt.'
            });
        }
        throw e;
    }
    try {
        let text = await graphText('GET', contentUrl(drive.driveId, cfg.catalogKursteamPath), token);
        if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
        const parsed = JSON.parse(text);
        const stored = buildStoredCatalog(parsed, {
            strict: false,
            updatedBy: parsed && parsed.updatedBy,
            updatedAt: parsed && parsed.updatedAt
        });
        return Object.assign(stored, {
            missing: false,
            message: '',
            library: cfg.catalogLibraryName,
            path: cfg.catalogKursteamPath,
            siteWebUrl: cfg.siteWebUrl,
            webUrl: drive.webUrl || ''
        });
    } catch (e) {
        if (e instanceof SyntaxError) {
            const err = new Error('Die zentrale Vorlagendatei ist kein gültiges JSON.');
            err.status = 502;
            throw err;
        }
        if (isMissingError(e)) {
            return emptyCatalog(cfg, {
                missing: true,
                message: 'Noch keine zentrale Vorlagendatei.'
            });
        }
        throw e;
    }
}

/**
 * Lokale Bibliothek des Betreibers als zentrale Datei speichern.
 * Legt die Dokumentbibliothek bei Bedarf an.
 * @param {unknown} body
 * @param {string} updatedBy
 */
async function writeKursteamCatalog(body, updatedBy) {
    const cfg = getConfig();
    const stored = buildStoredCatalog(body, {
        strict: true,
        updatedBy: updatedBy || '',
        updatedAt: new Date().toISOString()
    });
    const json = JSON.stringify(stored, null, 2);
    if (json.length > 2_000_000) {
        const err = new Error('Katalog ist zu groß (max. 2 MB).');
        err.status = 413;
        throw err;
    }
    const token = await getOperatorGraphToken();
    const drive = await resolveCatalogDrive(token, { create: true });
    await graphText(
        'PUT',
        contentUrl(drive.driveId, cfg.catalogKursteamPath),
        token,
        json,
        'application/json'
    );
    return Object.assign(stored, {
        missing: false,
        message: '',
        library: cfg.catalogLibraryName,
        path: cfg.catalogKursteamPath,
        siteWebUrl: cfg.siteWebUrl,
        webUrl: drive.webUrl || ''
    });
}

function clearCatalogCache() {
    driveCache = null;
}

module.exports = {
    readKursteamCatalog,
    writeKursteamCatalog,
    clearCatalogCache,
    encodeDrivePath
};
