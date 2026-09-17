/**
 * Graph-API: Stammdaten-Backup in die IT-Dokumentbibliothek schreiben / lesen.
 * Kein Seiten-DOM – nutzt ms365SpoGraph + ms365BrowserBackup.
 */
import {
    DEFAULT_FOLDER,
    CURRENT_FILE,
    buildDriveRelativePath,
    encodeDriveRootPath,
    describeRemoteBackup,
    isItLibraryConfigured
} from './stammdaten-sharepoint-sync-logic.js';

const SCOPES_GRAPH = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/Group.Read.All'
];

const IT_META_KEY = 'ms365-stammdaten-it-library-v1';
const SYNC_META_KEY = 'ms365-stammdaten-spo-sync-v1';

function getG() {
    const G = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!G) throw new Error('SharePoint-Graph-Helfer nicht geladen (spo-graph-shared.js).');
    return G;
}

export function loadItMeta() {
    try {
        const setup =
            window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup()
                : null;
        if (setup && setup.stammdatenItLibrary && typeof setup.stammdatenItLibrary === 'object') {
            return setup.stammdatenItLibrary;
        }
    } catch {
        /* ignore */
    }
    try {
        return JSON.parse(localStorage.getItem(IT_META_KEY) || '{}') || {};
    } catch {
        return {};
    }
}

export function saveItMeta(meta) {
    const m = meta || {};
    try {
        localStorage.setItem(IT_META_KEY, JSON.stringify(m));
    } catch {
        /* ignore */
    }
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup({ stammdatenItLibrary: m });
        }
    } catch {
        /* ignore */
    }
}

export function loadLocalSyncMeta() {
    try {
        return JSON.parse(localStorage.getItem(SYNC_META_KEY) || '{}') || {};
    } catch {
        return {};
    }
}

export function saveLocalSyncMeta(meta) {
    try {
        localStorage.setItem(SYNC_META_KEY, JSON.stringify(meta || {}));
    } catch {
        /* ignore */
    }
}

export function isReady() {
    return isItLibraryConfigured(loadItMeta());
}

/**
 * Relativer Pfad zur Übergabe-/Einrichtungsseite.
 * @param {string} [fromPath] document.location.pathname
 */
export function setupPageHref(fromPath) {
    const p = String(fromPath || (typeof location !== 'undefined' ? location.pathname : '') || '');
    const inTools = /\/tools\//i.test(p) || /\\tools\\/i.test(p);
    return (inTools ? 'stammdaten-uebergabe.html' : 'tools/stammdaten-uebergabe.html') + '#setup';
}

export function requireItLibrary() {
    const it = loadItMeta();
    if (!isItLibraryConfigured(it)) {
        const err = new Error('IT-Bibliothek noch nicht eingerichtet.');
        err.code = 'IT_LIBRARY_MISSING';
        throw err;
    }
    return it;
}

async function ensureGraphToken() {
    return getG().getGraphToken(SCOPES_GRAPH);
}

export async function putJsonOnDrive(driveId, relativePath, jsonText, token) {
    const G = getG();
    const enc = encodeDriveRootPath(relativePath);
    const url = G.graphBase('v1.0') + '/drives/' + encodeURIComponent(driveId) + '/' + enc + '/content';
    const res = await fetch(url, {
        method: 'PUT',
        headers: {
            Authorization: 'Bearer ' + token,
            'Content-Type': 'application/json; charset=utf-8'
        },
        body: jsonText
    });
    const text = await res.text();
    let data = null;
    try {
        data = text ? JSON.parse(text) : {};
    } catch {
        data = { raw: text };
    }
    if (!res.ok) {
        const msg =
            data && data.error && data.error.message ? data.error.message : text || String(res.status);
        throw new Error('Upload fehlgeschlagen: ' + msg);
    }
    return data;
}

export async function listDriveFolder(driveId, folder, token) {
    const G = getG();
    const rel = buildDriveRelativePath(folder, '');
    const enc = encodeDriveRootPath(rel.replace(/\/$/, '') || DEFAULT_FOLDER);
    const path =
        '/drives/' +
        encodeURIComponent(driveId) +
        '/' +
        enc +
        '/children?$select=id,name,size,lastModifiedDateTime,webUrl,file&$orderby=lastModifiedDateTime desc&$top=50';
    try {
        return await G.graphJson('GET', path, token, undefined, 'v1.0');
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e);
        if (/itemNotFound|404|not found/i.test(msg)) return { value: [] };
        throw e;
    }
}

export async function downloadDriveItem(driveId, itemId, token) {
    const G = getG();
    const url =
        G.graphBase('v1.0') +
        '/drives/' +
        encodeURIComponent(driveId) +
        '/items/' +
        encodeURIComponent(itemId) +
        '/content';
    const res = await fetch(url, { method: 'GET', headers: { Authorization: 'Bearer ' + token } });
    const text = await res.text();
    if (!res.ok) throw new Error('Download fehlgeschlagen: HTTP ' + res.status);
    return JSON.parse(text);
}

function buildPayloadJson() {
    const bb = window.ms365BrowserBackup;
    if (!bb || typeof bb.buildBackup !== 'function') throw new Error('Browser-Backup-Modul fehlt.');
    const payload = bb.buildBackup();
    return { payload: payload, text: JSON.stringify(payload, null, 2) };
}

/**
 * @param {{ folder?: string, keepDated?: boolean, siteUrl?: string }} [opts]
 */
export async function uploadCurrentBackup(opts) {
    const options = opts || {};
    const it = requireItLibrary();
    const folder = String(options.folder || DEFAULT_FOLDER).trim() || DEFAULT_FOLDER;
    const keepDated = !!options.keepDated;
    const built = buildPayloadJson();
    const token = await ensureGraphToken();
    const currentPath = buildDriveRelativePath(folder, CURRENT_FILE);
    const item = await putJsonOnDrive(it.driveId, currentPath, built.text, token);
    if (keepDated && window.ms365BrowserBackup && typeof window.ms365BrowserBackup.backupFilename === 'function') {
        const datedName = window.ms365BrowserBackup.backupFilename(new Date());
        await putJsonOnDrive(it.driveId, buildDriveRelativePath(folder, datedName), built.text, token);
    }
    const webUrl = (item && item.webUrl) || it.webUrl || '';
    const meta = {
        at: new Date().toISOString(),
        fileName: CURRENT_FILE,
        folder: folder,
        webUrl: webUrl,
        siteUrl: options.siteUrl || it.siteUrl || '',
        driveId: it.driveId,
        summary: describeRemoteBackup(built.payload)
    };
    saveLocalSyncMeta(meta);
    try {
        localStorage.setItem('ms365-last-backup-export-at', new Date().toISOString());
    } catch {
        /* ignore */
    }
    return { item: item, meta: meta, payload: built.payload };
}

/**
 * @param {{ folder?: string }} [opts]
 * @returns {Promise<{ id: string, name: string }>}
 */
export async function findCurrentBackupItem(opts) {
    const options = opts || {};
    const it = requireItLibrary();
    const folder = String(options.folder || DEFAULT_FOLDER).trim() || DEFAULT_FOLDER;
    const token = await ensureGraphToken();
    const data = await listDriveFolder(it.driveId, folder, token);
    const items = (data && data.value) || [];
    const cur = items.find(function (i) {
        return i && i.file && String(i.name || '') === CURRENT_FILE;
    });
    if (!cur) throw new Error('Datei „' + CURRENT_FILE + '“ nicht gefunden.');
    return cur;
}

/**
 * @param {{ folder?: string, apply?: boolean }} [opts]
 * apply=false: nur laden/prüfen, nicht importieren.
 */
export async function downloadCurrentBackup(opts) {
    const options = opts || {};
    const it = requireItLibrary();
    const token = await ensureGraphToken();
    const cur = await findCurrentBackupItem(options);
    const obj = await downloadDriveItem(it.driveId, cur.id, token);
    const bb = window.ms365BrowserBackup;
    if (!bb || typeof bb.isBackupPayload !== 'function') throw new Error('Backup-Modul fehlt.');
    if (!bb.isBackupPayload(obj) && !bb.isLegacyAppDataPayload(obj)) {
        throw new Error('Datei ist kein erkanntes MS365-Browser-Backup.');
    }
    if (options.apply !== false) {
        bb.importPayload(obj);
    }
    return { payload: obj, item: cur };
}

export default {
    SCOPES_GRAPH,
    loadItMeta,
    saveItMeta,
    loadLocalSyncMeta,
    saveLocalSyncMeta,
    isReady,
    setupPageHref,
    requireItLibrary,
    putJsonOnDrive,
    listDriveFolder,
    downloadDriveItem,
    uploadCurrentBackup,
    findCurrentBackupItem,
    downloadCurrentBackup
};
