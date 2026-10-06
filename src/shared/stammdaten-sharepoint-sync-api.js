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
    isItLibraryConfigured,
    normalizeItLibraryMeta
} from './stammdaten-sharepoint-sync-logic.js';
import {
    buildConfigBundleFromBackup,
    isConfigBundleManifest,
    parseConfigBundlePart,
    mergeConfigPartsLocalStorage
} from './stammdaten-sharepoint-config-bundle.js';

const SCOPES_GRAPH = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/Group.Read.All'
];

/** Legacy (ein Mandant pro Browser); wird in Mandanten-Map migriert. */
const IT_META_KEY = 'ms365-stammdaten-it-library-v1';
const SYNC_META_KEY = 'ms365-stammdaten-spo-sync-v1';
const IT_META_BY_TENANT_KEY = 'ms365-stammdaten-it-library-by-tenant-v2';
const SYNC_META_BY_TENANT_KEY = 'ms365-stammdaten-spo-sync-by-tenant-v2';
const FORM_DRAFT_KEY = 'ms365-su-form-draft-v1';

function getG() {
    const G = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!G) throw new Error('SharePoint-Graph-Helfer nicht geladen (spo-graph-shared.js).');
    return G;
}

function getTenantIdSync() {
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            if (info && info.tenantId) return String(info.tenantId).trim();
        }
    } catch {
        /* ignore */
    }
    return '';
}

function readJsonMap(storageKey) {
    try {
        const raw = localStorage.getItem(storageKey);
        const parsed = raw ? JSON.parse(raw) : {};
        return parsed && typeof parsed === 'object' ? parsed : {};
    } catch {
        return {};
    }
}

function writeJsonMap(storageKey, map) {
    try {
        localStorage.setItem(storageKey, JSON.stringify(map || {}));
    } catch {
        /* ignore */
    }
}

function readLegacyItMeta() {
    try {
        return normalizeItLibraryMeta(JSON.parse(localStorage.getItem(IT_META_KEY) || '{}') || {});
    } catch {
        return {};
    }
}

function readSetupItMeta() {
    try {
        const setup =
            window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup()
                : null;
        if (setup && setup.stammdatenItLibrary && typeof setup.stammdatenItLibrary === 'object') {
            return normalizeItLibraryMeta(setup.stammdatenItLibrary);
        }
    } catch {
        /* ignore */
    }
    return {};
}

export function readItLibraryFormDraft() {
    try {
        const raw = localStorage.getItem(FORM_DRAFT_KEY);
        if (!raw) return {};
        const parsed = JSON.parse(raw);
        return parsed && typeof parsed === 'object' ? parsed : {};
    } catch {
        return {};
    }
}

/**
 * @param {Partial<{ siteUrl: string, libraryTitle: string, itGroup: string, folder: string, keepDated: boolean }>} patch
 */
export function writeItLibraryFormDraft(patch) {
    const cur = readItLibraryFormDraft();
    const next = Object.assign({}, cur, patch || {});
    try {
        localStorage.setItem(FORM_DRAFT_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    try {
        sessionStorage.setItem(FORM_DRAFT_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

function mergeItMetaParts() {
    const parts = [readLegacyItMeta(), readSetupItMeta()];
    const out = {};
    parts.forEach(function (p) {
        Object.keys(p).forEach(function (k) {
            const v = p[k];
            if (v != null && v !== '') out[k] = v;
        });
    });
    return normalizeItLibraryMeta(out);
}

function loadItMetaFromTenantStore(tenantId) {
    if (!tenantId) return null;
    const map = readJsonMap(IT_META_BY_TENANT_KEY);
    const entry = map[tenantId];
    if (!entry || typeof entry !== 'object') return null;
    const meta = normalizeItLibraryMeta(entry);
    if (!Object.keys(meta).some(function (k) {
        return meta[k] != null && meta[k] !== '' && meta[k] !== false;
    })) {
        return null;
    }
    return meta;
}

function persistItMetaToTenantStore(tenantId, meta) {
    if (!tenantId) return;
    const map = readJsonMap(IT_META_BY_TENANT_KEY);
    map[tenantId] = normalizeItLibraryMeta(meta);
    writeJsonMap(IT_META_BY_TENANT_KEY, map);
}

function loadSyncMetaFromTenantStore(tenantId) {
    if (!tenantId) return null;
    const map = readJsonMap(SYNC_META_BY_TENANT_KEY);
    const entry = map[tenantId];
    return entry && typeof entry === 'object' ? entry : null;
}

function persistSyncMetaToTenantStore(tenantId, meta) {
    if (!tenantId) return;
    const map = readJsonMap(SYNC_META_BY_TENANT_KEY);
    map[tenantId] = meta || {};
    writeJsonMap(SYNC_META_BY_TENANT_KEY, map);
}

export function loadItMeta() {
    const tenantId = getTenantIdSync();
    let meta = loadItMetaFromTenantStore(tenantId);
    if (meta && isItLibraryConfigured(meta)) {
        return meta;
    }

    const migrated = mergeItMetaParts();
    if (isItLibraryConfigured(migrated)) {
        if (tenantId) persistItMetaToTenantStore(tenantId, migrated);
        return migrated;
    }

    if (
        meta &&
        Object.keys(meta).some(function (k) {
            return meta[k] != null && meta[k] !== '' && meta[k] !== false;
        })
    ) {
        return meta;
    }
    if (Object.keys(migrated).length) return migrated;
    return {};
}

export function saveItMeta(meta) {
    const m = normalizeItLibraryMeta(meta || {});
    const tenantId = getTenantIdSync();
    if (tenantId) {
        persistItMetaToTenantStore(tenantId, m);
    }
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
    const tenantId = getTenantIdSync();
    const fromTenant = loadSyncMetaFromTenantStore(tenantId);
    if (fromTenant) return fromTenant;
    try {
        return JSON.parse(localStorage.getItem(SYNC_META_KEY) || '{}') || {};
    } catch {
        return {};
    }
}

export function saveLocalSyncMeta(meta) {
    const m = meta || {};
    const tenantId = getTenantIdSync();
    if (tenantId) {
        persistSyncMetaToTenantStore(tenantId, m);
    }
    try {
        localStorage.setItem(SYNC_META_KEY, JSON.stringify(m));
    } catch {
        /* ignore */
    }
}

/**
 * Lokale Änderungen als „noch nicht auf SharePoint“ markieren.
 * @param {boolean} [dirty=true]
 */
export function setLocalDirty(dirty) {
    const cur = loadLocalSyncMeta();
    const next = Object.assign({}, cur, {
        dirty: dirty !== false,
        pendingError: dirty === false ? null : cur.pendingError || null
    });
    saveLocalSyncMeta(next);
    return next;
}

export function isLocalDirty() {
    const m = loadLocalSyncMeta();
    return !!(m && m.dirty);
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

export async function getJsonAtDrivePath(driveId, relativePath, token) {
    const G = getG();
    const enc = encodeDriveRootPath(relativePath);
    const url = G.graphBase('v1.0') + '/drives/' + encodeURIComponent(driveId) + '/' + enc + '/content';
    const res = await fetch(url, { method: 'GET', headers: { Authorization: 'Bearer ' + token } });
    if (res.status === 404) return null;
    const text = await res.text();
    if (!res.ok) {
        throw new Error('Download fehlgeschlagen (' + relativePath + '): HTTP ' + res.status);
    }
    if (!text) return null;
    try {
        return JSON.parse(text);
    } catch {
        throw new Error('Ungültiges JSON: ' + relativePath);
    }
}

/**
 * @param {object} payload Vollbackup
 * @param {{ tenantId?: string, backupsFolder?: string }} [opts]
 */
export async function uploadConfigBundle(payload, opts) {
    const options = opts || {};
    const it = requireItLibrary();
    const token = await ensureGraphToken();
    const backupsFolder = String(options.backupsFolder || DEFAULT_FOLDER).trim() || DEFAULT_FOLDER;
    const monolithPath = buildDriveRelativePath(backupsFolder, CURRENT_FILE);
    const bundle = buildConfigBundleFromBackup(payload, {
        tenantId: options.tenantId || getTenantIdSync(),
        monolithPath: monolithPath
    });
    for (let i = 0; i < bundle.files.length; i++) {
        const f = bundle.files[i];
        await putJsonOnDrive(it.driveId, f.path, JSON.stringify(f.body, null, 2), token);
    }
    await putJsonOnDrive(
        it.driveId,
        bundle.manifestPath,
        JSON.stringify(bundle.manifest, null, 2),
        token
    );
    return bundle;
}

/**
 * @returns {Promise<{ manifest: object|null, parts: object[], manifestPath: string }>}
 */
export async function downloadConfigBundle(opts) {
    const options = opts || {};
    const it = requireItLibrary();
    const token = await ensureGraphToken();
    const bundle = buildConfigBundleFromBackup({ localStorage: {} });
    const manifestPath = bundle.manifestPath;
    const manifest = await getJsonAtDrivePath(it.driveId, manifestPath, token);
    if (!isConfigBundleManifest(manifest)) {
        return { manifest: null, parts: [], manifestPath: manifestPath };
    }
    /** @type {object[]} */
    const parts = [];
    const files = Array.isArray(manifest.files) ? manifest.files : [];
    for (let i = 0; i < files.length; i++) {
        const entry = files[i] || {};
        const path = String(entry.path || '').trim();
        const id = String(entry.id || '').trim();
        if (!path || !id) continue;
        const raw = await getJsonAtDrivePath(it.driveId, path, token);
        const parsed = parseConfigBundlePart(raw, id);
        if (parsed) parts.push(parsed);
    }
    return { manifest: manifest, parts: parts, manifestPath: manifestPath };
}

/**
 * @param {{ reload?: boolean, syncMeta?: object }} [opts]
 */
export async function applyConfigBundleFromSharePoint(opts) {
    const options = opts || {};
    const downloaded = await downloadConfigBundle();
    if (!downloaded.manifest || !downloaded.parts.length) {
        return { applied: false, reason: 'no-manifest' };
    }
    const patch = mergeConfigPartsLocalStorage(downloaded.parts);
    const bb = window.ms365BrowserBackup;
    if (!bb || typeof bb.mergeLocalStoragePatch !== 'function') {
        throw new Error('Backup-Modul fehlt (mergeLocalStoragePatch).');
    }
    const result = bb.mergeLocalStoragePatch(patch);
    const meta =
        options.syncMeta ||
        Object.assign({}, loadLocalSyncMeta(), {
            at: new Date().toISOString(),
            direction: 'pull',
            dirty: false,
            pendingError: null,
            configBundleAppliedAt: new Date().toISOString(),
            configManifestFingerprint: String(downloaded.manifest.contentFingerprint || ''),
            configPartCount: downloaded.parts.length
        });
    saveLocalSyncMeta(meta);
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-tenant-settings-changed', {
                detail: { source: 'spo-config-bundle-import' }
            })
        );
    } catch {
        /* ignore */
    }
    if (options.reload !== false) {
        window.setTimeout(function () {
            window.location.reload();
        }, 80);
    }
    return { applied: true, manifest: downloaded.manifest, result: result, meta: meta };
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
        '/children?$select=id,name,size,lastModifiedDateTime,eTag,cTag,webUrl,file&$orderby=lastModifiedDateTime desc&$top=50';
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
 * keepDated: Standard true – zusätzlich zur aktuellen Datei eine datierte Kopie im gleichen Ordner.
 */
export async function uploadCurrentBackup(opts) {
    const options = opts || {};
    const it = requireItLibrary();
    const folder = String(options.folder || DEFAULT_FOLDER).trim() || DEFAULT_FOLDER;
    const keepDated = options.keepDated !== false;
    const built = buildPayloadJson();
    const token = await ensureGraphToken();
    const currentPath = buildDriveRelativePath(folder, CURRENT_FILE);
    const item = await putJsonOnDrive(it.driveId, currentPath, built.text, token);
    /** @type {string} */
    let configManifestFingerprint = '';
    try {
        const bundle = await uploadConfigBundle(built.payload, {
            tenantId: getTenantIdSync(),
            backupsFolder: folder
        });
        configManifestFingerprint = String(
            (bundle && bundle.manifest && bundle.manifest.contentFingerprint) || ''
        );
    } catch (e) {
        /* Config-Bundle optional – Monolith bleibt maßgeblich */
        if (typeof console !== 'undefined' && console.warn) {
            console.warn('Config-Bundle Upload:', e && e.message ? e.message : e);
        }
    }
    if (keepDated && window.ms365BrowserBackup && typeof window.ms365BrowserBackup.backupFilename === 'function') {
        const datedName = window.ms365BrowserBackup.backupFilename(new Date());
        await putJsonOnDrive(it.driveId, buildDriveRelativePath(folder, datedName), built.text, token);
    }
    const webUrl = (item && item.webUrl) || it.webUrl || '';
    const meta = {
        at: new Date().toISOString(),
        direction: 'push',
        fileName: CURRENT_FILE,
        folder: folder,
        webUrl: webUrl,
        siteUrl: options.siteUrl || it.siteUrl || '',
        driveId: it.driveId,
        summary: describeRemoteBackup(built.payload),
        remoteETag: (item && (item.eTag || item.cTag)) || '',
        remoteLastModified: (item && item.lastModifiedDateTime) || '',
        remoteExportedAt: (built.payload && built.payload.exportedAt) || '',
        contentFingerprint: (built.payload && built.payload.contentFingerprint) || '',
        configManifestFingerprint: configManifestFingerprint,
        dirty: false,
        pendingError: null
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
 * Metadaten der aktuellen Backup-Datei (ohne Inhalt).
 * @param {{ folder?: string }} [opts]
 * @returns {Promise<{ exists: boolean, id?: string, name?: string, lastModifiedDateTime?: string, eTag?: string, size?: number, webUrl?: string }>}
 */
export async function getCurrentBackupRemoteInfo(opts) {
    const options = opts || {};
    try {
        const cur = await findCurrentBackupItem(options);
        return {
            exists: true,
            id: cur.id,
            name: cur.name,
            lastModifiedDateTime: cur.lastModifiedDateTime || '',
            eTag: cur.eTag || cur.cTag || '',
            size: cur.size,
            webUrl: cur.webUrl || ''
        };
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e);
        if (/nicht gefunden|itemNotFound|404|not found/i.test(msg)) {
            return { exists: false };
        }
        throw e;
    }
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
    const meta = {
        at: new Date().toISOString(),
        direction: 'pull',
        fileName: CURRENT_FILE,
        folder: String(options.folder || DEFAULT_FOLDER).trim() || DEFAULT_FOLDER,
        webUrl: (cur && cur.webUrl) || it.webUrl || '',
        siteUrl: it.siteUrl || '',
        driveId: it.driveId,
        summary: describeRemoteBackup(obj),
        remoteETag: (cur && (cur.eTag || cur.cTag)) || '',
        remoteLastModified: (cur && cur.lastModifiedDateTime) || '',
        remoteExportedAt: (obj && obj.exportedAt) || '',
        contentFingerprint: (obj && obj.contentFingerprint) || '',
        dirty: false,
        pendingError: null
    };
    if (options.apply !== false || options.recordMeta) {
        saveLocalSyncMeta(meta);
    }
    return { payload: obj, item: cur, meta: meta };
}

export default {
    SCOPES_GRAPH,
    loadItMeta,
    saveItMeta,
    loadLocalSyncMeta,
    saveLocalSyncMeta,
    setLocalDirty,
    isLocalDirty,
    isReady,
    setupPageHref,
    requireItLibrary,
    putJsonOnDrive,
    listDriveFolder,
    downloadDriveItem,
    uploadCurrentBackup,
    getCurrentBackupRemoteInfo,
    findCurrentBackupItem,
    downloadCurrentBackup,
    getJsonAtDrivePath,
    uploadConfigBundle,
    downloadConfigBundle,
    applyConfigBundleFromSharePoint
};
