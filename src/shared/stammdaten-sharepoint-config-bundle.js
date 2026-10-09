/**
 * Aufgeteilte Schul-Konfiguration in der IT-Bibliothek (neben dem Vollbackup-Monolithen).
 * Kein Graph/DOM – testbar.
 */

export const CONFIG_FOLDER = 'config';
export const CONFIG_MANIFEST_FILE = 'manifest.json';
export const CONFIG_BUNDLE_KIND = 'ms365-spo-config-bundle-v1';
export const CONFIG_BUNDLE_SCHEMA_VERSION = 1;

/** @typedef {{ id: string, fileName: string, keys: string[] }} ConfigBundlePartDef */

/** @type {ConfigBundlePartDef[]} */
export const CONFIG_BUNDLE_PARTS = [
    {
        id: 'stammdaten-core',
        fileName: 'stammdaten-core.json',
        keys: [
            'ms365-schooltool-data-v2',
            'ms365-tenant-settings-v1',
            'ms365-school-email-domain-v1',
            'ms365-class-nick-schema-v1',
            'ms365-schueler-lehrer-gruppen-v1',
            'ms365-schueler-lehrer-gruppen-v2',
            'ms365-schulstruktur-sync-v1',
            'ms365-schulstruktur-match-v1',
            'ms365-schulstruktur-tenant-cache-v1',
            'ms365-schulstruktur-sync-ui-mode-v1',
            'ms365-schulstruktur-ad-flags-v1',
            'ms365-stammdaten-it-library-v1',
            'ms365-stammdaten-it-library-by-tenant-v2',
            'ms365-stammdaten-listen-perms-v1'
        ]
    },
    {
        id: 'dashboard-access',
        fileName: 'dashboard-access.json',
        keys: [
            'ms365-dashboard-audience-groups-v1',
            'ms365-dashboard-tool-access-v1',
            'ms365-dashboard-favorites-v1',
            'ms365-dashboard-recent-tools-v1',
            'ms365-dashboard-order-catalog-v1',
            'ms365-dashboard-order-m365-v3',
            'ms365-dashboard-order-sharepoint-v2',
            'ms365-dashboard-order-v1',
            'ms365-dash-catalog-fold-v1',
            'ms365-dashboard-category-tab-v1',
            'ms365-dashboard-expert-open-v1',
            'ms365-dashboard-view-v1',
            'ms365-dashboard-it-preview-v1',
            'ms365-dashboard-setup-dismissed-v1'
        ]
    },
    {
        id: 'permissions-schularbeiten',
        fileName: 'permissions-schularbeiten.json',
        keys: ['ms365-schularbeiten-perms-v1', 'ms365-sa-settings-v1']
    },
    {
        id: 'permissions-freistellung',
        fileName: 'permissions-freistellung.json',
        keys: [
            'ms365-freistellung-perms-v1',
            'ms365-freistellung-setup-v1',
            'ms365-freistellung-setup-step-v1',
            'ms365-freistellung-setup-step-v2',
            'ms365-freistellung-setup-step-v3',
            'ms365-freistellung-kategorien-extra-v1',
            'ms365-freistellung-planer-role-v1',
            'ms365-freistellung-planer-site-v1',
            'ms365-freistellung-planer-demo-klasse-v1',
            'ms365-freistellung-student-klasse-pick-v1'
        ]
    },
    {
        id: 'permissions-lehrer-freistellung',
        fileName: 'permissions-lehrer-freistellung.json',
        keys: [
            'ms365-lfr-perms-v1',
            'ms365-lfr-setup-v1',
            'ms365-lfr-setup-step-v1',
            'ms365-lfr-planer-role-v1',
            'ms365-lfr-planer-site-v1',
            'ms365-lfr-demo-items-v1',
            'ms365-lfr-outlook-event-v1',
            'ms365-pa-done-lehrer-freistellung'
        ]
    }
];

/**
 * @param {Record<string, string>} localSlice
 */
export function fingerprintLocalStorageSlice(localSlice) {
    const local = localSlice && typeof localSlice === 'object' ? localSlice : {};
    const parts = [];
    Object.keys(local)
        .sort()
        .forEach(function (k) {
            parts.push(k + '\0' + String(local[k] == null ? '' : local[k]));
        });
    const s = parts.join('\n');
    let h = 5381;
    for (let i = 0; i < s.length; i++) {
        h = (Math.imul(h, 33) ^ s.charCodeAt(i)) | 0;
    }
    return (h >>> 0).toString(16) + ':' + parts.length;
}

/**
 * @param {string} folderName
 * @param {string} fileName
 */
export function configPartRelativePath(folderName, fileName) {
    const folder = String(folderName || CONFIG_FOLDER)
        .trim()
        .replace(/^\/+|\/+$/g, '')
        .replace(/\\/g, '/');
    const name = String(fileName || '').trim();
    if (!folder) return name;
    if (!name) return folder;
    return folder + '/' + name;
}

/**
 * @param {object} backupPayload Vollbackup (buildBackup)
 * @param {{ configFolder?: string, monolithPath?: string, tenantId?: string }} [opts]
 */
export function buildConfigBundleFromBackup(backupPayload, opts) {
    const options = opts || {};
    const payload = backupPayload && typeof backupPayload === 'object' ? backupPayload : {};
    const store = payload.localStorage && typeof payload.localStorage === 'object' ? payload.localStorage : {};
    const configFolder = String(options.configFolder || CONFIG_FOLDER).trim() || CONFIG_FOLDER;
    const monolithPath = String(options.monolithPath || '').trim();

    /** @type {Array<{ id: string, path: string, body: object, keyCount: number, fingerprint: string }>} */
    const files = [];

    CONFIG_BUNDLE_PARTS.forEach(function (part) {
        /** @type {Record<string, string>} */
        const slice = {};
        (part.keys || []).forEach(function (key) {
            if (!Object.prototype.hasOwnProperty.call(store, key)) return;
            slice[key] = store[key];
        });
        const body = {
            schemaVersion: CONFIG_BUNDLE_SCHEMA_VERSION,
            bundleKind: CONFIG_BUNDLE_KIND,
            bundlePart: part.id,
            exportedAt: payload.exportedAt || new Date().toISOString(),
            schoolName: payload.schoolName || null,
            domain: payload.domain || null,
            keyCount: Object.keys(slice).length,
            fingerprint: fingerprintLocalStorageSlice(slice),
            localStorage: slice
        };
        files.push({
            id: part.id,
            path: configPartRelativePath(configFolder, part.fileName),
            body: body,
            keyCount: Object.keys(slice).length,
            fingerprint: body.fingerprint
        });
    });

    const manifest = {
        schemaVersion: CONFIG_BUNDLE_SCHEMA_VERSION,
        bundleKind: CONFIG_BUNDLE_KIND,
        exportedAt: payload.exportedAt || new Date().toISOString(),
        tenantId: String(options.tenantId || '').trim() || null,
        schoolName: payload.schoolName || null,
        domain: payload.domain || null,
        contentFingerprint: String(payload.contentFingerprint || '').trim() || null,
        monolithPath: monolithPath || null,
        files: files.map(function (f) {
            return {
                id: f.id,
                path: f.path,
                keyCount: f.keyCount,
                fingerprint: f.fingerprint
            };
        })
    };

    return {
        manifest: manifest,
        manifestPath: configPartRelativePath(configFolder, CONFIG_MANIFEST_FILE),
        files: files
    };
}

/**
 * @param {unknown} manifest
 */
export function isConfigBundleManifest(manifest) {
    const m = manifest && typeof manifest === 'object' ? manifest : null;
    if (!m) return false;
    if (m.bundleKind !== CONFIG_BUNDLE_KIND) return false;
    if (!Array.isArray(m.files) || !m.files.length) return false;
    return true;
}

/**
 * @param {unknown} partBody
 * @param {string} expectedPartId
 */
export function parseConfigBundlePart(partBody, expectedPartId) {
    const o = partBody && typeof partBody === 'object' ? partBody : null;
    if (!o || o.bundleKind !== CONFIG_BUNDLE_KIND) return null;
    if (expectedPartId && String(o.bundlePart || '') !== String(expectedPartId)) return null;
    const ls = o.localStorage && typeof o.localStorage === 'object' ? o.localStorage : {};
    return {
        bundlePart: String(o.bundlePart || expectedPartId || ''),
        exportedAt: String(o.exportedAt || ''),
        fingerprint: String(o.fingerprint || ''),
        localStorage: ls
    };
}

/**
 * @param {Array<{ localStorage: Record<string, string> }>} parts
 */
export function mergeConfigPartsLocalStorage(parts) {
    /** @type {Record<string, string>} */
    const out = {};
    (parts || []).forEach(function (p) {
        const ls = p && p.localStorage && typeof p.localStorage === 'object' ? p.localStorage : {};
        Object.keys(ls).forEach(function (k) {
            out[k] = ls[k];
        });
    });
    return out;
}
