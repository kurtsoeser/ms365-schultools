/**
 * SharePoint-Backup lokal anwenden (nach Nutzerentscheidung).
 */
import { saveLocalSyncMeta } from './stammdaten-sharepoint-sync-api.js';
import { applySelectiveLocalStoragePatch } from './stammdaten-sharepoint-selective-import.js';

function saveSyncMetaClean(syncMeta) {
    if (!syncMeta) return;
    saveLocalSyncMeta(
        Object.assign({}, syncMeta, {
            dirty: false,
            pendingError: null
        })
    );
}

/**
 * @param {object} payload Backup-JSON
 * @param {object} syncMeta Sync-Metadaten für localStorage
 * @param {{ reload?: boolean }} [opts]
 */
export function applySharePointBackupLocally(payload, syncMeta, opts) {
    const options = opts || {};
    const bb = window.ms365BrowserBackup;
    if (!bb || typeof bb.importPayload !== 'function') {
        throw new Error('Backup-Modul fehlt.');
    }
    bb.importPayload(payload);
    saveSyncMetaClean(syncMeta);
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-tenant-settings-changed', {
                detail: { source: 'spo-backup-import' }
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
}

/**
 * Nur ausgewählte Bereiche aus dem SharePoint-Backup übernehmen.
 * @param {object} payload Backup-JSON
 * @param {string[]} selectedIds Bereichs-IDs aus buildSelectiveImportPlan
 * @param {object|null} syncMeta
 * @param {{ reload?: boolean, localPayload?: object }} [opts]
 */
export function applySharePointSelectiveImportLocally(payload, selectedIds, syncMeta, opts) {
    const options = opts || {};
    const applied = applySelectiveLocalStoragePatch(payload, selectedIds, {
        localPayload: options.localPayload
    });
    saveSyncMetaClean(syncMeta);
    if (options.reload === true) {
        window.setTimeout(function () {
            window.location.reload();
        }, 80);
    }
    return applied;
}

export default { applySharePointBackupLocally, applySharePointSelectiveImportLocally };
