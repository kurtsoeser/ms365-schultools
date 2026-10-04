/**
 * SharePoint-Backup lokal anwenden (nach Nutzerentscheidung).
 */
import { saveLocalSyncMeta } from './stammdaten-sharepoint-sync-api.js';

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
    if (syncMeta) {
        saveLocalSyncMeta(
            Object.assign({}, syncMeta, {
                dirty: false,
                pendingError: null
            })
        );
    }
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

export default { applySharePointBackupLocally };
