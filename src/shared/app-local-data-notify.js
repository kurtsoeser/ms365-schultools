/**
 * Signalisiert lokale App-Datenänderungen für Browser-Backup-Hinweise und SharePoint-Auto-Sync.
 * @param {string} [source]
 * @param {object} [detail]
 */
export function notifyAppLocalDataChanged(source, detail) {
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365BrowserBackup &&
            typeof window.ms365BrowserBackup.notifyLocalDataChanged === 'function'
        ) {
            window.ms365BrowserBackup.notifyLocalDataChanged(source, detail);
            return;
        }
        const payload = Object.assign({}, detail || {}, {
            source: String(source || 'local-data'),
            at: new Date().toISOString()
        });
        window.dispatchEvent(new CustomEvent('ms365-local-data-changed', { detail: payload }));
    } catch {
        /* ignore */
    }
}
