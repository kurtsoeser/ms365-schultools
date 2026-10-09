/**
 * Tenant-Panel Synchron: Brücke SharePoint ↔ Stammdaten-Ampel.
 */
export function mountTenantSpoPanel() {
    if (typeof window === 'undefined') return;

    window.addEventListener('ms365-app-local-data-changed', function (ev) {
        const src = ev && ev.detail && ev.detail.source;
        if (
            src === 'spo-auto-pull' ||
            src === 'spo-sync' ||
            src === 'tenant-settings' ||
            src === 'browser-backup-import'
        ) {
            try {
                window.dispatchEvent(new CustomEvent('ms365-spo-sync-status'));
            } catch {
                /* ignore */
            }
        }
    });
}
