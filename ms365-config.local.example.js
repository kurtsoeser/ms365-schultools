/**
 * Lokale Geheimnisse – nicht ins Git committen.
 *
 * === Lokal (Entwicklung) ===
 *   1. Diese Datei kopieren:  ms365-config.local.example.js  →  ms365-config.local.js
 *   2. functionKey eintragen (Azure Portal → Function App → App-Schlüssel → default)
 *
 * === GitHub Pages (Live-App) ===
 *   Repository → Settings → Secrets and variables → Actions
 *   Secret anlegen:  KURSTEAMS_FUNCTION_KEY  =  Ihr App-Schlüssel
 *   Beim Deploy erzeugt der Workflow automatisch dist/ms365-config.local.js.
 *
 * ms365-config.js lädt ms365-config.local.js automatisch, wenn die Datei neben ms365-config.js liegt.
 */
window.MS365_CONFIG_LOCAL = {
    MS365_KURSTEAMS_API: {
        functionKey: ''
    }
};
