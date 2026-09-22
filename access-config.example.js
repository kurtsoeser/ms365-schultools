/**
 * Zugangskonfiguration (Vorlage → als access-config.js kopieren).
 *
 * User-Tools: Freischaltung über Tenant-Lizenz (License-API) nach MS365-Login.
 * PIN-Sperre ist optional (enabled: false = aus).
 * Admin: nur operatorUpns (MS365-Konto), kein Master-PIN nötig.
 */
window.MS365_ACCESS_CONFIG = {
    /** false = keine PIN-Abfrage */
    enabled: false,
    /**
     * Optional: Gültige PINs, falls enabled: true.
     * Vergleich ohne Groß-/Kleinschreibung; Trim am Rand.
     */
    pins: ['MS365-Schule', 'IT-Team'],
    /**
     * Legacy-Master-PIN nur relevant, wenn PIN-Sperre aktiv ist.
     */
    adminPin: 'DEIN_ADMIN_MASTER_PIN',
    /**
     * Betreiber-Konten (UPN). Sehen „Admin“ im Menü und dürfen admin.html öffnen.
     */
    operatorUpns: ['kurt@kurtsoeser.at']
};
