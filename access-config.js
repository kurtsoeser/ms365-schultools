/**
 * Zugangskonfiguration.
 * User-Tools: Tenant-Lizenz (License-API) nach MS365-Login – PIN optional.
 * Admin: nur Betreiber-UPNs (MS365), kein Master-PIN nötig.
 */
window.MS365_ACCESS_CONFIG = {
    /** false = keine PIN-Abfrage (Freischaltung über Tenant-Lizenz) */
    enabled: false,
    pins: ['MS365-Schule', 'IT-Team', '#kurtrocks', '#KurtRocks!', 'HLAEbensee', 'HAK-Steyr'],
    /** Legacy: nur wenn PIN-Sperre wieder aktiviert wird */
    adminPin: '#kurtrocksMS365',
    /** Betreiber: Admin ohne PIN, nur mit diesem MS365-Konto */
    operatorUpns: ['kurt@kurtsoeser.at']
};
