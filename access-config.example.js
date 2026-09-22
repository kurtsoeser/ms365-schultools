/**
 * Zugangskonfiguration (Vorlage → als access-config.js kopieren).
 *
 * User-Tools: Freischaltung über Tenant-Lizenz (License-API) nach MS365-Login.
 * PIN-Sperre ist optional (enabled: false = aus). Bei enabled: true nur lokal
 * per Admin-Override (localStorage) setzen – keine echten PINs committen.
 * Admin: Betreiber-Liste nur in Azure (LICENSE_OPERATOR_UPNS / LICENSE_OPERATOR_OIDS).
 */
window.MS365_ACCESS_CONFIG = {
    /** false = keine PIN-Abfrage */
    enabled: false,
    pins: [],
    adminPins: []
};
