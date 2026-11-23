/**
 * Zugangskonfiguration (Vorlage → als access-config.js kopieren).
 *
 * User-Tools: Freischaltung über Tenant-Lizenz (License-API) nach MS365-Login.
 * PIN-Sperre optional (enabled: false = aus). Keine echten PINs committen.
 * Betreiber-Admin: nur Azure LICENSE_OPERATOR_UPNS / LICENSE_OPERATOR_OIDS –
 * der frühere Admin-Tab „Zugänge“ ist entfernt.
 */
window.MS365_ACCESS_CONFIG = {
    /** false = keine PIN-Abfrage */
    enabled: false,
    pins: [],
    adminPins: []
};
