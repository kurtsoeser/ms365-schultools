/**
 * Zugangskonfiguration (öffentlich ausgeliefert).
 * User-Tools: Tenant-Lizenz nach MS365-Login (License-API) – kein PIN im Alltag.
 * PIN-Sperre aus (enabled: false); Legacy-Override nur noch in pin-gate, nicht im Admin-UI.
 * Betreiber-Admin: nur serverseitig (LICENSE_OPERATOR_UPNS / OIDS) – Tab „Zugänge“ entfernt.
 */
window.MS365_ACCESS_CONFIG = {
    /** false = keine PIN-Abfrage */
    enabled: false,
    /** Leer lassen – PINs gehören nicht ins Repo / auf den Host */
    pins: [],
    /** Legacy: ungenutzt, solange enabled: false */
    adminPins: []
};
