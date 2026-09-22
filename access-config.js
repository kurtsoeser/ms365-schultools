/**
 * Zugangskonfiguration (öffentlich ausgeliefert).
 * User-Tools: Tenant-Lizenz nach MS365-Login (License-API).
 * PIN-Sperre aus – keine PINs in dieser Datei.
 * Admin: Betreiber nur serverseitig (LICENSE_OPERATOR_UPNS / OIDS in Azure).
 */
window.MS365_ACCESS_CONFIG = {
    /** false = keine PIN-Abfrage */
    enabled: false,
    /** Leer lassen – PINs gehören nicht ins Repo / auf den Host */
    pins: [],
    /** Legacy: ungenutzt, solange enabled: false */
    adminPins: []
};
