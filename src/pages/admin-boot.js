/*
 * Seiten-Skript für admin.html (klassisches Skript, kein Modul).
 * Theme und Marke setzt src/shared/theme-boot.js davor.
 *
 * Betreiber: License-API Operator (UPN/OID). Schulen: Tenant-Lizenz nach Login.
 * Kein PIN-Tab mehr – Freischaltung nicht über access-config-PINs.
 */
(function () {
    try {
        document.documentElement.removeAttribute('data-ms365-admin-boot');
    } catch (e) {}
})();
