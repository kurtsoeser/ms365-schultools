/**
 * Lokale Overrides – nicht ins Git committen.
 *
 * Kopieren:  ms365-config.local.example.js  →  ms365-config.local.js
 *
 * Kursteams-Backend: kein Function Key. Die Seite sendet das Anmelde-Token.
 * Lizenz-API: baseUrl und optional functionKey nur hier, nie ins Repo.
 *
 * ms365-config.js lädt diese Datei automatisch, wenn sie daneben liegt.
 */
window.MS365_CONFIG_LOCAL = {
    MS365_LICENSE_API: {
        baseUrl: '',
        functionKey: ''
    }
};
