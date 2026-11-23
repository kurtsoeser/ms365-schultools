/**
 * Lokale Overrides – nicht ins Git committen.
 *
 * Kopieren:  ms365-config.local.example.js  →  ms365-config.local.js
 *
 * Hier dürfen -dev-URLs stehen. Produktion setzt die URLs über Deploy-Secrets
 * (ohne Function Key, a9).
 */
window.MS365_CONFIG_LOCAL = {
    MS365_LICENSE_API: {
        baseUrl: 'https://func-ms365-license-dev.azurewebsites.net/api/license',
        functionKey: '',
        /** Schul-Hosting ohne Betreiber-Portal: true unterdrückt /admin/me-Aufrufe */
        skipOperatorCheck: false
    },
    MS365_KURSTEAMS_API: {
        baseUrl:
            'https://func-ms365-kursteams-dev-cmatbeawgqf8daaq.westeurope-01.azurewebsites.net/api/kursteams'
    },
    MS365_ELECTION_API: {
        baseUrl: 'https://func-ms365-election-dev.azurewebsites.net/api/election'
    }
};
