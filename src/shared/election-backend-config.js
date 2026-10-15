/**
 * Schulsprecherwahl-API (Azure Function + Azure Table Storage).
 * Stimmen liegen NICHT im Schul-SharePoint (Anonymität), sondern in Tabellen
 * hinter der Function (Partition = Schul-Tenant).
 *
 * baseUrl ohne trailing slash, endet typisch auf /api/election
 * functionKey nur in ms365-config.local.js
 */
(function (global) {
    'use strict';

    global.MS365_ELECTION_API = global.MS365_ELECTION_API || {
        baseUrl: 'https://func-ms365-election-dev.azurewebsites.net/api/election',
        /**
         * Delegierter Scope der Backend-App.
         * Election.Access bevorzugt; License.Access als Fallback (gleiche App).
         */
        scope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/Election.Access',
        /** optionaler Fallback-Scope */
        fallbackScope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/License.Access',
        functionKey: ''
    };
})(typeof window !== 'undefined' ? window : globalThis);
