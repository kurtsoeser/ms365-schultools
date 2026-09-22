/**
 * Betreiber-Backend für Schul-Lizenzen (Phase 1+).
 * Daten liegen in SharePoint auf dem Betreiber-Tenant – nicht im Schul-Tenant.
 * Zentraler Vorlagen-Katalog: Dokumentbibliothek MS365-Katalog
 * (Datei vorlagen/kursteam-kanaele.json), gelesen über die License-API.
 */
(function (global) {
    'use strict';

    global.MS365_LICENSE_BACKEND = {
        /** SharePoint-Website (Betreiber) */
        siteWebUrl: 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools',
        /** Anzeigename der Liste */
        listDisplayName: 'MS365-Schultools-Lizenzen',
        /** Status-Werte (Choice) */
        statusChoices: ['trial', 'active', 'expired', 'blocked'],
        /** Erlaubte Status für Tool-Zugang (Phase 3) */
        allowedStatuses: ['trial', 'active'],
        /**
         * Spalten-Hinweis:
         * PrimaryDomain = eine Hauptdomain
         * AdditionalDomains = weitere Domains/URLs (eine pro Zeile)
         */
        domainColumns: ['PrimaryDomain', 'AdditionalDomains']
    };

    /**
     * License-API (Azure Function). baseUrl ohne trailing slash, endet auf /api/license.
     * functionKey nur in ms365-config.local.js (nicht committen).
     */
    global.MS365_LICENSE_API = global.MS365_LICENSE_API || {
        baseUrl: '',
        functionKey: ''
    };
})(typeof window !== 'undefined' ? window : globalThis);
