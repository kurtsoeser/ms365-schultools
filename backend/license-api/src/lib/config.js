'use strict';

function env(name, fallback) {
    const v = process.env[name];
    if (v === undefined || v === null || String(v).trim() === '') {
        return fallback;
    }
    return String(v).trim();
}

function parseCsv(raw) {
    return String(raw || '')
        .split(/[,;\s]+/)
        .map((s) => s.trim())
        .filter(Boolean);
}

function getConfig() {
    const tenantId = env('AZURE_TENANT_ID');
    const clientId = env('AZURE_CLIENT_ID');
    const clientSecret = env('AZURE_CLIENT_SECRET');
    if (!tenantId || !clientId || !clientSecret) {
        throw new Error(
            'Umgebungsvariablen AZURE_TENANT_ID, AZURE_CLIENT_ID und AZURE_CLIENT_SECRET sind erforderlich.'
        );
    }

    const spaClientId = env('LICENSE_SPA_CLIENT_ID', '');
    const audiences = parseCsv(env('LICENSE_TOKEN_AUDIENCES', ''));
    if (spaClientId && !audiences.includes(spaClientId)) {
        audiences.push(spaClientId);
    }
    if (!audiences.length) {
        audiences.push(
            '00000003-0000-0000-c000-000000000000',
            'https://graph.microsoft.com'
        );
    }

    return {
        tenantId,
        clientId,
        clientSecret,
        siteWebUrl: env(
            'LICENSE_SITE_WEB_URL',
            'https://kurtrocks.sharepoint.com/sites/MS365-Schultools'
        ),
        listDisplayName: env('LICENSE_LIST_NAME', 'MS365schule-Lizenzen'),
        catalogLibraryName: env('CATALOG_LIBRARY_NAME', 'MS365-Katalog'),
        catalogKursteamPath: env('CATALOG_KURSTEAM_PATH', 'vorlagen/kursteam-kanaele.json'),
        catalogMaterialsRoot: env('CATALOG_MATERIALS_ROOT', 'materialien'),
        allowedStatuses: parseCsv(env('LICENSE_ALLOWED_STATUSES', 'trial,active')).map((s) =>
            s.toLowerCase()
        ),
        tokenAudiences: audiences,
        spaClientId,
        operatorUpns: parseCsv(
            env('LICENSE_OPERATOR_UPNS', 'kurt@kurtsoeser.at')
        ).map((s) => s.toLowerCase())
    };
}

module.exports = { getConfig, env, parseCsv };
