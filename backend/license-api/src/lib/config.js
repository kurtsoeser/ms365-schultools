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

/**
 * Nur die License-Backend-App als Audience – kein Graph-Fallback.
 * @param {string} clientId
 */
function tokenAudiences(clientId) {
    const fromEnv = parseCsv(env('LICENSE_TOKEN_AUDIENCES', ''));
    const filtered = fromEnv.filter((a) => {
        const s = String(a).toLowerCase();
        return (
            s !== '00000003-0000-0000-c000-000000000000' &&
            s !== 'https://graph.microsoft.com' &&
            s !== 'https://graph.microsoft.com/'
        );
    });
    if (filtered.length) return filtered;
    const id = String(clientId || '').trim();
    if (!id) return [];
    return [id, 'api://' + id];
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

    const audiences = tokenAudiences(clientId);
    if (!audiences.length) {
        throw new Error(
            'LICENSE_TOKEN_AUDIENCES ist leer oder enthält nur Graph-Audiences. ' +
                'Bitte die License-App-ID bzw. api://… setzen.'
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
        spaClientId: env('LICENSE_SPA_CLIENT_ID', ''),
        operatorUpns: parseCsv(env('LICENSE_OPERATOR_UPNS', '')).map((s) => s.toLowerCase()),
        operatorOids: parseCsv(env('LICENSE_OPERATOR_OIDS', '')).map((s) => s.toLowerCase())
    };
}

module.exports = { getConfig, env, parseCsv, tokenAudiences };
