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

function parseTenantAllowlist() {
    const raw = env('KURSTEAMS_ALLOWED_TENANT_IDS', env('AZURE_TENANT_ID', ''));
    return parseCsv(raw).map((s) => s.toLowerCase());
}

/** Globaler Administrator, Teams-Administrator. */
const DEFAULT_OPERATOR_ROLE_TEMPLATE_IDS = [
    '62e90394-69f5-4237-9190-012177145e10',
    '69091246-20e8-4a56-aa4d-066075b2a7a8'
];

function tokenAudiences(clientId) {
    const fromEnv = parseCsv(env('KURSTEAMS_TOKEN_AUDIENCES', ''));
    if (fromEnv.length) return fromEnv;
    const id = String(clientId || '').trim();
    if (!id) return [];
    return [id, 'api://' + id];
}

function operatorRoleTemplateIds() {
    const fromEnv = parseCsv(env('KURSTEAMS_OPERATOR_ROLE_TEMPLATE_IDS', ''));
    const list = fromEnv.length ? fromEnv : DEFAULT_OPERATOR_ROLE_TEMPLATE_IDS;
    return list.map((s) => s.toLowerCase());
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
    return {
        tenantId,
        clientId,
        clientSecret,
        allowedTenantIds: parseTenantAllowlist(),
        tokenAudiences: tokenAudiences(clientId),
        operatorRoleTemplateIds: operatorRoleTemplateIds(),
        licenseApiBaseUrl: env(
            'LICENSE_API_BASE_URL',
            'https://func-ms365-license-dev.azurewebsites.net/api/license'
        ),
        storageConnectionString: env('AzureWebJobsStorage'),
        queueName: env('KURSTEAMS_JOB_QUEUE', 'kursteam-jobs'),
        blobContainer: env('KURSTEAMS_JOB_CONTAINER', 'kursteam-jobs')
    };
}

function isTenantAllowed(tenantId) {
    const cfg = getConfig();
    const tid = String(tenantId || '').trim().toLowerCase();
    if (!tid) return false;
    return cfg.allowedTenantIds.includes(tid);
}

module.exports = {
    getConfig,
    isTenantAllowed,
    env,
    tokenAudiences,
    operatorRoleTemplateIds,
    DEFAULT_OPERATOR_ROLE_TEMPLATE_IDS
};
