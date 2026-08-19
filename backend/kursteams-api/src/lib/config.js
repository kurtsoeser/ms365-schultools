'use strict';

function env(name, fallback) {
    const v = process.env[name];
    if (v === undefined || v === null || String(v).trim() === '') {
        return fallback;
    }
    return String(v).trim();
}

function parseTenantAllowlist() {
    const raw = env('KURSTEAMS_ALLOWED_TENANT_IDS', env('AZURE_TENANT_ID', ''));
    return raw
        .split(/[,;\s]+/)
        .map((s) => s.trim().toLowerCase())
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
    return {
        tenantId,
        clientId,
        clientSecret,
        allowedTenantIds: parseTenantAllowlist(),
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

module.exports = { getConfig, isTenantAllowed, env };
