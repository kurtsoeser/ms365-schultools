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
 * @param {string} clientId
 */
function tokenAudiences(clientId) {
    const fromEnv = parseCsv(env('ELECTION_TOKEN_AUDIENCES', ''));
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
    const clientId = env('AZURE_CLIENT_ID', env('ELECTION_API_CLIENT_ID', ''));
    const audiences = tokenAudiences(clientId);
    if (!audiences.length) {
        throw new Error(
            'ELECTION_TOKEN_AUDIENCES / AZURE_CLIENT_ID fehlt. Bitte die API-App-ID setzen.'
        );
    }

    const storage =
        env('ELECTION_STORAGE_CONNECTION', '') || env('AzureWebJobsStorage', '');

    return {
        clientId,
        tokenAudiences: audiences,
        requiredScopes: parseCsv(
            env('ELECTION_REQUIRED_SCOPES', 'Election.Access,License.Access')
        ).map((s) => s.toLowerCase()),
        voterPepper: env('ELECTION_VOTER_PEPPER', 'ms365-election-dev-pepper'),
        storageConnection: storage
    };
}

module.exports = { getConfig, env, parseCsv, tokenAudiences };
