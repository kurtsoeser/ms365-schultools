'use strict';

const { ConfidentialClientApplication } = require('@azure/msal-node');
const { getConfig } = require('./config');

const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

/** @type {Map<string, import('@azure/msal-node').ConfidentialClientApplication>} */
const apps = new Map();

/** @type {Map<string, { token: string, expiresAt: number }>} */
const tokenCache = new Map();

function getMsalApp(tenantId) {
    const cfg = getConfig();
    const tid = String(tenantId || cfg.tenantId).trim();
    if (!tid) {
        throw new Error('tenantId fehlt für App-Only-Token.');
    }
    if (apps.has(tid)) return apps.get(tid);

    const app = new ConfidentialClientApplication({
        auth: {
            clientId: cfg.clientId,
            authority: 'https://login.microsoftonline.com/' + tid,
            clientSecret: cfg.clientSecret
        }
    });
    apps.set(tid, app);
    return app;
}

/**
 * @param {string} [tenantId] Zielmandant des Jobs (Multimandanten-App mit Admin Consent)
 * @returns {Promise<string>}
 */
async function getAppOnlyToken(tenantId) {
    const cfg = getConfig();
    const tid = String(tenantId || cfg.tenantId).trim();
    const now = Date.now();
    const cached = tokenCache.get(tid);
    if (cached && cached.expiresAt > now + 60_000) {
        return cached.token;
    }

    const msal = getMsalApp(tid);
    const result = await msal.acquireTokenByClientCredential({
        scopes: [GRAPH_SCOPE]
    });
    if (!result || !result.accessToken) {
        throw new Error(
            'MSAL: Kein Access Token für Mandant ' +
                tid +
                ' (Admin Consent für die Backend-App in diesem Mandanten?)'
        );
    }
    const expiresOn =
        result.expiresOn instanceof Date ? result.expiresOn.getTime() : now + 3_500_000;
    tokenCache.set(tid, { token: result.accessToken, expiresAt: expiresOn });
    return result.accessToken;
}

function clearTokenCache() {
    tokenCache.clear();
}

module.exports = { getAppOnlyToken, clearTokenCache, GRAPH_SCOPE };
