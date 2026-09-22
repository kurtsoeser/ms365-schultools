'use strict';

const { ConfidentialClientApplication } = require('@azure/msal-node');
const { getConfig } = require('./config');

const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

/** @type {import('@azure/msal-node').ConfidentialClientApplication | null} */
let app = null;

/** @type {{ token: string, expiresAt: number } | null} */
let cached = null;

function getMsalApp() {
    if (app) return app;
    const cfg = getConfig();
    app = new ConfidentialClientApplication({
        auth: {
            clientId: cfg.clientId,
            authority: 'https://login.microsoftonline.com/' + cfg.tenantId,
            clientSecret: cfg.clientSecret
        }
    });
    return app;
}

/**
 * App-Only-Token im Betreiber-Tenant (SharePoint-Liste lesen).
 * @returns {Promise<string>}
 */
async function getOperatorGraphToken() {
    const now = Date.now();
    if (cached && cached.expiresAt > now + 60_000) {
        return cached.token;
    }
    const msal = getMsalApp();
    const result = await msal.acquireTokenByClientCredential({
        scopes: [GRAPH_SCOPE]
    });
    if (!result || !result.accessToken) {
        throw new Error('MSAL: Kein App-Only-Token für den Betreiber-Mandanten.');
    }
    const expiresOn =
        result.expiresOn instanceof Date ? result.expiresOn.getTime() : now + 3_500_000;
    cached = { token: result.accessToken, expiresAt: expiresOn };
    return result.accessToken;
}

function clearTokenCache() {
    cached = null;
}

module.exports = { getOperatorGraphToken, clearTokenCache, GRAPH_SCOPE };
