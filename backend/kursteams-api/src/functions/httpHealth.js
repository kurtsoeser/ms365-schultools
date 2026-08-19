'use strict';

const { app } = require('@azure/functions');
const { getAppOnlyToken } = require('../lib/msal-app-only');
const { CORS_HEADERS } = require('../lib/http-utils');

app.http('httpHealthOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'kursteams/health',
    handler: async () => ({
        status: 204,
        headers: {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type',
            'Access-Control-Max-Age': '86400'
        }
    })
});

app.http('httpHealth', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'kursteams/health',
    handler: async (_request, context) => {
        try {
            await getAppOnlyToken();
            return {
                status: 200,
                headers: CORS_HEADERS,
                jsonBody: { ok: true, graphToken: 'acquired' }
            };
        } catch (e) {
            context.error('Health check fehlgeschlagen:', e);
            return {
                status: 503,
                headers: CORS_HEADERS,
                jsonBody: { ok: false, error: e.message || String(e) }
            };
        }
    }
});
