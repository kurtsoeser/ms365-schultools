'use strict';

const { app } = require('@azure/functions');
const { getOperatorGraphToken } = require('../lib/msal-app-only');
const { CORS_HEADERS } = require('../lib/http-utils');

app.http('httpLicenseHealthOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/health',
    handler: async () => ({
        status: 204,
        headers: {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization',
            'Access-Control-Max-Age': '86400'
        }
    })
});

app.http('httpLicenseHealth', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/health',
    handler: async (_request, context) => {
        try {
            await getOperatorGraphToken();
            return {
                status: 200,
                headers: CORS_HEADERS,
                jsonBody: { ok: true, graphToken: 'acquired' }
            };
        } catch (e) {
            context.error('License health fehlgeschlagen:', e);
            return {
                status: 503,
                headers: CORS_HEADERS,
                jsonBody: { ok: false, error: e.message || String(e) }
            };
        }
    }
});
