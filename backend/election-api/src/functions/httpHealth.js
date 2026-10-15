'use strict';

const { app } = require('@azure/functions');
const { CORS_HEADERS, jsonResponse } = require('../lib/http-utils');
const { ensureTables } = require('../lib/tables');
const { getConfig } = require('../lib/config');

app.http('httpElectionHealthOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'election/health',
    handler: async () => ({
        status: 204,
        headers: {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, x-functions-key',
            'Access-Control-Max-Age': '86400'
        }
    })
});

app.http('httpElectionHealth', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'election/health',
    handler: async (_request, context) => {
        try {
            getConfig();
            await ensureTables();
            return jsonResponse(200, { ok: true, storage: 'azure-tables' });
        } catch (e) {
            context.error('Election health fehlgeschlagen:', e);
            return {
                status: 503,
                headers: CORS_HEADERS,
                jsonBody: { ok: false, error: 'Storage oder Config nicht bereit.' }
            };
        }
    }
});
