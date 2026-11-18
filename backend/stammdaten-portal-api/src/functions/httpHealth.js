'use strict';

const { app } = require('@azure/functions');
const { jsonResponse, corsPreflightResponse } = require('../lib/http-utils');

app.http('health', {
    methods: ['GET', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'health',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        return jsonResponse(request, 200, { ok: true, service: 'stammdaten-portal-api' });
    }
});
