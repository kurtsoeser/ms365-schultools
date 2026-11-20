'use strict';

const { app } = require('@azure/functions');
const { jsonResponse, corsPreflightResponse, errorJsonResponse } = require('../lib/http-utils');

const MATURA_URL = 'https://www.matura.gv.at/pruefungstermine';

app.http('httpMaturaTermineOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/matura-termine',
    handler: async () => corsPreflightResponse()
});

app.http('httpMaturaTermine', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/matura-termine',
    handler: async (_request, context) => {
        try {
            const res = await fetch(MATURA_URL, {
                headers: {
                    Accept: 'text/html',
                    'User-Agent': 'MS365-Schul-Tools'
                }
            });
            if (!res.ok) {
                const err = new Error('matura.gv.at antwortet mit HTTP ' + res.status);
                err.status = 502;
                throw err;
            }
            const html = await res.text();
            if (!html || html.length < 500) {
                const err = new Error('matura.gv.at hat keine Termine geliefert.');
                err.status = 502;
                throw err;
            }
            return jsonResponse(200, {
                html: html.slice(0, 200000),
                fetchedAt: new Date().toISOString(),
                source: MATURA_URL
            });
        } catch (e) {
            return errorJsonResponse(context, 'matura-termine', e);
        }
    }
});
