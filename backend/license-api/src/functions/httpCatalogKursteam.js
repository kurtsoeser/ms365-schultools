'use strict';

const { app } = require('@azure/functions');
const {
    jsonResponse,
    corsPreflightResponse,
    bearerTokenFromRequest,
    readJsonBody
} = require('../lib/http-utils');
const { requireOperatorCaller, requireCatalogReader } = require('../lib/require-operator');
const { readKursteamCatalog, writeKursteamCatalog } = require('../lib/sharepoint-catalog');

app.http('httpCatalogKursteamOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/catalog/kursteam-templates',
    handler: async () => corsPreflightResponse()
});

app.http('httpCatalogKursteamGet', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/catalog/kursteam-templates',
    handler: async (request, context) => {
        try {
            await requireCatalogReader(bearerTokenFromRequest(request));
            const catalog = await readKursteamCatalog();
            return jsonResponse(200, catalog);
        } catch (e) {
            const status = e.status && Number.isFinite(e.status) ? e.status : 500;
            if (status >= 500) context.error('catalog/kursteam-templates GET:', e);
            return jsonResponse(status >= 400 && status < 600 ? status : 500, {
                error: e.message || String(e),
                templates: [],
                schoolForms: [],
                missing: true
            });
        }
    }
});

app.http('httpCatalogKursteamPut', {
    methods: ['PUT'],
    authLevel: 'anonymous',
    route: 'license/catalog/kursteam-templates',
    handler: async (request, context) => {
        try {
            const caller = await requireOperatorCaller(bearerTokenFromRequest(request));
            const body = await readJsonBody(request);
            const catalog = await writeKursteamCatalog(body, caller.upn || '');
            return jsonResponse(200, catalog);
        } catch (e) {
            const status = e.status && Number.isFinite(e.status) ? e.status : 500;
            if (status >= 500) context.error('catalog/kursteam-templates PUT:', e);
            const message = e.message || String(e);
            const hint =
                status === 403 || /access denied|accessdenied|forbidden/i.test(message)
                    ? ' Die App braucht Schreibrecht auf der Website (Sites.ReadWrite.All oder Sites.Selected mit Schreiben).'
                    : '';
            return jsonResponse(status >= 400 && status < 600 ? status : 500, {
                error: message + hint
            });
        }
    }
});
