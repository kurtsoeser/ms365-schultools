'use strict';

const { app } = require('@azure/functions');
const {
    jsonResponse,
    corsPreflightResponse,
    bearerTokenFromRequest,
    readJsonBody
} = require('../lib/http-utils');
const { requireOperatorCaller } = require('../lib/require-operator');
const {
    listLicenses,
    createLicense,
    updateLicense,
    deleteLicense
} = require('../lib/sharepoint-license');

app.http('httpLicenseAdminSchoolsOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/admin/schools',
    handler: async () => corsPreflightResponse()
});

app.http('httpLicenseAdminSchoolsOptionsId', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/admin/schools/{id}',
    handler: async () => corsPreflightResponse()
});

app.http('httpLicenseAdminSchoolsList', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/admin/schools',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const schools = await listLicenses();
            return jsonResponse(200, { schools });
        } catch (e) {
            const status = e.status && Number.isFinite(e.status) ? e.status : 500;
            if (status >= 500) context.error('license/admin/schools GET:', e);
            return jsonResponse(status >= 400 && status < 600 ? status : 500, {
                error: e.message || String(e),
                schools: []
            });
        }
    }
});

app.http('httpLicenseAdminSchoolsCreate', {
    methods: ['POST'],
    authLevel: 'anonymous',
    route: 'license/admin/schools',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const body = await readJsonBody(request);
            const school = await createLicense(body);
            return jsonResponse(201, { school });
        } catch (e) {
            const status = e.status && Number.isFinite(e.status) ? e.status : 500;
            if (status >= 500) context.error('license/admin/schools POST:', e);
            return jsonResponse(status >= 400 && status < 600 ? status : 500, {
                error: e.message || String(e)
            });
        }
    }
});

app.http('httpLicenseAdminSchoolsUpdate', {
    methods: ['PATCH'],
    authLevel: 'anonymous',
    route: 'license/admin/schools/{id}',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const id = request.params && request.params.id ? String(request.params.id) : '';
            const body = await readJsonBody(request);
            const school = await updateLicense(id, body);
            return jsonResponse(200, { school });
        } catch (e) {
            const status = e.status && Number.isFinite(e.status) ? e.status : 500;
            if (status >= 500) context.error('license/admin/schools PATCH:', e);
            return jsonResponse(status >= 400 && status < 600 ? status : 500, {
                error: e.message || String(e)
            });
        }
    }
});

app.http('httpLicenseAdminSchoolsDelete', {
    methods: ['DELETE'],
    authLevel: 'anonymous',
    route: 'license/admin/schools/{id}',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const id = request.params && request.params.id ? String(request.params.id) : '';
            const result = await deleteLicense(id);
            return jsonResponse(200, result);
        } catch (e) {
            const status = e.status && Number.isFinite(e.status) ? e.status : 500;
            if (status >= 500) context.error('license/admin/schools DELETE:', e);
            return jsonResponse(status >= 400 && status < 600 ? status : 500, {
                error: e.message || String(e)
            });
        }
    }
});
