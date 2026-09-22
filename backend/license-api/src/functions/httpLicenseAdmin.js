'use strict';

const { app } = require('@azure/functions');
const {
    jsonResponse,
    corsPreflightResponse,
    bearerTokenFromRequest,
    readJsonBody,
    errorJsonResponse
} = require('../lib/http-utils');
const { requireOperatorCaller } = require('../lib/require-operator');
const {
    listLicenses,
    createLicense,
    updateLicense,
    deleteLicense,
    listExtraColumns,
    createExtraColumn,
    deleteExtraColumn
} = require('../lib/sharepoint-license');

app.http('httpLicenseAdminMeOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/admin/me',
    handler: async () => corsPreflightResponse()
});

app.http('httpLicenseAdminMe', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/admin/me',
    handler: async (request, context) => {
        try {
            const caller = await requireOperatorCaller(bearerTokenFromRequest(request));
            return jsonResponse(200, {
                operator: true,
                user: {
                    oid: caller.oid || null,
                    upn: caller.upn || null,
                    name: caller.name || null,
                    tid: caller.tid || null
                }
            });
        } catch (e) {
            return errorJsonResponse(context, 'license/admin/me GET:', e, {
                operator: false,
                user: null
            });
        }
    }
});

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

app.http('httpLicenseAdminColumnsOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/admin/columns',
    handler: async () => corsPreflightResponse()
});

app.http('httpLicenseAdminColumnsOptionsName', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/admin/columns/{name}',
    handler: async () => corsPreflightResponse()
});

app.http('httpLicenseAdminSchoolsList', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/admin/schools',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const result = await listLicenses();
            return jsonResponse(200, {
                schools: result.schools || [],
                columns: result.columns || []
            });
        } catch (e) {
            return errorJsonResponse(context, 'license/admin/schools GET:', e, {
                schools: [],
                columns: []
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
            return errorJsonResponse(context, 'license/admin/schools POST:', e);
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
            return errorJsonResponse(context, 'license/admin/schools PATCH:', e);
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
            return errorJsonResponse(context, 'license/admin/schools DELETE:', e);
        }
    }
});

app.http('httpLicenseAdminColumnsList', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/admin/columns',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const columns = await listExtraColumns({ force: true });
            return jsonResponse(200, { columns });
        } catch (e) {
            return errorJsonResponse(context, 'license/admin/columns GET:', e, { columns: [] });
        }
    }
});

app.http('httpLicenseAdminColumnsCreate', {
    methods: ['POST'],
    authLevel: 'anonymous',
    route: 'license/admin/columns',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const body = await readJsonBody(request);
            const column = await createExtraColumn(body);
            return jsonResponse(201, { column });
        } catch (e) {
            return errorJsonResponse(context, 'license/admin/columns POST:', e);
        }
    }
});

app.http('httpLicenseAdminColumnsDelete', {
    methods: ['DELETE'],
    authLevel: 'anonymous',
    route: 'license/admin/columns/{name}',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const name = request.params && request.params.name ? String(request.params.name) : '';
            const result = await deleteExtraColumn(name);
            return jsonResponse(200, result);
        } catch (e) {
            return errorJsonResponse(context, 'license/admin/columns DELETE:', e);
        }
    }
});
