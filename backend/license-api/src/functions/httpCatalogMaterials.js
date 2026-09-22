'use strict';

const { app } = require('@azure/functions');
const {
    jsonResponse,
    corsPreflightResponse,
    bearerTokenFromRequest,
    readJsonBody,
    errorJsonResponse
} = require('../lib/http-utils');
const { requireCatalogReader, requireOperatorCaller } = require('../lib/require-operator');
const {
    listMaterials,
    readMaterialFile,
    createMaterialFolder,
    ensureMaterialFolder,
    writeMaterialFile,
    deleteMaterialItem,
    defaultMaterialsPathForTemplate,
    MAX_MATERIAL_BYTES
} = require('../lib/sharepoint-catalog');

app.http('httpCatalogMaterialsOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/catalog/materials',
    handler: async () => corsPreflightResponse()
});

app.http('httpCatalogMaterialsFileOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/catalog/materials/file',
    handler: async () => corsPreflightResponse()
});

app.http('httpCatalogMaterialsFolderOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/catalog/materials/folder',
    handler: async () => corsPreflightResponse()
});

app.http('httpCatalogMaterialsList', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/catalog/materials',
    handler: async (request, context) => {
        try {
            await requireCatalogReader(bearerTokenFromRequest(request));
            const path = request.query.get('path') || '';
            const result = await listMaterials(path);
            return jsonResponse(200, result);
        } catch (e) {
            return errorJsonResponse(context, 'catalog/materials GET:', e, {
                path: '',
                missing: true,
                items: []
            });
        }
    }
});

app.http('httpCatalogMaterialsDelete', {
    methods: ['DELETE'],
    authLevel: 'anonymous',
    route: 'license/catalog/materials',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const path = request.query.get('path') || '';
            if (!path) return jsonResponse(400, { error: 'Query path fehlt.' });
            const result = await deleteMaterialItem(path);
            return jsonResponse(200, result);
        } catch (e) {
            return errorJsonResponse(context, 'catalog/materials DELETE:', e);
        }
    }
});

app.http('httpCatalogMaterialsFolder', {
    methods: ['POST'],
    authLevel: 'anonymous',
    route: 'license/catalog/materials/folder',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const body = await readJsonBody(request);
            if (body.ensurePath) {
                const ensured = await ensureMaterialFolder(String(body.ensurePath || ''));
                const listed = await listMaterials(ensured.path);
                return jsonResponse(200, Object.assign(listed, { ensured: true }));
            }
            if (body.templateId) {
                const path = defaultMaterialsPathForTemplate(String(body.templateId));
                const ensured = await ensureMaterialFolder(path);
                const listed = await listMaterials(ensured.path);
                return jsonResponse(200, Object.assign(listed, { path: ensured.path, forTemplate: true }));
            }
            const parent = String(body.parentPath || body.path || '');
            const name = String(body.name || body.folderName || '').trim();
            if (!name) return jsonResponse(400, { error: 'Ordnername fehlt.' });
            const listed = await createMaterialFolder(parent, name);
            return jsonResponse(200, listed);
        } catch (e) {
            return errorJsonResponse(context, 'catalog/materials/folder POST:', e, { items: [] });
        }
    }
});

app.http('httpCatalogMaterialsFile', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/catalog/materials/file',
    handler: async (request, context) => {
        try {
            await requireCatalogReader(bearerTokenFromRequest(request));
            const path = request.query.get('path') || '';
            if (!path) {
                return jsonResponse(400, { error: 'Query path fehlt.' });
            }
            const file = await readMaterialFile(path);
            const asciiName = String(file.name || 'datei')
                .replace(/[^\x20-\x7E]/g, '_')
                .replace(/"/g, '');
            return {
                status: 200,
                headers: {
                    'Content-Type': file.contentType || 'application/octet-stream',
                    'Content-Disposition':
                        'attachment; filename="' +
                        asciiName +
                        '"; filename*=UTF-8\'\'' +
                        encodeURIComponent(file.name),
                    'Access-Control-Allow-Origin': '*',
                    'Access-Control-Expose-Headers': 'Content-Disposition, Content-Type',
                    'Cache-Control': 'no-store',
                    'X-Catalog-Path': file.path
                },
                body: file.bytes
            };
        } catch (e) {
            return errorJsonResponse(context, 'catalog/materials/file GET:', e);
        }
    }
});

app.http('httpCatalogMaterialsUpload', {
    methods: ['PUT'],
    authLevel: 'anonymous',
    route: 'license/catalog/materials/file',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const path = request.query.get('path') || '';
            if (!path) return jsonResponse(400, { error: 'Query path fehlt.' });
            const buf = Buffer.from(await request.arrayBuffer());
            if (buf.length > MAX_MATERIAL_BYTES) {
                return jsonResponse(413, { error: 'Datei ist größer als 8 MB.' });
            }
            const ct = request.headers.get('content-type') || 'application/octet-stream';
            const saved = await writeMaterialFile(path, buf, ct);
            return jsonResponse(200, saved);
        } catch (e) {
            return errorJsonResponse(context, 'catalog/materials/file PUT:', e);
        }
    }
});
