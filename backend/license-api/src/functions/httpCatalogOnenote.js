'use strict';

const { app } = require('@azure/functions');
const {
    jsonResponse,
    corsPreflightResponse,
    bearerTokenFromRequest,
    errorJsonResponse,
    readJsonBody
} = require('../lib/http-utils');
const { requireCatalogReader, requireOperatorCaller } = require('../lib/require-operator');
const {
    listCatalogNotebooks,
    loadCatalogNotebookTree,
    listCatalogSectionPages,
    getCatalogPagePreview,
    getCatalogPageContent,
    getCatalogSectionExport
} = require('../lib/onenote-catalog');
const { publishSnapshot } = require('../lib/onenote-snapshot');

app.http('httpCatalogOnenoteNotebooksOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/notebooks',
    handler: async () => corsPreflightResponse()
});

app.http('httpCatalogOnenoteTreeOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/notebooks/{notebookId}/tree',
    handler: async () => corsPreflightResponse()
});

app.http('httpCatalogOnenoteSectionPagesOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/sections/{sectionId}/pages',
    handler: async () => corsPreflightResponse()
});

app.http('httpCatalogOnenotePagePreviewOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/pages/{pageId}/preview',
    handler: async () => corsPreflightResponse()
});

app.http('httpCatalogOnenotePageContentOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/pages/{pageId}/content',
    handler: async () => corsPreflightResponse()
});

app.http('httpCatalogOnenoteSectionExportOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/sections/{sectionId}/export',
    handler: async () => corsPreflightResponse()
});

app.http('httpCatalogOnenoteNotebooks', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/notebooks',
    handler: async (request, context) => {
        try {
            await requireCatalogReader(bearerTokenFromRequest(request));
            const result = await listCatalogNotebooks();
            return jsonResponse(200, result);
        } catch (e) {
            return errorJsonResponse(context, 'catalog/onenote/notebooks GET:', e, {
                notebooks: [],
                missing: true
            });
        }
    }
});

app.http('httpCatalogOnenoteTree', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/notebooks/{notebookId}/tree',
    handler: async (request, context) => {
        try {
            await requireCatalogReader(bearerTokenFromRequest(request));
            const notebookId = request.params.notebookId || '';
            const result = await loadCatalogNotebookTree(notebookId);
            return jsonResponse(200, result);
        } catch (e) {
            return errorJsonResponse(context, 'catalog/onenote/tree GET:', e, {
                sections: [],
                groups: []
            });
        }
    }
});

app.http('httpCatalogOnenoteSectionPages', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/sections/{sectionId}/pages',
    handler: async (request, context) => {
        try {
            await requireCatalogReader(bearerTokenFromRequest(request));
            const sectionId = request.params.sectionId || '';
            const result = await listCatalogSectionPages(sectionId);
            return jsonResponse(200, result);
        } catch (e) {
            return errorJsonResponse(context, 'catalog/onenote/pages GET:', e, { pages: [] });
        }
    }
});

app.http('httpCatalogOnenotePagePreview', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/pages/{pageId}/preview',
    handler: async (request, context) => {
        try {
            await requireCatalogReader(bearerTokenFromRequest(request));
            const pageId = request.params.pageId || '';
            const result = await getCatalogPagePreview(pageId);
            return jsonResponse(200, result);
        } catch (e) {
            return errorJsonResponse(context, 'catalog/onenote/preview GET:', e, {
                previewText: ''
            });
        }
    }
});

app.http('httpCatalogOnenotePageContent', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/pages/{pageId}/content',
    handler: async (request, context) => {
        try {
            await requireCatalogReader(bearerTokenFromRequest(request));
            const pageId = request.params.pageId || '';
            const result = await getCatalogPageContent(pageId);
            return jsonResponse(200, result);
        } catch (e) {
            return errorJsonResponse(context, 'catalog/onenote/content GET:', e);
        }
    }
});

app.http('httpCatalogOnenoteSectionExport', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/sections/{sectionId}/export',
    handler: async (request, context) => {
        try {
            await requireCatalogReader(bearerTokenFromRequest(request));
            const sectionId = request.params.sectionId || '';
            const result = await getCatalogSectionExport(sectionId);
            return jsonResponse(200, result);
        } catch (e) {
            return errorJsonResponse(context, 'catalog/onenote/export GET:', e, {
                pages: []
            });
        }
    }
});

app.http('httpCatalogOnenoteSnapshotOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/snapshot',
    handler: async () => corsPreflightResponse()
});

app.http('httpCatalogOnenoteSnapshotPut', {
    methods: ['PUT'],
    authLevel: 'anonymous',
    route: 'license/catalog/onenote/snapshot',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const body = await readJsonBody(request);
            const result = await publishSnapshot(body);
            return jsonResponse(200, result);
        } catch (e) {
            return errorJsonResponse(context, 'catalog/onenote/snapshot PUT:', e);
        }
    }
});
