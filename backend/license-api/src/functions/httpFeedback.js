'use strict';

const { app } = require('@azure/functions');
const {
    jsonResponse,
    corsPreflightResponse,
    bearerTokenFromRequest,
    readJsonBody,
    errorJsonResponse
} = require('../lib/http-utils');
const { getConfig } = require('../lib/config');
const { validateCallerToken } = require('../lib/validate-token');
const { requireOperatorCaller } = require('../lib/require-operator');
const { lookupLicenseFields } = require('../lib/sharepoint-license');
const { evaluateLicense } = require('../lib/evaluate-license');
const {
    createFeedback,
    listFeedback,
    updateFeedback,
    ensureFeedbackSchema,
    readFeedbackFile,
    CATEGORIES,
    STATUSES
} = require('../lib/sharepoint-feedback');

/**
 * Domain grob aus UPN ableiten.
 * @param {string} upn
 */
function domainFromUpn(upn) {
    const s = String(upn || '');
    const at = s.lastIndexOf('@');
    return at >= 0 ? s.slice(at + 1).trim().toLowerCase() : '';
}

app.http('httpFeedbackSubmitOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/feedback',
    handler: async () => corsPreflightResponse()
});

/**
 * Feedback einreichen – jeder angemeldete Nutzer (License.Access Token).
 * Absender-Daten kommen aus dem Token / Lizenz, nicht aus dem Client-Vertrauen.
 */
app.http('httpFeedbackSubmit', {
    methods: ['POST'],
    authLevel: 'anonymous',
    route: 'license/feedback',
    handler: async (request, context) => {
        try {
            const cfg = getConfig();
            const token = bearerTokenFromRequest(request);
            const caller = await validateCallerToken(token, cfg.tokenAudiences);
            const body = await readJsonBody(request);

            let schoolName = '';
            let domain = domainFromUpn(caller.upn);
            try {
                const fields = await lookupLicenseFields(caller.tid);
                const result = evaluateLicense({
                    tenantId: caller.tid,
                    fields,
                    allowedStatuses: cfg.allowedStatuses
                });
                schoolName = String(result.schoolName || '').trim();
                if (result.primaryDomain) domain = String(result.primaryDomain).trim();
            } catch {
                /* Lizenz optional – Feedback trotzdem annehmen */
            }

            const item = await createFeedback({
                title: body.title,
                description: body.description || body.body || body.text,
                category: body.category || body.kind,
                pageUrl: body.pageUrl || body.url,
                userAgent: body.userAgent,
                reporterEmail: caller.upn || '',
                reporterName: caller.name || '',
                tenantId: caller.tid || '',
                schoolName,
                domain,
                attachments: body.attachments
            });

            return jsonResponse(201, {
                ok: true,
                item: {
                    id: item.id,
                    title: item.title,
                    category: item.category,
                    status: item.status,
                    createdAt: item.createdAt
                }
            });
        } catch (e) {
            return errorJsonResponse(context, 'license/feedback POST:', e);
        }
    }
});

app.http('httpFeedbackAdminListOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/admin/feedback',
    handler: async () => corsPreflightResponse()
});

app.http('httpFeedbackAdminList', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/admin/feedback',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const status = request.query.get('status') || '';
            const top = Number(request.query.get('top') || 100);
            const data = await listFeedback({ status: status || undefined, top });
            return jsonResponse(200, {
                ok: true,
                items: data.items,
                listDisplayName: data.listDisplayName,
                categories: CATEGORIES,
                statuses: STATUSES
            });
        } catch (e) {
            return errorJsonResponse(context, 'license/admin/feedback GET:', e, { items: [] });
        }
    }
});

app.http('httpFeedbackAdminUpdateOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/admin/feedback/{id}',
    handler: async () => corsPreflightResponse()
});

app.http('httpFeedbackAdminUpdate', {
    methods: ['PATCH'],
    authLevel: 'anonymous',
    route: 'license/admin/feedback/{id}',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const id = request.params && request.params.id;
            const body = await readJsonBody(request);
            const item = await updateFeedback(id, body);
            return jsonResponse(200, { ok: true, item });
        } catch (e) {
            return errorJsonResponse(context, 'license/admin/feedback PATCH:', e);
        }
    }
});

app.http('httpFeedbackAdminFileOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/admin/feedback/file',
    handler: async () => corsPreflightResponse()
});

app.http('httpFeedbackAdminFile', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/admin/feedback/file',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const path = request.query.get('path') || '';
            const file = await readFeedbackFile(path);
            return {
                status: 200,
                headers: {
                    'Content-Type': file.contentType,
                    'Cache-Control': 'private, max-age=120',
                    'Access-Control-Allow-Origin': '*',
                    'Access-Control-Allow-Headers': 'Content-Type, Authorization, x-functions-key'
                },
                body: file.bytes
            };
        } catch (e) {
            return errorJsonResponse(context, 'license/admin/feedback/file GET:', e);
        }
    }
});

app.http('httpFeedbackEnsureSchemaOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/admin/feedback/ensure-schema',
    handler: async () => corsPreflightResponse()
});

app.http('httpFeedbackEnsureSchema', {
    methods: ['POST'],
    authLevel: 'anonymous',
    route: 'license/admin/feedback/ensure-schema',
    handler: async (request, context) => {
        try {
            await requireOperatorCaller(bearerTokenFromRequest(request));
            const schema = await ensureFeedbackSchema();
            return jsonResponse(200, { ok: true, schema });
        } catch (e) {
            return errorJsonResponse(context, 'license/admin/feedback/ensure-schema POST:', e);
        }
    }
});
