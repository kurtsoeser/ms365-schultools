'use strict';

const { app } = require('@azure/functions');
const { getConfig } = require('../lib/config');
const {
    jsonResponse,
    corsPreflightResponse,
    bearerTokenFromRequest,
    publicErrorMessage
} = require('../lib/http-utils');
const { validateCallerToken } = require('../lib/validate-token');
const { lookupLicenseFields } = require('../lib/sharepoint-license');
const {
    evaluateLicense,
    upnMatchesDesignatedItContact
} = require('../lib/evaluate-license');

app.http('httpLicenseMeOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/me',
    handler: async () => corsPreflightResponse()
});

app.http('httpLicenseMe', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/me',
    handler: async (request, context) => {
        try {
            const cfg = getConfig();
            const token = bearerTokenFromRequest(request);
            const caller = await validateCallerToken(token, cfg.tokenAudiences);
            const fields = await lookupLicenseFields(caller.tid);
            const result = evaluateLicense({
                tenantId: caller.tid,
                fields,
                allowedStatuses: cfg.allowedStatuses
            });

            const contactEmails = result.contactEmails || [];
            const isDesignatedSchoolIt =
                result.allowed &&
                upnMatchesDesignatedItContact(caller.upn, contactEmails);

            return jsonResponse(200, {
                allowed: result.allowed,
                reason: result.reason,
                message: result.message,
                tenantId: caller.tid,
                isDesignatedSchoolIt,
                user: {
                    oid: caller.oid || null,
                    upn: caller.upn || null,
                    name: caller.name || null
                },
                license: {
                    schoolName: result.schoolName,
                    status: result.status,
                    validUntil: result.validUntil,
                    primaryDomain: result.primaryDomain,
                    domains: result.domains || [],
                    contactEmail: result.contactEmail,
                    contactEmails,
                    /** Nur URL-Zeiger auf Schul-SharePoint – keine Stammdaten-Inhalte. */
                    toolSiteUrl: result.toolSiteUrl || null
                },
                /** Explizit für Clients: wo die Schule ihre Tool-/Stammdaten-Site hat. */
                sites: {
                    toolSiteUrl: result.toolSiteUrl || null
                },
                entitlements: result.entitlements || {
                    packages: [],
                    features: [],
                    denied: [],
                    featureCount: 0
                }
            });
        } catch (e) {
            const status = e.status && Number.isFinite(e.status) ? e.status : 500;
            if (status >= 500) {
                context.error('license/me fehlgeschlagen:', e);
            }
            return jsonResponse(status >= 400 && status < 600 ? status : 500, {
                allowed: false,
                reason: status === 401 ? 'unauthorized' : 'error',
                message: publicErrorMessage(e, status === 401 ? 'Anmeldung fehlt.' : 'Lizenzprüfung fehlgeschlagen.'),
                tenantId: null,
                user: null,
                license: null
            });
        }
    }
});
