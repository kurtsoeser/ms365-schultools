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
const { lookupLicenseItem, updateLicense, pickToolSiteUrlFromFields } = require('../lib/sharepoint-license');
const {
    evaluateLicense,
    upnMatchesDesignatedItContact,
    canReportToolSiteUrl,
    normalizeReportedToolSiteUrl
} = require('../lib/evaluate-license');
const {
    pickSchoolStammdatenFromFields,
    normalizeSchoolProfileMePatch
} = require('../lib/school-profile-sync');

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
            const hit = await lookupLicenseItem(caller.tid);
            const fields = hit ? hit.fields : null;
            const result = evaluateLicense({
                tenantId: caller.tid,
                fields,
                allowedStatuses: cfg.allowedStatuses
            });

            const contactEmails = result.contactEmails || [];
            const isDesignatedSchoolIt =
                result.allowed &&
                upnMatchesDesignatedItContact(caller.upn, contactEmails);

            const schoolStammdaten =
                (result.schoolStammdaten && Object.keys(result.schoolStammdaten).length
                    ? result.schoolStammdaten
                    : null) || pickSchoolStammdatenFromFields(fields);

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
                    /** Nur URL-Zeiger auf Schul-SharePoint – keine Listeninhalte. */
                    toolSiteUrl: result.toolSiteUrl || null,
                    /** Adresse/Rollen/Kontakt aus Betreiber-Liste (kein Schülerdaten). */
                    schoolStammdaten: schoolStammdaten || {}
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

/**
 * Schul-IT meldet Tool-Site-URL und/oder Schulprofil-Stammdaten an den Betreiber.
 * Body: { toolSiteUrl?: string, schoolName?: string, schoolStammdaten?: { Strasse, PLZ, … } }
 */
app.http('httpLicenseMePatch', {
    methods: ['PATCH'],
    authLevel: 'anonymous',
    route: 'license/me',
    handler: async (request, context) => {
        try {
            const cfg = getConfig();
            const token = bearerTokenFromRequest(request);
            const caller = await validateCallerToken(token, cfg.tokenAudiences);

            let body = {};
            try {
                body = (await request.json()) || {};
            } catch {
                body = {};
            }

            const hasToolSite =
                body.toolSiteUrl != null || body.ToolSiteUrl != null;
            const hasProfile =
                body.schoolName != null ||
                body.Title != null ||
                (body.schoolStammdaten && typeof body.schoolStammdaten === 'object') ||
                (body.extra && typeof body.extra === 'object');

            if (!hasToolSite && !hasProfile) {
                return jsonResponse(400, {
                    ok: false,
                    reason: 'empty',
                    message:
                        'Bitte toolSiteUrl und/oder schoolName bzw. schoolStammdaten angeben.'
                });
            }

            /** @type {{ ok: true, url: string } | { ok: false, reason: string } | null} */
            let normalizedTool = null;
            if (hasToolSite) {
                normalizedTool = normalizeReportedToolSiteUrl(
                    body.toolSiteUrl != null ? body.toolSiteUrl : body.ToolSiteUrl
                );
                if (!normalizedTool.ok) {
                    return jsonResponse(400, {
                        ok: false,
                        reason: normalizedTool.reason,
                        message:
                            normalizedTool.reason === 'empty'
                                ? 'toolSiteUrl fehlt.'
                                : 'Bitte eine gültige https-SharePoint-Webadresse angeben (…/sites/… oder …/teams/…).'
                    });
                }
            }

            /** @type {{ ok: true, patch: Record<string, unknown> } | { ok: false, reason: string } | null} */
            let profileNorm = null;
            if (hasProfile) {
                profileNorm = normalizeSchoolProfileMePatch(body);
                if (!profileNorm.ok) {
                    return jsonResponse(400, {
                        ok: false,
                        reason: profileNorm.reason,
                        message:
                            profileNorm.reason === 'empty_school_name'
                                ? 'Schulname darf nicht leer sein.'
                                : 'Keine gültigen Profilfelder zum Speichern.'
                    });
                }
            }

            const hit = await lookupLicenseItem(caller.tid);
            if (!hit || !hit.id) {
                return jsonResponse(404, {
                    ok: false,
                    reason: 'tenant_not_found',
                    message: 'Für diesen Tenant ist noch kein Lizenz-Eintrag vorhanden.'
                });
            }

            const result = evaluateLicense({
                tenantId: caller.tid,
                fields: hit.fields,
                allowedStatuses: cfg.allowedStatuses
            });
            const contactEmails = result.contactEmails || [];
            const authz = canReportToolSiteUrl({
                allowed: result.allowed,
                contactEmails,
                upn: caller.upn
            });
            if (!authz.ok) {
                const msg =
                    authz.reason === 'not_designated_it'
                        ? 'Nur hinterlegte IT-Kontakte dürfen Site-URL und Schulprofil melden. Bitte den Betreiber um Eintrag unter IT-Kontakt bitten.'
                        : 'Lizenz nicht freigeschaltet – Meldung nicht möglich.';
                return jsonResponse(403, {
                    ok: false,
                    reason: authz.reason,
                    message: msg
                });
            }

            /** @type {Record<string, unknown>} */
            const updateBody = {};
            let toolUnchanged = true;
            let previousTool = null;

            if (normalizedTool && normalizedTool.ok) {
                previousTool =
                    pickToolSiteUrlFromFields(hit.fields) ||
                    result.toolSiteUrl ||
                    null;
                const nextUrl = normalizedTool.url;
                if (previousTool && String(previousTool).replace(/\/+$/, '') === nextUrl) {
                    toolUnchanged = true;
                } else {
                    toolUnchanged = false;
                    updateBody.toolSiteUrl = nextUrl;
                }
            }

            if (profileNorm && profileNorm.ok) {
                Object.assign(updateBody, profileNorm.patch);
            }

            if (!Object.keys(updateBody).length) {
                return jsonResponse(200, {
                    ok: true,
                    unchanged: true,
                    toolSiteUrl: previousTool || (normalizedTool && normalizedTool.url) || result.toolSiteUrl || null,
                    previousToolSiteUrl: previousTool,
                    schoolName: result.schoolName,
                    schoolStammdaten: result.schoolStammdaten || pickSchoolStammdatenFromFields(hit.fields),
                    authz: authz.reason,
                    reportedBy: caller.upn || null
                });
            }

            const updated = await updateLicense(hit.id, updateBody);
            const stammdatenFromUpdated = pickSchoolStammdatenFromFields(
                Object.assign(
                    {},
                    hit.fields || {},
                    (updated && updated.extra) || {},
                    updateBody.extra || {}
                )
            );

            return jsonResponse(200, {
                ok: true,
                unchanged: false,
                toolSiteUrl:
                    (updated && updated.toolSiteUrl) ||
                    updateBody.toolSiteUrl ||
                    previousTool ||
                    result.toolSiteUrl ||
                    null,
                previousToolSiteUrl: previousTool,
                toolSiteUnchanged: toolUnchanged,
                schoolName: (updated && updated.schoolName) || updateBody.schoolName || result.schoolName,
                schoolStammdaten: stammdatenFromUpdated,
                authz: authz.reason,
                reportedBy: caller.upn || null
            });
        } catch (e) {
            const status = e.status && Number.isFinite(e.status) ? e.status : 500;
            if (status >= 500) {
                context.error('license/me PATCH fehlgeschlagen:', e);
            }
            return jsonResponse(status >= 400 && status < 600 ? status : 500, {
                ok: false,
                reason: status === 401 ? 'unauthorized' : 'error',
                message: publicErrorMessage(
                    e,
                    status === 401 ? 'Anmeldung fehlt.' : 'Meldung konnte nicht gespeichert werden.'
                )
            });
        }
    }
});
