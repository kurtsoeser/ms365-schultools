'use strict';

const { app } = require('@azure/functions');
const { createJob } = require('../lib/job-store');
const {
    jsonResponse,
    validateTeamsPayload,
    corsPreflightResponse,
    bearerTokenFromRequest,
    errorResponse
} = require('../lib/http-utils');
const {
    requireKursteamCaller,
    assertCallerMayCreateTeams
} = require('../lib/require-kursteam-caller');
const {
    assertTenantLicenseAllowed,
    licenseTokenFromRequest
} = require('../lib/require-tenant-license');

app.http('httpCreateJobOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'kursteams/jobs',
    handler: async () => corsPreflightResponse()
});

app.http('httpCreateJob', {
    methods: ['POST'],
    authLevel: 'anonymous',
    route: 'kursteams/jobs',
    handler: async (request, context) => {
        try {
            const caller = await requireKursteamCaller(bearerTokenFromRequest(request));
            let body;
            try {
                body = await request.json();
            } catch {
                return jsonResponse(400, { error: 'Ungültiges JSON.' });
            }

            const validated = validateTeamsPayload(body);
            if (validated.error) {
                return jsonResponse(400, { error: validated.error });
            }

            await assertCallerMayCreateTeams(caller);
            await assertTenantLicenseAllowed(caller, licenseTokenFromRequest(request));

            const job = await createJob({
                tenantId: caller.tid,
                createdByOid: caller.oid,
                teams: validated.teams,
                mailDomain: validated.mailDomain
            });
            context.log('Kursteam-Job angelegt:', job.id, 'Teams:', job.total);
            return jsonResponse(202, {
                jobId: job.id,
                status: job.status,
                total: job.total,
                pollUrl: '/api/kursteams/jobs/' + job.id
            });
        } catch (e) {
            return errorResponse(context, 'createJob fehlgeschlagen:', e);
        }
    }
});
