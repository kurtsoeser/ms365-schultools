'use strict';

const { app } = require('@azure/functions');
const { isTenantAllowed } = require('../lib/config');
const { createJob } = require('../lib/job-store');
const { jsonResponse, validateTeamsPayload, corsPreflightResponse } = require('../lib/http-utils');

app.http('httpCreateJobOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'kursteams/jobs',
    handler: async () => corsPreflightResponse()
});

app.http('httpCreateJob', {
    methods: ['POST'],
    authLevel: 'function',
    route: 'kursteams/jobs',
    handler: async (request, context) => {
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

        if (!isTenantAllowed(validated.tenantId)) {
            return jsonResponse(403, {
                error: 'tenantId ist für dieses Backend nicht freigeschaltet (KURSTEAMS_ALLOWED_TENANT_IDS).'
            });
        }

        try {
            const job = await createJob({
                tenantId: validated.tenantId,
                teams: validated.teams
            });
            context.log('Kursteam-Job angelegt:', job.id, 'Teams:', job.total);
            return jsonResponse(202, {
                jobId: job.id,
                status: job.status,
                total: job.total,
                pollUrl: '/api/kursteams/jobs/' + job.id
            });
        } catch (e) {
            context.error('createJob fehlgeschlagen:', e);
            return jsonResponse(500, { error: e.message || String(e) });
        }
    }
});
