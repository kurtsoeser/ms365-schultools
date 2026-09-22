'use strict';

const { app } = require('@azure/functions');
const { getJob } = require('../lib/job-store');
const {
    jsonResponse,
    corsPreflightResponse,
    bearerTokenFromRequest,
    errorResponse
} = require('../lib/http-utils');
const { requireKursteamCaller, jobVisibleToCaller } = require('../lib/require-kursteam-caller');

app.http('httpGetJobOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'kursteams/jobs/{jobId}',
    handler: async () => corsPreflightResponse()
});

app.http('httpGetJob', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'kursteams/jobs/{jobId}',
    handler: async (request, context) => {
        const jobId = request.params.jobId;
        if (!jobId) {
            return jsonResponse(400, { error: 'jobId fehlt.' });
        }

        try {
            const caller = await requireKursteamCaller(bearerTokenFromRequest(request));
            const job = await getJob(jobId);
            if (!jobVisibleToCaller(job, caller)) {
                return jsonResponse(404, { error: 'Job nicht gefunden.' });
            }
            return jsonResponse(200, job);
        } catch (e) {
            return errorResponse(context, 'getJob fehlgeschlagen:', e);
        }
    }
});
