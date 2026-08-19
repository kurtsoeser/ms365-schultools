'use strict';

const { app } = require('@azure/functions');
const { getJob } = require('../lib/job-store');
const { jsonResponse, corsPreflightResponse } = require('../lib/http-utils');

app.http('httpGetJobOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'kursteams/jobs/{jobId}',
    handler: async () => corsPreflightResponse()
});

app.http('httpGetJob', {
    methods: ['GET'],
    authLevel: 'function',
    route: 'kursteams/jobs/{jobId}',
    handler: async (request, context) => {
        const jobId = request.params.jobId;
        if (!jobId) {
            return jsonResponse(400, { error: 'jobId fehlt.' });
        }

        try {
            const job = await getJob(jobId);
            if (!job) {
                return jsonResponse(404, { error: 'Job nicht gefunden.' });
            }
            return jsonResponse(200, job);
        } catch (e) {
            context.error('getJob fehlgeschlagen:', e);
            return jsonResponse(500, { error: e.message || String(e) });
        }
    }
});
