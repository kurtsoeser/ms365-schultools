'use strict';

const { app } = require('@azure/functions');
const { getJob, saveJob } = require('../lib/job-store');
const { createSingleKursteam } = require('../lib/kursteam-create');
const { sleep } = require('../lib/graph-client');

app.storageQueue('queueProcessJob', {
    queueName: '%KURSTEAMS_JOB_QUEUE%',
    connection: 'AzureWebJobsStorage',
    handler: async (message, context) => {
        let jobId;
        try {
            const raw =
                typeof message === 'string'
                    ? message
                    : message && typeof message === 'object' && message.jobId
                      ? message
                      : String(message || '');
            const parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
            jobId = parsed && parsed.jobId;
        } catch (e) {
            context.error('Queue-Nachricht ungültig:', message, e);
            throw e;
        }

        if (!jobId) {
            context.error('Queue-Nachricht ohne jobId:', message);
            return;
        }

        const job = await getJob(jobId);
        if (!job) {
            context.error('Job nicht gefunden:', jobId);
            return;
        }

        if (job.status === 'completed' || job.status === 'failed') {
            context.log('Job bereits abgeschlossen, überspringe:', jobId);
            return;
        }

        job.status = 'running';
        await saveJob(job);
        context.log('Verarbeite Kursteam-Job', jobId, 'Teams:', job.total);

        for (const entry of job.entries) {
            if (entry.status === 'ok') {
                continue;
            }

            entry.status = 'running';
            await saveJob(job);

            const log = (msg) => context.log('[' + entry.index + '/' + job.total + '] ' + msg);

            try {
                const result = await createSingleKursteam(
                    {
                        teamName: entry.teamName,
                        gruppenmail: entry.gruppenmail,
                        besitzer: entry.besitzer
                    },
                    log,
                    job.tenantId
                );
                entry.status = 'ok';
                entry.groupId = result.groupId;
                entry.message = 'OK → ' + entry.gruppenmail;
                job.completed++;
            } catch (e) {
                entry.status = 'error';
                entry.message = e.message || String(e);
                job.failed++;
                context.error('Fehler bei', entry.teamName, e);
            }

            await saveJob(job);
            await sleep(2000);
        }

        job.status = job.failed > 0 && job.completed === 0 ? 'failed' : 'completed';
        await saveJob(job);
        context.log(
            'Job fertig:',
            jobId,
            'OK:',
            job.completed,
            'Fehler:',
            job.failed
        );
    }
});
