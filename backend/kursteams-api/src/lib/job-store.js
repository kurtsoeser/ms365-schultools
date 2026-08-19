'use strict';

const { BlobServiceClient } = require('@azure/storage-blob');
const { QueueClient } = require('@azure/storage-queue');
const { randomUUID } = require('crypto');
const { getConfig } = require('./config');

function getBlobContainer() {
    const cfg = getConfig();
    if (!cfg.storageConnectionString) {
        throw new Error('AzureWebJobsStorage ist nicht konfiguriert.');
    }
    const blob = BlobServiceClient.fromConnectionString(cfg.storageConnectionString);
    return blob.getContainerClient(cfg.blobContainer);
}

function getQueueClient() {
    const cfg = getConfig();
    if (!cfg.storageConnectionString) {
        throw new Error('AzureWebJobsStorage ist nicht konfiguriert.');
    }
    return new QueueClient(cfg.storageConnectionString, cfg.queueName);
}

async function ensureContainer() {
    const container = getBlobContainer();
    await container.createIfNotExists();
    return container;
}

async function ensureQueue() {
    const queue = getQueueClient();
    await queue.createIfNotExists();
    return queue;
}

function blobName(jobId) {
    return jobId + '.json';
}

/**
 * @param {{ tenantId: string, teams: Array<{ teamName: string, gruppenmail: string, besitzer: string }>, mailDomain?: string }} input
 */
async function createJob(input) {
    const id = randomUUID();
    const now = new Date().toISOString();
    const job = {
        id,
        tenantId: input.tenantId,
        mailDomain: String(input.mailDomain || '').trim(),
        status: 'queued',
        createdAt: now,
        updatedAt: now,
        total: input.teams.length,
        completed: 0,
        failed: 0,
        entries: input.teams.map((t, index) => ({
            index: index + 1,
            teamName: t.teamName,
            gruppenmail: t.gruppenmail,
            besitzer: t.besitzer,
            status: 'pending',
            message: '',
            groupId: null
        }))
    };

    const container = await ensureContainer();
    const block = container.getBlockBlobClient(blobName(id));
    await block.upload(JSON.stringify(job, null, 2), Buffer.byteLength(JSON.stringify(job, null, 2)), {
        blobHTTPHeaders: { blobContentType: 'application/json' }
    });

    const queue = await ensureQueue();
    await queue.sendMessage(JSON.stringify({ jobId: id }));

    return job;
}

async function getJob(jobId) {
    const container = await ensureContainer();
    const block = container.getBlockBlobClient(blobName(jobId));
    if (!(await block.exists())) {
        return null;
    }
    const buf = await block.downloadToBuffer();
    return JSON.parse(buf.toString('utf8'));
}

async function saveJob(job) {
    job.updatedAt = new Date().toISOString();
    const container = await ensureContainer();
    const block = container.getBlockBlobClient(blobName(job.id));
    const json = JSON.stringify(job, null, 2);
    await block.upload(json, Buffer.byteLength(json), {
        blobHTTPHeaders: { blobContentType: 'application/json' }
    });
    return job;
}

module.exports = { createJob, getJob, saveJob };
