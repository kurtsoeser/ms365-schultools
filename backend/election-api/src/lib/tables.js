'use strict';

const { TableClient, TableServiceClient } = require('@azure/data-tables');
const { createHash, randomUUID } = require('node:crypto');
const { getConfig } = require('./config');

const TABLE_ELECTIONS = 'ElectionMeta';
const TABLE_VOTERS = 'ElectionVoters';
const TABLE_BALLOTS = 'ElectionBallots';

/** @type {Map<string, TableClient>} */
const clients = new Map();
let ensured = false;

/**
 * @param {string} name
 */
function tableClient(name) {
    const cfg = getConfig();
    const conn = cfg.storageConnection;
    if (!conn) {
        const err = new Error('Azure Storage Connection fehlt (AzureWebJobsStorage).');
        err.status = 503;
        err.code = 'storage_config';
        throw err;
    }
    if (!clients.has(name)) {
        clients.set(name, TableClient.fromConnectionString(conn, name));
    }
    return clients.get(name);
}

async function ensureTables() {
    if (ensured) return;
    const cfg = getConfig();
    const conn = cfg.storageConnection;
    if (!conn) return;
    const svc = TableServiceClient.fromConnectionString(conn);
    for (const name of [TABLE_ELECTIONS, TABLE_VOTERS, TABLE_BALLOTS]) {
        try {
            await svc.createTable(name);
        } catch (e) {
            const code = e && (e.statusCode || e.code);
            if (code !== 409 && String(e && e.message || '').indexOf('TableAlreadyExists') < 0) {
                // Azurite / bestehende Tabelle
                if (Number(code) !== 409) {
                    /* ignore already exists variants */
                }
            }
        }
    }
    ensured = true;
}

/**
 * @param {string} tid
 * @param {string} electionId
 * @param {string} roundId
 * @param {string} oid
 */
function voterKey(tid, electionId, roundId, oid) {
    const cfg = getConfig();
    const payload = [
        cfg.voterPepper,
        String(tid || '').toLowerCase(),
        String(electionId || ''),
        String(roundId || ''),
        String(oid || '').toLowerCase()
    ].join('|');
    return createHash('sha256').update(payload, 'utf8').digest('hex');
}

function pkVoters(tid, electionId, roundId) {
    return `${String(tid).toLowerCase()}_${electionId}_${roundId}`;
}

/**
 * @param {string} tid
 */
async function listElections(tid) {
    await ensureTables();
    const client = tableClient(TABLE_ELECTIONS);
    const out = [];
    const iter = client.listEntities({
        queryOptions: { filter: `PartitionKey eq '${escapeOData(String(tid).toLowerCase())}'` }
    });
    for await (const ent of iter) {
        out.push(parseElectionEntity(ent));
    }
    out.sort((a, b) => String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')));
    return out;
}

/**
 * @param {string} tid
 * @param {string} electionId
 */
async function getElection(tid, electionId) {
    await ensureTables();
    const client = tableClient(TABLE_ELECTIONS);
    try {
        const ent = await client.getEntity(String(tid).toLowerCase(), String(electionId));
        return parseElectionEntity(ent);
    } catch (e) {
        if (e && (e.statusCode === 404 || e.code === 'ResourceNotFound')) return null;
        throw e;
    }
}

/**
 * @param {object} election
 * @param {string} tid
 * @param {string[]} adminOids
 */
async function upsertElection(election, tid, adminOids) {
    await ensureTables();
    const client = tableClient(TABLE_ELECTIONS);
    const now = new Date().toISOString();
    const payload = Object.assign({}, election, {
        updatedAt: now,
        createdAt: election.createdAt || now
    });
    const entity = {
        partitionKey: String(tid).toLowerCase(),
        rowKey: String(election.id),
        payloadJson: JSON.stringify(payload),
        adminOidsJson: JSON.stringify(adminOids || []),
        updatedAt: now
    };
    await client.upsertEntity(entity, 'Replace');
    return payload;
}

/**
 * @param {string} tid
 * @param {string} electionId
 */
async function deleteElection(tid, electionId) {
    await ensureTables();
    const eClient = tableClient(TABLE_ELECTIONS);
    try {
        await eClient.deleteEntity(String(tid).toLowerCase(), String(electionId));
    } catch (e) {
        if (!(e && (e.statusCode === 404 || e.code === 'ResourceNotFound'))) throw e;
    }
    // Voters + Ballots: Partitionen pro Round – wir scannen Prefix
    await deleteByPartitionPrefix(TABLE_VOTERS, `${String(tid).toLowerCase()}_${electionId}_`);
    await deleteByPartitionPrefix(TABLE_BALLOTS, `${String(tid).toLowerCase()}_${electionId}_`);
}

/**
 * @param {string} tableName
 * @param {string} prefix
 */
async function deleteByPartitionPrefix(tableName, prefix) {
    const client = tableClient(tableName);
    const seen = new Set();
    const iter = client.listEntities({
        queryOptions: { select: ['partitionKey', 'rowKey'] }
    });
    for await (const ent of iter) {
        const pk = String(ent.partitionKey || '');
        if (!pk.startsWith(prefix)) continue;
        const key = pk + '|' + ent.rowKey;
        if (seen.has(key)) continue;
        seen.add(key);
        try {
            await client.deleteEntity(pk, String(ent.rowKey));
        } catch {
            /* ignore */
        }
    }
}

/**
 * @param {string} tid
 * @param {string} electionId
 * @param {string} roundId
 * @param {string} oid
 */
async function hasVoted(tid, electionId, roundId, oid) {
    await ensureTables();
    const client = tableClient(TABLE_VOTERS);
    const rk = voterKey(tid, electionId, roundId, oid);
    try {
        await client.getEntity(pkVoters(tid, electionId, roundId), rk);
        return true;
    } catch (e) {
        if (e && (e.statusCode === 404 || e.code === 'ResourceNotFound')) return false;
        throw e;
    }
}

/**
 * Atomar: Wähler-Vermerk anlegen (409 wenn schon vorhanden), dann Ballot.
 * @param {object} args
 */
async function castAnonymousVote(args) {
    await ensureTables();
    const { tid, electionId, roundId, oid, ballotPayload } = args;
    const vClient = tableClient(TABLE_VOTERS);
    const bClient = tableClient(TABLE_BALLOTS);
    const pk = pkVoters(tid, electionId, roundId);
    const rk = voterKey(tid, electionId, roundId, oid);
    const at = new Date().toISOString();

    try {
        await vClient.createEntity({
            partitionKey: pk,
            rowKey: rk,
            at
        });
    } catch (e) {
        if (e && (e.statusCode === 409 || e.code === 'EntityAlreadyExists')) {
            const err = new Error('Sie haben in diesem Wahlgang bereits abgestimmt.');
            err.status = 409;
            err.code = 'already_voted';
            throw err;
        }
        throw e;
    }

    const ballotId = randomUUID().replace(/-/g, '');
    await bClient.createEntity({
        partitionKey: pk,
        rowKey: ballotId,
        payloadJson: JSON.stringify(ballotPayload),
        at
    });

    return { ok: true, at };
}

/**
 * @param {string} tid
 * @param {string} electionId
 * @param {string} roundId
 */
async function listBallots(tid, electionId, roundId) {
    await ensureTables();
    const client = tableClient(TABLE_BALLOTS);
    const pk = pkVoters(tid, electionId, roundId);
    const out = [];
    const iter = client.listEntities({
        queryOptions: { filter: `PartitionKey eq '${escapeOData(pk)}'` }
    });
    for await (const ent of iter) {
        let payload = {};
        try {
            payload = JSON.parse(String(ent.payloadJson || '{}'));
        } catch {
            payload = {};
        }
        out.push({
            id: String(ent.rowKey),
            rankings: payload.rankings,
            choice: payload.choice,
            at: String(ent.at || '')
        });
    }
    return out;
}

/**
 * @param {string} tid
 * @param {string} electionId
 * @param {string} roundId
 */
async function countVoters(tid, electionId, roundId) {
    await ensureTables();
    const client = tableClient(TABLE_VOTERS);
    const pk = pkVoters(tid, electionId, roundId);
    let n = 0;
    const iter = client.listEntities({
        queryOptions: {
            filter: `PartitionKey eq '${escapeOData(pk)}'`,
            select: ['rowKey']
        }
    });
    for await (const _ent of iter) n += 1;
    return n;
}

/**
 * @param {Record<string, unknown>} ent
 */
function parseElectionEntity(ent) {
    let election = {};
    let adminOids = [];
    try {
        election = JSON.parse(String(ent.payloadJson || '{}'));
    } catch {
        election = {};
    }
    try {
        adminOids = JSON.parse(String(ent.adminOidsJson || '[]'));
    } catch {
        adminOids = [];
    }
    if (!Array.isArray(adminOids)) adminOids = [];
    election._adminOids = adminOids.map((x) => String(x).toLowerCase());
    return election;
}

/**
 * @param {string} value
 */
function escapeOData(value) {
    return String(value).replace(/'/g, "''");
}

/**
 * @param {object} election
 * @param {string} oid
 */
function isElectionAdmin(election, oid) {
    const list = Array.isArray(election && election._adminOids) ? election._adminOids : [];
    return list.includes(String(oid || '').toLowerCase());
}

module.exports = {
    listElections,
    getElection,
    upsertElection,
    deleteElection,
    hasVoted,
    castAnonymousVote,
    listBallots,
    countVoters,
    voterKey,
    isElectionAdmin,
    ensureTables,
    TABLE_ELECTIONS,
    TABLE_VOTERS,
    TABLE_BALLOTS
};
