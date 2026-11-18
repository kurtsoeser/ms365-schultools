'use strict';

const { graphConfigured } = require('./graph');

/** @type {Map<string, object>} */
const g = globalThis;
if (!g.__sspMemoryStore) {
    g.__sspMemoryStore = new Map();
}

function useSharePoint() {
    const mode = String(process.env.STAMMDATEN_STORE || '').toLowerCase();
    if (mode === 'memory') return false;
    if (mode === 'sharepoint') return true;
    return graphConfigured();
}

function memory() {
    return g.__sspMemoryStore;
}

async function getByGuid(guid) {
    if (useSharePoint()) {
        return require('./sharepoint-store').getByGuid(guid);
    }
    return memory().get(String(guid || '')) || null;
}

async function upsert(record) {
    if (!record?.uniqueGuid) throw new Error('uniqueGuid fehlt');
    if (useSharePoint()) {
        return require('./sharepoint-store').upsert(record);
    }
    memory().set(String(record.uniqueGuid), { ...record });
    return record;
}

async function listAll() {
    if (useSharePoint()) {
        return require('./sharepoint-store').listAll();
    }
    return [...memory().values()];
}

/**
 * @param {{ siteUrl?: string, siteId?: string, listName?: string }} [opts]
 */
async function ensureList(opts = {}) {
    if (!useSharePoint()) {
        return {
            mode: 'memory',
            ok: true,
            hint:
                'Gateway OK, aber Store=memory – in SharePoint wurde nichts angelegt. Azure App Setting STAMMDATEN_STORE=sharepoint setzen und Graph-Credentials (Tenant/Client/Secret) hinterlegen, dann erneut „Liste anlegen / prüfen“.'
        };
    }
    return require('./sharepoint-store').ensureList(opts);
}

function backendMode() {
    return useSharePoint() ? 'sharepoint' : 'memory';
}

module.exports = {
    getByGuid,
    upsert,
    listAll,
    ensureList,
    backendMode
};
