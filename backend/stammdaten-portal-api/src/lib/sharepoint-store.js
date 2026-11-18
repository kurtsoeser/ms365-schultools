'use strict';

const { env, graphJson, resolveSiteId } = require('./graph');

const LIST_NAME = () => env('STAMMDATEN_SP_LIST_NAME', 'SchuelerAnmeldungen');

const COLUMN_DEFS = [
    { name: 'UniqueGUID', typed: { text: {} } },
    { name: 'AccessPinHash', typed: { text: {} } },
    {
        name: 'Status',
        typed: {
            choice: {
                allowTextEntry: false,
                choices: ['Vorerfasst', 'In Bearbeitung', 'Vollständig', 'Geprüft', 'Importiert']
            }
        }
    },
    { name: 'PayloadJSON', typed: { text: { allowMultipleLines: true, textType: 'plain' } } },
    { name: 'FailedAttempts', typed: { number: {} } },
    { name: 'LockedUntil', typed: { text: {} } },
    { name: 'CreatedAt', typed: { text: {} } },
    { name: 'UpdatedAt', typed: { text: {} } },
    { name: 'ReviewedAt', typed: { text: {} } },
    { name: 'ReviewedBy', typed: { text: {} } },
    { name: 'ReviewedByName', typed: { text: {} } }
];

/** @type {Map<string, string>} */
const listIdCache = new Map();

/**
 * @param {string} siteId
 * @param {{ listName?: string, listId?: string, skipCache?: boolean }} [opts]
 */
async function getListId(siteId, opts = {}) {
    const fixed = String(opts.listId || env('STAMMDATEN_SP_LIST_ID') || '').trim();
    if (fixed) return fixed;
    const name = String(opts.listName || LIST_NAME()).trim() || 'SchuelerAnmeldungen';
    const cacheKey = `${siteId}::${name}`;
    if (!opts.skipCache && listIdCache.has(cacheKey)) return listIdCache.get(cacheKey);
    const data = await graphJson(
        'GET',
        `/sites/${encodeURIComponent(siteId)}/lists?$filter=displayName eq '${name.replace(/'/g, "''")}'`
    );
    const list = (data.value || [])[0];
    if (!list?.id) {
        throw Object.assign(new Error(`SharePoint-Liste „${name}“ nicht gefunden. Bitte Liste anlegen.`), {
            status: 404,
            code: 'list_missing'
        });
    }
    if (!opts.skipCache) listIdCache.set(cacheKey, list.id);
    return list.id;
}

function fieldsToRecord(item) {
    const f = item.fields || {};
    let payload = {};
    try {
        payload = f.PayloadJSON ? JSON.parse(f.PayloadJSON) : {};
    } catch {
        payload = {};
    }
    return {
        itemId: item.id,
        uniqueGuid: f.UniqueGUID || '',
        accessPinHash: f.AccessPinHash || '',
        status: f.Status || 'Vorerfasst',
        title: f.Title || '',
        payload,
        failedAttempts: Number(f.FailedAttempts || 0),
        lockedUntil: f.LockedUntil || null,
        createdAt: f.CreatedAt || null,
        updatedAt: f.UpdatedAt || null,
        reviewedAt: f.ReviewedAt || null,
        reviewedBy: f.ReviewedBy || '',
        reviewedByName: f.ReviewedByName || ''
    };
}

function recordToFields(record) {
    return {
        Title: record.title || 'Anmeldung',
        UniqueGUID: record.uniqueGuid,
        AccessPinHash: record.accessPinHash,
        Status: record.status || 'Vorerfasst',
        PayloadJSON: JSON.stringify(record.payload || {}),
        FailedAttempts: Number(record.failedAttempts || 0),
        LockedUntil: record.lockedUntil || '',
        CreatedAt: record.createdAt || '',
        UpdatedAt: record.updatedAt || '',
        ReviewedAt: record.reviewedAt || '',
        ReviewedBy: record.reviewedBy || '',
        ReviewedByName: record.reviewedByName || ''
    };
}

async function getByGuid(guid) {
    const wanted = String(guid || '').trim();
    if (!wanted) return null;
    const siteId = await resolveSiteId();
    const listId = await getListId(siteId);
    const g = wanted.replace(/'/g, "''");
    try {
        const data = await graphJson(
            'GET',
            `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items?$expand=fields&$filter=fields/UniqueGUID eq '${g}'&$top=1`
        );
        const item = (data.value || [])[0];
        if (item) return fieldsToRecord(item);
    } catch {
        /* Filter kann auf manchen Listen scheitern – Fallback scan */
    }
    const all = await listAll();
    return all.find((r) => String(r.uniqueGuid || '') === wanted) || null;
}

async function upsert(record) {
    const siteId = await resolveSiteId();
    const listId = await getListId(siteId);
    const fields = recordToFields(record);
    let itemId = record.itemId;
    if (!itemId) {
        const existing = await getByGuid(record.uniqueGuid);
        itemId = existing?.itemId;
    }
    if (itemId) {
        await graphJson(
            'PATCH',
            `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}/fields`,
            fields
        );
        return { ...record, itemId };
    }
    const created = await graphJson(
        'POST',
        `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`,
        { fields }
    );
    return { ...record, itemId: created.id };
}

async function listAll() {
    const siteId = await resolveSiteId();
    const listId = await getListId(siteId);
    const out = [];
    let url = `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items?$expand=fields&$top=100`;
    while (url) {
        const data = await graphJson('GET', url);
        for (const item of data.value || []) out.push(fieldsToRecord(item));
        url = data['@odata.nextLink'] || null;
        if (url && url.startsWith('https://graph.microsoft.com/v1.0')) {
            url = url.replace('https://graph.microsoft.com/v1.0', '');
        }
    }
    return out;
}

/**
 * @param {{ siteUrl?: string, siteId?: string, listName?: string }} [opts]
 */
async function ensureList(opts = {}) {
    const name = String(opts.listName || LIST_NAME()).trim() || 'SchuelerAnmeldungen';
    const siteId = await resolveSiteId({ siteUrl: opts.siteUrl, siteId: opts.siteId });
    let listId = '';
    try {
        listId = await getListId(siteId, { listName: name, skipCache: true });
    } catch {
        listId = '';
    }
    let created = false;
    if (!listId) {
        const createdList = await graphJson('POST', `/sites/${encodeURIComponent(siteId)}/lists`, {
            displayName: name,
            description:
                'Schüler-Stammdaten-Portal: Anmeldungen mit GUID/PIN. Zugriff nur Sekretariat + App-Registrierung.',
            list: { template: 'genericList' }
        });
        listId = createdList.id;
        created = true;
    }
    listIdCache.set(`${siteId}::${name}`, listId);

    const cols = await graphJson(
        'GET',
        `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/columns?$top=200`
    );
    const existing = new Set((cols.value || []).map((c) => String(c.name || '').toLowerCase()));
    let added = 0;
    for (const def of COLUMN_DEFS) {
        if (existing.has(def.name.toLowerCase())) continue;
        await graphJson('POST', `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/columns`, {
            name: def.name,
            displayName: def.name,
            ...def.typed
        });
        added += 1;
    }

    try {
        await graphJson(
            'PATCH',
            `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/columns/UniqueGUID`,
            { indexed: true }
        );
    } catch {
        /* optional */
    }

    return {
        siteId,
        listId,
        listName: name,
        siteUrl: opts.siteUrl || env('STAMMDATEN_SP_SITE_URL') || '',
        created,
        columnsAdded: added
    };
}

module.exports = {
    getByGuid,
    upsert,
    listAll,
    ensureList,
    LIST_NAME
};
