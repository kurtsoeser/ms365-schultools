/**
 * SharePoint-CRUD für Raumplan / Raeume (Live-Anbindung).
 * Transport: shared/graph-client.js (R22 / b11); Token/Site-Helfer weiter über spo-graph-shared.
 */
import { graphJson as sharedGraphJson } from '../../shared/graph-client.js';

import { LIST_TITLES } from './raumplan-schema.js';
import { normalizeRaum } from './raumplan-model.js';
import { loadRaumplanSpSettings, resolveRaumplanSiteUrl } from './raumplan-settings.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

async function graphJson(method, pathOrUrl, token, body, version) {
    return sharedGraphJson(method, pathOrUrl, token, body, undefined, {
        version: version === 'beta' ? 'beta' : 'v1.0'
    });
}

async function token() {
    return await G().getGraphToken(SCOPES);
}

function fieldStr(fields, key) {
    const v = fields && fields[key];
    return v == null ? '' : String(v).trim();
}

function fieldNum(fields, key) {
    const v = fields && fields[key];
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
}

async function findListByTitle(tok, siteId, listTitle) {
    const title = String(listTitle || '').trim();
    const path =
        G().graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await graphJson('GET', path, tok, undefined, 'v1.0');
    return ((data && data.value) || [])[0] || null;
}

async function fetchAllItems(tok, siteId, listId) {
    let path =
        G().graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(listId) +
        '/items?$expand=fields&$top=100';
    const out = [];
    while (path) {
        const data = await graphJson(
            'GET',
            path.indexOf('http') === 0 ? path : path,
            tok,
            undefined,
            'v1.0'
        );
        const rows = (data && data.value) || [];
        for (let i = 0; i < rows.length; i++) out.push(rows[i]);
        path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
    }
    return out;
}

/**
 * @param {object} item Graph list item
 */
export function mapRaumFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeRaum({
        id: fieldStr(f, 'RaumId') || (item && item.id != null ? 'spo-' + item.id : ''),
        code: fieldStr(f, 'Code'),
        name: fieldStr(f, 'Name') || fieldStr(f, 'Title'),
        gebaeude: fieldStr(f, 'Gebaeude'),
        stockwerk: fieldStr(f, 'Stockwerk'),
        kapazitaet: fieldNum(f, 'Kapazitaet'),
        typ: fieldStr(f, 'Typ'),
        notiz: fieldStr(f, 'Notiz'),
        seedTag: fieldStr(f, 'SeedTag'),
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/**
 * @param {object} raum
 * @param {{ omitOptional?: boolean }} [opts]
 */
export function mapRaumToFields(raum, opts = {}) {
    const r = normalizeRaum(raum);
    /** @type {Record<string, string|number>} */
    const fields = {
        Title: r.name || r.code || 'Raum',
        RaumId: r.id,
        Code: r.code,
        Name: r.name,
        Gebaeude: r.gebaeude,
        Stockwerk: r.stockwerk,
        Typ: r.typ,
        Notiz: r.notiz
    };
    if (r.kapazitaet != null && Number.isFinite(Number(r.kapazitaet))) {
        fields.Kapazitaet = Number(r.kapazitaet);
    }
    if (!opts.omitOptional) {
        if (r.seedTag) fields.SeedTag = r.seedTag;
    }
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = r.name || r.code || 'Raum';
    if (r.id) fields.RaumId = r.id;
    if (r.code) fields.Code = r.code;
    return fields;
}

/**
 * @param {string} [webUrl]
 */
export async function resolveRaumplanContext(webUrl) {
    const settings = loadRaumplanSpSettings();
    const url = String(webUrl || settings.siteUrl || resolveRaumplanSiteUrl() || '').trim();
    if (!url) throw new Error('Keine SharePoint-Site hinterlegt. Bitte Listen-Assistent oder System nutzen.');
    const listTitle = settings.listRaeume || LIST_TITLES.raeume;
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const list = await findListByTitle(tok, siteId, listTitle);
    if (!list) {
        throw new Error(
            'Liste „' + listTitle + '“ fehlt. Bitte zuerst unter „Listen einrichten“ anlegen.'
        );
    }
    return {
        webUrl: url,
        siteId,
        siteName: site.displayName || '',
        list: {
            id: String(list.id),
            webUrl: list.webUrl || '',
            title: list.displayName || listTitle
        }
    };
}

/**
 * @param {{ webUrl?: string }} [opts]
 */
export async function loadAllRaeume(opts = {}) {
    const ctx = await resolveRaumplanContext(opts.webUrl);
    const tok = await token();
    const rows = await fetchAllItems(tok, ctx.siteId, ctx.list.id);
    return { ctx, items: rows.map(mapRaumFromItem) };
}

/**
 * @param {object} ctx
 * @param {object} raum
 */
export async function createRaumItem(ctx, raum) {
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(ctx.list.id) + '/items';
    let fields = mapRaumToFields(raum, { omitOptional: false });
    let created;
    try {
        created = await graphJson('POST', path, tok, { fields }, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/SeedTag|RaumId/i.test(msg)) {
            fields = mapRaumToFields(raum, { omitOptional: true });
            created = await graphJson('POST', path, tok, { fields }, 'v1.0');
        } else {
            throw e;
        }
    }
    const mapped = mapRaumFromItem(created);
    const local = normalizeRaum(raum);
    return {
        ...local,
        ...mapped,
        id: local.id || mapped.id,
        spoItemId: mapped.spoItemId || local.spoItemId
    };
}

/**
 * @param {object} ctx
 * @param {object} raum
 */
export async function updateRaumItem(ctx, raum) {
    const spoId = String(raum.spoItemId || '').trim();
    if (!spoId) throw new Error('spoItemId fehlt – Raum noch nicht in SharePoint.');
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.list.id) +
        '/items/' +
        encodeURIComponent(spoId) +
        '/fields';
    let fields = mapRaumToFields(raum, { omitOptional: false });
    try {
        await graphJson('PATCH', path, tok, fields, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/SeedTag|RaumId/i.test(msg)) {
            fields = mapRaumToFields(raum, { omitOptional: true });
            await graphJson('PATCH', path, tok, fields, 'v1.0');
        } else {
            throw e;
        }
    }
    return { ...normalizeRaum(raum), spoItemId: spoId };
}

/**
 * @param {object} ctx
 * @param {string} spoItemId
 */
export async function deleteRaumItem(ctx, spoItemId) {
    const spoId = String(spoItemId || '').trim();
    if (!spoId) return;
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.list.id) +
        '/items/' +
        encodeURIComponent(spoId);
    await graphJson('DELETE', path, tok, undefined, 'v1.0');
}

export { SCOPES as RAUMPLAN_GRAPH_SCOPES };
