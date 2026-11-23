/**
 * SharePoint-CRUD für Technik-Meldungen (Live-Anbindung).
 * Transport: shared/graph-client.js (R22 / b11); Token/Site-Helfer weiter über spo-graph-shared.
 */
import { graphJson as sharedGraphJson } from '../../shared/graph-client.js';

import { LIST_TITLES } from './technik-meldungen-schema.js';
import { normalizeMeldung } from './technik-meldungen-model.js';
import { loadTechnikSpSettings, resolveTechnikSiteUrl } from './technik-meldungen-settings.js';

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
export function mapMeldungFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeMeldung({
        id: fieldStr(f, 'MeldungId') || (item && item.id != null ? 'spo-' + item.id : ''),
        title: fieldStr(f, 'Title'),
        beschreibung: fieldStr(f, 'Beschreibung'),
        raumCode: fieldStr(f, 'RaumCode'),
        raumName: fieldStr(f, 'RaumName'),
        kategorie: fieldStr(f, 'Kategorie'),
        prioritaet: fieldStr(f, 'Prioritaet'),
        status: fieldStr(f, 'Status'),
        melderName: fieldStr(f, 'MelderName'),
        melderEmail: fieldStr(f, 'MelderEmail'),
        gemeldetAm: fieldStr(f, 'GemeldetAm') || (item && item.createdDateTime) || '',
        erledigtAm: fieldStr(f, 'ErledigtAm'),
        kommentar: fieldStr(f, 'Kommentar'),
        bearbeitetVonName: fieldStr(f, 'BearbeitetVonName'),
        bearbeitetVonEmail: fieldStr(f, 'BearbeitetVonEmail'),
        bearbeitetAm: fieldStr(f, 'BearbeitetAm'),
        seedTag: fieldStr(f, 'SeedTag'),
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/**
 * @param {object} meldung
 * @param {{ omitOptional?: boolean }} [opts]
 */
export function mapMeldungToFields(meldung, opts = {}) {
    const m = normalizeMeldung(meldung);
    /** @type {Record<string, string>} */
    const fields = {
        Title: m.title || 'Technik-Meldung',
        MeldungId: m.id,
        Beschreibung: m.beschreibung,
        RaumCode: m.raumCode,
        RaumName: m.raumName,
        Kategorie: m.kategorie,
        Prioritaet: m.prioritaet,
        Status: m.status,
        MelderName: m.melderName,
        MelderEmail: m.melderEmail,
        Kommentar: m.kommentar
    };
    if (m.gemeldetAm) fields.GemeldetAm = m.gemeldetAm;
    if (m.erledigtAm) fields.ErledigtAm = m.erledigtAm;
    if (!opts.omitOptional) {
        if (m.bearbeitetVonName) fields.BearbeitetVonName = m.bearbeitetVonName;
        if (m.bearbeitetVonEmail) fields.BearbeitetVonEmail = m.bearbeitetVonEmail;
        if (m.bearbeitetAm) fields.BearbeitetAm = m.bearbeitetAm;
        if (m.seedTag) fields.SeedTag = m.seedTag;
    }
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = m.title || 'Technik-Meldung';
    if (m.id) fields.MeldungId = m.id;
    return fields;
}

/**
 * @param {string} [webUrl]
 */
export async function resolveTechnikContext(webUrl) {
    const settings = loadTechnikSpSettings();
    const url = String(webUrl || settings.siteUrl || resolveTechnikSiteUrl() || '').trim();
    if (!url) throw new Error('Keine SharePoint-Site hinterlegt. Bitte Listen-Assistent oder System nutzen.');
    const listTitle = settings.listMeldungen || LIST_TITLES.meldungen;
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
export async function loadAllMeldungen(opts = {}) {
    const ctx = await resolveTechnikContext(opts.webUrl);
    const tok = await token();
    const rows = await fetchAllItems(tok, ctx.siteId, ctx.list.id);
    return { ctx, items: rows.map(mapMeldungFromItem) };
}

/**
 * @param {object} ctx
 * @param {object} meldung
 */
export async function createMeldungItem(ctx, meldung) {
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(ctx.list.id) + '/items';
    let fields = mapMeldungToFields(meldung, { omitOptional: false });
    let created;
    try {
        created = await graphJson('POST', path, tok, { fields }, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/BearbeitetVon|SeedTag|MeldungId/i.test(msg)) {
            fields = mapMeldungToFields(meldung, { omitOptional: true });
            created = await graphJson('POST', path, tok, { fields }, 'v1.0');
        } else {
            throw e;
        }
    }
    const mapped = mapMeldungFromItem(created);
    const local = normalizeMeldung(meldung);
    return {
        ...local,
        ...mapped,
        id: local.id || mapped.id,
        spoItemId: mapped.spoItemId || local.spoItemId
    };
}

/**
 * @param {object} ctx
 * @param {object} meldung
 */
export async function updateMeldungItem(ctx, meldung) {
    const spoId = String(meldung.spoItemId || '').trim();
    if (!spoId) throw new Error('spoItemId fehlt – Meldung noch nicht in SharePoint.');
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.list.id) +
        '/items/' +
        encodeURIComponent(spoId) +
        '/fields';
    let fields = mapMeldungToFields(meldung, { omitOptional: false });
    try {
        await graphJson('PATCH', path, tok, fields, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/BearbeitetVon|SeedTag|MeldungId/i.test(msg)) {
            fields = mapMeldungToFields(meldung, { omitOptional: true });
            await graphJson('PATCH', path, tok, fields, 'v1.0');
        } else {
            throw e;
        }
    }
    return { ...normalizeMeldung(meldung), spoItemId: spoId };
}

/**
 * @param {object} ctx
 * @param {string} spoItemId
 */
export async function deleteMeldungItem(ctx, spoItemId) {
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

export { SCOPES as TECHNIK_GRAPH_SCOPES };
