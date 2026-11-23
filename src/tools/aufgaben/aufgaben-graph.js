/**
 * SharePoint-CRUD für Aufgaben (Live-Anbindung).
 * Transport: shared/graph-client.js (R22 / b11); Token/Site-Helfer weiter über spo-graph-shared.
 */
import { graphJson as sharedGraphJson } from '../../shared/graph-client.js';

import { LIST_TITLES } from './aufgaben-schema.js';
import { normalizeAufgabe } from './aufgaben-model.js';
import { loadAufgabenSpSettings, resolveAufgabenSiteUrl } from './aufgaben-settings.js';

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
export function mapAufgabeFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeAufgabe({
        id: fieldStr(f, 'AufgabeId') || (item && item.id != null ? 'spo-' + item.id : ''),
        title: fieldStr(f, 'Title'),
        beschreibung: fieldStr(f, 'Beschreibung'),
        zuweisungName: fieldStr(f, 'ZuweisungName'),
        zuweisungEmail: fieldStr(f, 'ZuweisungEmail'),
        rolle: fieldStr(f, 'Rolle'),
        prioritaet: fieldStr(f, 'Prioritaet'),
        frist: fieldStr(f, 'Frist'),
        status: fieldStr(f, 'Status'),
        erstelltVon: fieldStr(f, 'ErstelltVon'),
        erstelltVonEmail: fieldStr(f, 'ErstelltVonEmail'),
        erstelltAm: fieldStr(f, 'ErstelltAm') || (item && item.createdDateTime) || '',
        kommentar: fieldStr(f, 'Kommentar'),
        bearbeitetVonName: fieldStr(f, 'BearbeitetVonName'),
        bearbeitetVonEmail: fieldStr(f, 'BearbeitetVonEmail'),
        bearbeitetAm: fieldStr(f, 'BearbeitetAm'),
        seedTag: fieldStr(f, 'SeedTag'),
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/**
 * @param {object} aufgabe
 * @param {{ omitOptional?: boolean }} [opts]
 */
export function mapAufgabeToFields(aufgabe, opts = {}) {
    const a = normalizeAufgabe(aufgabe);
    /** @type {Record<string, string>} */
    const fields = {
        Title: a.title || 'Aufgabe',
        AufgabeId: a.id,
        Beschreibung: a.beschreibung,
        ZuweisungName: a.zuweisungName,
        ZuweisungEmail: a.zuweisungEmail,
        Rolle: a.rolle,
        Prioritaet: a.prioritaet,
        Status: a.status,
        ErstelltVon: a.erstelltVon,
        ErstelltVonEmail: a.erstelltVonEmail,
        Kommentar: a.kommentar
    };
    if (a.frist) fields.Frist = a.frist;
    if (a.erstelltAm) fields.ErstelltAm = a.erstelltAm;
    if (!opts.omitOptional) {
        if (a.bearbeitetVonName) fields.BearbeitetVonName = a.bearbeitetVonName;
        if (a.bearbeitetVonEmail) fields.BearbeitetVonEmail = a.bearbeitetVonEmail;
        if (a.bearbeitetAm) fields.BearbeitetAm = a.bearbeitetAm;
        if (a.seedTag) fields.SeedTag = a.seedTag;
    }
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = a.title || 'Aufgabe';
    if (a.id) fields.AufgabeId = a.id;
    return fields;
}

/**
 * @param {string} [webUrl]
 */
export async function resolveAufgabenContext(webUrl) {
    const settings = loadAufgabenSpSettings();
    const url = String(webUrl || settings.siteUrl || resolveAufgabenSiteUrl() || '').trim();
    if (!url) throw new Error('Keine SharePoint-Site hinterlegt. Bitte Listen-Assistent oder System nutzen.');
    const listTitle = settings.listAufgaben || LIST_TITLES.aufgaben;
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
export async function loadAllAufgaben(opts = {}) {
    const ctx = await resolveAufgabenContext(opts.webUrl);
    const tok = await token();
    const rows = await fetchAllItems(tok, ctx.siteId, ctx.list.id);
    return { ctx, items: rows.map(mapAufgabeFromItem) };
}

/**
 * @param {object} ctx
 * @param {object} aufgabe
 */
export async function createAufgabeItem(ctx, aufgabe) {
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(ctx.list.id) + '/items';
    let fields = mapAufgabeToFields(aufgabe, { omitOptional: false });
    let created;
    try {
        created = await graphJson('POST', path, tok, { fields }, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/BearbeitetVon|SeedTag|AufgabeId/i.test(msg)) {
            fields = mapAufgabeToFields(aufgabe, { omitOptional: true });
            created = await graphJson('POST', path, tok, { fields }, 'v1.0');
        } else {
            throw e;
        }
    }
    const mapped = mapAufgabeFromItem(created);
    const local = normalizeAufgabe(aufgabe);
    return {
        ...local,
        ...mapped,
        id: local.id || mapped.id,
        spoItemId: mapped.spoItemId || local.spoItemId
    };
}

/**
 * @param {object} ctx
 * @param {object} aufgabe
 */
export async function updateAufgabeItem(ctx, aufgabe) {
    const spoId = String(aufgabe.spoItemId || '').trim();
    if (!spoId) throw new Error('spoItemId fehlt – Aufgabe noch nicht in SharePoint.');
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.list.id) +
        '/items/' +
        encodeURIComponent(spoId) +
        '/fields';
    let fields = mapAufgabeToFields(aufgabe, { omitOptional: false });
    try {
        await graphJson('PATCH', path, tok, fields, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/BearbeitetVon|SeedTag|AufgabeId/i.test(msg)) {
            fields = mapAufgabeToFields(aufgabe, { omitOptional: true });
            await graphJson('PATCH', path, tok, fields, 'v1.0');
        } else {
            throw e;
        }
    }
    return { ...normalizeAufgabe(aufgabe), spoItemId: spoId };
}

/**
 * @param {object} ctx
 * @param {string} spoItemId
 */
export async function deleteAufgabeItem(ctx, spoItemId) {
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

export { SCOPES as AUFGABEN_GRAPH_SCOPES };
