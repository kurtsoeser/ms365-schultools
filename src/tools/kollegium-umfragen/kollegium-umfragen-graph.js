/**
 * SharePoint-CRUD für Kollegiumsumfragen (Umfragen + Antworten, Live).
 * Transport: shared/graph-client.js (R22 / b11); Token/Site-Helfer weiter über spo-graph-shared.
 */
import { graphJson as sharedGraphJson } from '../../shared/graph-client.js';

import { LIST_TITLES } from './kollegium-umfragen-schema.js';
import { normalizeAntwort, normalizeUmfrage } from './kollegium-umfragen-model.js';
import {
    loadKollegiumSpSettings,
    resolveKollegiumSiteUrl
} from './kollegium-umfragen-settings.js';

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

function fieldBool(fields, key) {
    const v = fields && fields[key];
    if (v === true || v === 'true' || v === 1 || v === '1') return true;
    if (v === false || v === 'false' || v === 0 || v === '0') return false;
    return !!v;
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
export function mapUmfrageFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeUmfrage({
        id: fieldStr(f, 'UmfrageId') || (item && item.id != null ? 'spo-' + item.id : ''),
        titel: fieldStr(f, 'Title'),
        beschreibung: fieldStr(f, 'Beschreibung'),
        vorlageId: fieldStr(f, 'VorlageId'),
        komplexitaet: fieldStr(f, 'Komplexitaet'),
        zielgruppe: fieldStr(f, 'Zielgruppe'),
        abteilung: fieldStr(f, 'Abteilung'),
        klasseCodes: fieldStr(f, 'KlasseCodes'),
        empfaengerMails: fieldStr(f, 'EmpfaengerMails'),
        fragen: fieldStr(f, 'FragenJson') || '[]',
        faelligAm: fieldStr(f, 'FaelligAm'),
        anonym: fieldBool(f, 'Anonym'),
        anhangErlaubt: fieldBool(f, 'AnhangErlaubt'),
        status: fieldStr(f, 'Status'),
        erstellerEmail: fieldStr(f, 'ErstellerEmail'),
        erstellerName: fieldStr(f, 'ErstellerName'),
        quelle: fieldStr(f, 'Quelle') || 'ui',
        seedTag: fieldStr(f, 'SeedTag'),
        spoItemId: item && item.id != null ? String(item.id) : '',
        createdAt: (item && item.createdDateTime) || ''
    });
}

/**
 * @param {object} umfrage
 * @param {{ omitOptional?: boolean }} [opts]
 */
export function mapUmfrageToFields(umfrage, opts = {}) {
    const u = normalizeUmfrage(umfrage);
    /** @type {Record<string, string|boolean>} */
    const fields = {
        Title: u.titel || 'Umfrage',
        UmfrageId: u.id,
        Beschreibung: u.beschreibung,
        VorlageId: u.vorlageId,
        Komplexitaet: u.komplexitaet,
        Zielgruppe: u.zielgruppe,
        Abteilung: u.abteilung,
        KlasseCodes: u.klasseCodes,
        EmpfaengerMails: u.empfaengerMails,
        FragenJson: JSON.stringify(u.fragen || []),
        Anonym: !!u.anonym,
        AnhangErlaubt: !!u.anhangErlaubt,
        Status: u.status,
        ErstellerEmail: u.erstellerEmail,
        ErstellerName: u.erstellerName,
        Quelle: u.quelle || 'ui'
    };
    if (u.faelligAm) fields.FaelligAm = u.faelligAm;
    if (!opts.omitOptional) {
        if (u.seedTag) fields.SeedTag = u.seedTag;
    }
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = u.titel || 'Umfrage';
    if (u.id) fields.UmfrageId = u.id;
    fields.Anonym = !!u.anonym;
    fields.AnhangErlaubt = !!u.anhangErlaubt;
    return fields;
}

/**
 * @param {object} item Graph list item
 */
export function mapAntwortFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeAntwort({
        id: fieldStr(f, 'AntwortId') || (item && item.id != null ? 'spo-a-' + item.id : ''),
        umfrageId: fieldStr(f, 'UmfrageId'),
        lehrerEmail: fieldStr(f, 'LehrerEmail'),
        lehrerName: fieldStr(f, 'LehrerName'),
        werte: fieldStr(f, 'AntwortenJson') || '{}',
        status: fieldStr(f, 'Status'),
        abgegebenAm: fieldStr(f, 'AbgegebenAm'),
        anhangName: fieldStr(f, 'AnhangName'),
        anhangMime: fieldStr(f, 'AnhangMime'),
        anhangUrl: fieldStr(f, 'AnhangUrl'),
        seedTag: fieldStr(f, 'SeedTag'),
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/**
 * @param {object} antwort
 * @param {{ omitOptional?: boolean }} [opts]
 */
export function mapAntwortToFields(antwort, opts = {}) {
    const a = normalizeAntwort(antwort);
    /** @type {Record<string, string>} */
    const fields = {
        Title: a.lehrerName || a.lehrerEmail || a.id || 'Antwort',
        AntwortId: a.id,
        UmfrageId: a.umfrageId,
        LehrerEmail: a.lehrerEmail,
        LehrerName: a.lehrerName,
        AntwortenJson: JSON.stringify(a.werte || {}),
        Status: a.status,
        AnhangName: a.anhangName,
        AnhangMime: a.anhangMime,
        AnhangUrl: a.anhangUrl
    };
    if (a.abgegebenAm) fields.AbgegebenAm = a.abgegebenAm;
    if (!opts.omitOptional) {
        if (a.seedTag) fields.SeedTag = a.seedTag;
    }
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = a.lehrerName || a.lehrerEmail || a.id || 'Antwort';
    if (a.id) fields.AntwortId = a.id;
    if (a.umfrageId) fields.UmfrageId = a.umfrageId;
    return fields;
}

/**
 * @param {string} [webUrl]
 */
export async function resolveKollegiumContext(webUrl) {
    const settings = loadKollegiumSpSettings();
    const url = String(webUrl || settings.siteUrl || resolveKollegiumSiteUrl() || '').trim();
    if (!url) throw new Error('Keine SharePoint-Site hinterlegt. Bitte Listen-Assistent oder System nutzen.');
    const listUmfragen = settings.listUmfragen || LIST_TITLES.umfragen;
    const listAntworten = settings.listAntworten || LIST_TITLES.antworten;
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const umfragenList = await findListByTitle(tok, siteId, listUmfragen);
    if (!umfragenList) {
        throw new Error(
            'Liste „' + listUmfragen + '“ fehlt. Bitte zuerst unter „Listen einrichten“ anlegen.'
        );
    }
    const antwortenList = await findListByTitle(tok, siteId, listAntworten);
    if (!antwortenList) {
        throw new Error(
            'Liste „' + listAntworten + '“ fehlt. Bitte zuerst unter „Listen einrichten“ anlegen.'
        );
    }
    return {
        webUrl: url,
        siteId,
        siteName: site.displayName || '',
        listUmfragen: {
            id: String(umfragenList.id),
            webUrl: umfragenList.webUrl || '',
            title: umfragenList.displayName || listUmfragen
        },
        listAntworten: {
            id: String(antwortenList.id),
            webUrl: antwortenList.webUrl || '',
            title: antwortenList.displayName || listAntworten
        }
    };
}

/**
 * @param {{ webUrl?: string }} [opts]
 */
export async function loadAllKollegium(opts = {}) {
    const ctx = await resolveKollegiumContext(opts.webUrl);
    const tok = await token();
    const umfrageRows = await fetchAllItems(tok, ctx.siteId, ctx.listUmfragen.id);
    const antwortRows = await fetchAllItems(tok, ctx.siteId, ctx.listAntworten.id);
    return {
        ctx,
        umfragen: umfrageRows.map(mapUmfrageFromItem),
        antworten: antwortRows.map(mapAntwortFromItem)
    };
}

/**
 * @param {object} ctx
 * @param {object} umfrage
 */
export async function createUmfrageItem(ctx, umfrage) {
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listUmfragen.id) +
        '/items';
    let fields = mapUmfrageToFields(umfrage, { omitOptional: false });
    let created;
    try {
        created = await graphJson('POST', path, tok, { fields }, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/SeedTag|UmfrageId|FragenJson/i.test(msg)) {
            fields = mapUmfrageToFields(umfrage, { omitOptional: true });
            created = await graphJson('POST', path, tok, { fields }, 'v1.0');
        } else {
            throw e;
        }
    }
    const mapped = mapUmfrageFromItem(created);
    const local = normalizeUmfrage(umfrage);
    return {
        ...local,
        id: local.id || mapped.id,
        spoItemId: mapped.spoItemId || local.spoItemId,
        // Graph-Create liefert Felder oft unvollständig – Domäne lokal behalten
        fragen: local.fragen.length ? local.fragen : mapped.fragen
    };
}

/**
 * @param {object} ctx
 * @param {object} umfrage
 */
export async function updateUmfrageItem(ctx, umfrage) {
    const spoId = String(umfrage.spoItemId || '').trim();
    if (!spoId) throw new Error('spoItemId fehlt – Umfrage noch nicht in SharePoint.');
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listUmfragen.id) +
        '/items/' +
        encodeURIComponent(spoId) +
        '/fields';
    let fields = mapUmfrageToFields(umfrage, { omitOptional: false });
    try {
        await graphJson('PATCH', path, tok, fields, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/SeedTag|UmfrageId|FragenJson/i.test(msg)) {
            fields = mapUmfrageToFields(umfrage, { omitOptional: true });
            await graphJson('PATCH', path, tok, fields, 'v1.0');
        } else {
            throw e;
        }
    }
    return { ...normalizeUmfrage(umfrage), spoItemId: spoId };
}

/**
 * @param {object} ctx
 * @param {string} spoItemId
 */
export async function deleteUmfrageItem(ctx, spoItemId) {
    const spoId = String(spoItemId || '').trim();
    if (!spoId) return;
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listUmfragen.id) +
        '/items/' +
        encodeURIComponent(spoId);
    await graphJson('DELETE', path, tok, undefined, 'v1.0');
}

/**
 * @param {object} ctx
 * @param {object} antwort
 */
export async function createAntwortItem(ctx, antwort) {
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listAntworten.id) +
        '/items';
    let fields = mapAntwortToFields(antwort, { omitOptional: false });
    let created;
    try {
        created = await graphJson('POST', path, tok, { fields }, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/SeedTag|AntwortId|AntwortenJson/i.test(msg)) {
            fields = mapAntwortToFields(antwort, { omitOptional: true });
            created = await graphJson('POST', path, tok, { fields }, 'v1.0');
        } else {
            throw e;
        }
    }
    const mapped = mapAntwortFromItem(created);
    const local = normalizeAntwort(antwort);
    return {
        ...local,
        id: local.id || mapped.id,
        spoItemId: mapped.spoItemId || local.spoItemId,
        werte:
            local.werte && Object.keys(local.werte).length
                ? local.werte
                : mapped.werte
    };
}

/**
 * @param {object} ctx
 * @param {object} antwort
 */
export async function updateAntwortItem(ctx, antwort) {
    const spoId = String(antwort.spoItemId || '').trim();
    if (!spoId) throw new Error('spoItemId fehlt – Antwort noch nicht in SharePoint.');
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listAntworten.id) +
        '/items/' +
        encodeURIComponent(spoId) +
        '/fields';
    let fields = mapAntwortToFields(antwort, { omitOptional: false });
    try {
        await graphJson('PATCH', path, tok, fields, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/SeedTag|AntwortId|AntwortenJson/i.test(msg)) {
            fields = mapAntwortToFields(antwort, { omitOptional: true });
            await graphJson('PATCH', path, tok, fields, 'v1.0');
        } else {
            throw e;
        }
    }
    return { ...normalizeAntwort(antwort), spoItemId: spoId };
}

/**
 * @param {object} ctx
 * @param {string} spoItemId
 */
export async function deleteAntwortItem(ctx, spoItemId) {
    const spoId = String(spoItemId || '').trim();
    if (!spoId) return;
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listAntworten.id) +
        '/items/' +
        encodeURIComponent(spoId);
    await graphJson('DELETE', path, tok, undefined, 'v1.0');
}

export { SCOPES as KOLLEGIUM_GRAPH_SCOPES };
