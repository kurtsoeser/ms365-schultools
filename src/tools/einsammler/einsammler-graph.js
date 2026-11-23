/**
 * SharePoint-CRUD für Einsammler (Köpfe + Antworten, Live-Anbindung).
 * Transport: shared/graph-client.js (R22 / b11); Token/Site-Helfer weiter über spo-graph-shared.
 */
import { graphJson as sharedGraphJson } from '../../shared/graph-client.js';

import { LIST_TITLES } from './einsammler-schema.js';
import { normalizeAntwort, normalizeEinsammler } from './einsammler-model.js';
import {
    loadEinsammlerSpSettings,
    resolveEinsammlerSiteUrl
} from './einsammler-settings.js';

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
    if (v === true || v === false) return v;
    if (v === 'true' || v === '1' || v === 1) return true;
    if (v === 'false' || v === '0' || v === 0) return false;
    return false;
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
export function mapKopfFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeEinsammler({
        id: fieldStr(f, 'EinsammlerId') || (item && item.id != null ? 'spo-' + item.id : ''),
        titel: fieldStr(f, 'Title'),
        modus: fieldStr(f, 'Modus'),
        typ: fieldStr(f, 'Typ'),
        textInhalt: fieldStr(f, 'TextInhalt'),
        klasseCodes: fieldStr(f, 'KlasseCodes'),
        faelligAm: fieldStr(f, 'FaelligAm'),
        status: fieldStr(f, 'Status'),
        uploadErlaubt: fieldBool(f, 'UploadErlaubt'),
        uploadPflicht: fieldBool(f, 'UploadPflicht'),
        erstellerEmail: fieldStr(f, 'ErstellerEmail'),
        erstellerName: fieldStr(f, 'ErstellerName'),
        quelle: fieldStr(f, 'Quelle') || 'ui',
        seedTag: fieldStr(f, 'SeedTag'),
        spoItemId: item && item.id != null ? String(item.id) : '',
        createdAt: (item && item.createdDateTime) || ''
    });
}

/**
 * @param {object} item Graph list item
 */
export function mapAntwortFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeAntwort({
        id: fieldStr(f, 'AntwortId') || (item && item.id != null ? 'spo-ant-' + item.id : ''),
        einsammlerId: fieldStr(f, 'EinsammlerId'),
        schuelerName: fieldStr(f, 'SchuelerName'),
        klasseCode: fieldStr(f, 'KlasseCode'),
        elternEmail: fieldStr(f, 'ElternEmail'),
        status: fieldStr(f, 'Status'),
        antwortAm: fieldStr(f, 'AntwortAm'),
        kommentar: fieldStr(f, 'Kommentar'),
        anhangName: fieldStr(f, 'AnhangName'),
        anhangMime: fieldStr(f, 'AnhangMime'),
        anhangUrl: fieldStr(f, 'AnhangUrl'),
        quelle: fieldStr(f, 'Quelle') || 'ui',
        seedTag: fieldStr(f, 'SeedTag'),
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/**
 * @param {object} kopf
 * @param {{ omitOptional?: boolean }} [opts]
 */
export function mapKopfToFields(kopf, opts = {}) {
    const m = normalizeEinsammler(kopf);
    const quelle = m.quelle === 'demo' ? 'ui' : m.quelle || 'ui';
    /** @type {Record<string, string|boolean>} */
    const fields = {
        Title: m.titel || 'Einsammler',
        EinsammlerId: m.id,
        Modus: m.modus,
        Typ: m.typ,
        TextInhalt: m.textInhalt,
        KlasseCodes: m.klasseCodes,
        Status: m.status,
        ErstellerEmail: m.erstellerEmail,
        ErstellerName: m.erstellerName,
        UploadErlaubt: !!m.uploadErlaubt,
        UploadPflicht: !!m.uploadPflicht,
        Quelle: ['ui', 'public', 'import'].includes(quelle) ? quelle : 'ui'
    };
    if (m.faelligAm) fields.FaelligAm = m.faelligAm;
    if (!opts.omitOptional && m.seedTag) fields.SeedTag = m.seedTag;
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = m.titel || 'Einsammler';
    if (m.id) fields.EinsammlerId = m.id;
    fields.UploadErlaubt = !!m.uploadErlaubt;
    fields.UploadPflicht = !!m.uploadPflicht;
    return fields;
}

/**
 * @param {object} antwort
 * @param {{ omitOptional?: boolean }} [opts]
 */
export function mapAntwortToFields(antwort, opts = {}) {
    const m = normalizeAntwort(antwort);
    const quelle = m.quelle === 'demo' ? 'ui' : m.quelle || 'ui';
    /** @type {Record<string, string>} */
    const fields = {
        Title: m.schuelerName || m.id || 'Antwort',
        AntwortId: m.id,
        EinsammlerId: m.einsammlerId,
        SchuelerName: m.schuelerName,
        KlasseCode: m.klasseCode,
        ElternEmail: m.elternEmail,
        Status: m.status,
        Kommentar: m.kommentar,
        AnhangName: m.anhangName,
        AnhangMime: m.anhangMime,
        AnhangUrl: m.anhangUrl,
        Quelle: ['ui', 'public', 'import'].includes(quelle) ? quelle : 'ui'
    };
    if (m.antwortAm) fields.AntwortAm = m.antwortAm;
    if (!opts.omitOptional && m.seedTag) fields.SeedTag = m.seedTag;
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = m.schuelerName || m.id || 'Antwort';
    if (m.id) fields.AntwortId = m.id;
    if (m.einsammlerId) fields.EinsammlerId = m.einsammlerId;
    return fields;
}

/**
 * @param {string} [webUrl]
 */
export async function resolveEinsammlerContext(webUrl) {
    const settings = loadEinsammlerSpSettings();
    const url = String(webUrl || settings.siteUrl || resolveEinsammlerSiteUrl() || '').trim();
    if (!url) {
        throw new Error('Keine SharePoint-Site hinterlegt. Bitte Listen-Assistent oder System nutzen.');
    }
    const titleKoepfe = settings.listKoepfe || LIST_TITLES.koepfe;
    const titleAntworten = settings.listAntworten || LIST_TITLES.antworten;
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const listKoepfe = await findListByTitle(tok, siteId, titleKoepfe);
    if (!listKoepfe) {
        throw new Error(
            'Liste „' + titleKoepfe + '“ fehlt. Bitte zuerst unter „Listen einrichten“ anlegen.'
        );
    }
    const listAntworten = await findListByTitle(tok, siteId, titleAntworten);
    if (!listAntworten) {
        throw new Error(
            'Liste „' + titleAntworten + '“ fehlt. Bitte zuerst unter „Listen einrichten“ anlegen.'
        );
    }
    return {
        webUrl: url,
        siteId,
        siteName: site.displayName || '',
        listKoepfe: {
            id: String(listKoepfe.id),
            webUrl: listKoepfe.webUrl || '',
            title: listKoepfe.displayName || titleKoepfe
        },
        listAntworten: {
            id: String(listAntworten.id),
            webUrl: listAntworten.webUrl || '',
            title: listAntworten.displayName || titleAntworten
        }
    };
}

/**
 * @param {{ webUrl?: string }} [opts]
 */
export async function loadAllEinsammler(opts = {}) {
    const ctx = await resolveEinsammlerContext(opts.webUrl);
    const tok = await token();
    const [kopfRows, antRows] = await Promise.all([
        fetchAllItems(tok, ctx.siteId, ctx.listKoepfe.id),
        fetchAllItems(tok, ctx.siteId, ctx.listAntworten.id)
    ]);
    return {
        ctx,
        einsammler: kopfRows.map(mapKopfFromItem),
        antworten: antRows.map(mapAntwortFromItem)
    };
}

/**
 * @param {object} ctx
 * @param {object} kopf
 */
export async function createKopfItem(ctx, kopf) {
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listKoepfe.id) +
        '/items';
    let fields = mapKopfToFields(kopf, { omitOptional: false });
    let created;
    try {
        created = await graphJson('POST', path, tok, { fields }, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/SeedTag|UploadErlaubt|UploadPflicht|EinsammlerId/i.test(msg)) {
            fields = mapKopfToFields(kopf, { omitOptional: true });
            created = await graphJson('POST', path, tok, { fields }, 'v1.0');
        } else {
            throw e;
        }
    }
    const mapped = mapKopfFromItem(created);
    const local = normalizeEinsammler(kopf);
    return {
        ...local,
        ...mapped,
        id: local.id || mapped.id,
        spoItemId: mapped.spoItemId || local.spoItemId
    };
}

/**
 * @param {object} ctx
 * @param {object} kopf
 */
export async function updateKopfItem(ctx, kopf) {
    const spoId = String(kopf.spoItemId || '').trim();
    if (!spoId) throw new Error('spoItemId fehlt – Einsammler noch nicht in SharePoint.');
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listKoepfe.id) +
        '/items/' +
        encodeURIComponent(spoId) +
        '/fields';
    let fields = mapKopfToFields(kopf, { omitOptional: false });
    try {
        await graphJson('PATCH', path, tok, fields, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/SeedTag|UploadErlaubt|UploadPflicht|EinsammlerId/i.test(msg)) {
            fields = mapKopfToFields(kopf, { omitOptional: true });
            await graphJson('PATCH', path, tok, fields, 'v1.0');
        } else {
            throw e;
        }
    }
    return { ...normalizeEinsammler(kopf), spoItemId: spoId };
}

/**
 * @param {object} ctx
 * @param {string} spoItemId
 */
export async function deleteKopfItem(ctx, spoItemId) {
    const spoId = String(spoItemId || '').trim();
    if (!spoId) return;
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listKoepfe.id) +
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
        if (/SeedTag|AnhangUrl|AnhangMime|AntwortId/i.test(msg)) {
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
        ...mapped,
        id: local.id || mapped.id,
        spoItemId: mapped.spoItemId || local.spoItemId
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
        if (/SeedTag|AnhangUrl|AnhangMime|AntwortId/i.test(msg)) {
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

export { SCOPES as EINSAMMLER_GRAPH_SCOPES };
