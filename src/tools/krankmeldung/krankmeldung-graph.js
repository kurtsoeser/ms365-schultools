/**
 * SharePoint-CRUD für Krankmeldungen (Live-Anbindung).
 */
import { LIST_TITLES } from './krankmeldung-schema.js';
import { normalizeMeldung } from './krankmeldung-model.js';
import { loadKrankmeldungSpSettings, resolveKrankmeldungSiteUrl } from './krankmeldung-settings.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

async function token() {
    return await G().getGraphToken(SCOPES);
}

function fieldStr(fields, key) {
    const v = fields && fields[key];
    return v == null ? '' : String(v).trim();
}

function fieldDateOnly(fields, key) {
    const s = fieldStr(fields, key);
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    return '';
}

async function findListByTitle(tok, siteId, listTitle) {
    const title = String(listTitle || '').trim();
    const path =
        G().graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await G().graphJson('GET', path, tok, undefined, 'v1.0');
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
        const data = await G().graphJson(
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

function buildTitle(m) {
    const name = String(m.schuelerName || '').trim() || 'Krankmeldung';
    const klasse = String(m.klasseCode || '').trim();
    return klasse ? name + ' (' + klasse + ')' : name;
}

/**
 * @param {object} item Graph list item
 */
export function mapMeldungFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeMeldung({
        id: fieldStr(f, 'MeldungId') || (item && item.id != null ? 'spo-' + item.id : ''),
        schuelerName: fieldStr(f, 'SchuelerName') || fieldStr(f, 'Title'),
        klasseCode: fieldStr(f, 'KlasseCode'),
        von: fieldDateOnly(f, 'Von'),
        bis: fieldDateOnly(f, 'Bis'),
        kategorie: fieldStr(f, 'Kategorie'),
        gemeldetAm: fieldStr(f, 'GemeldetAm') || (item && item.createdDateTime) || '',
        gemeldetVonEmail: fieldStr(f, 'GemeldetVonEmail'),
        gemeldetVonName: fieldStr(f, 'GemeldetVonName'),
        status: fieldStr(f, 'Status'),
        kenntnisLehrerEmail: fieldStr(f, 'KenntnisLehrerEmail'),
        notiz: fieldStr(f, 'Notiz'),
        anhangName: fieldStr(f, 'AnhangName'),
        anhangMime: fieldStr(f, 'AnhangMime'),
        anhangUrl: fieldStr(f, 'AnhangUrl'),
        quelle: fieldStr(f, 'Quelle') || 'public',
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
        Title: buildTitle(m),
        MeldungId: m.id,
        SchuelerName: m.schuelerName,
        KlasseCode: m.klasseCode,
        Kategorie: m.kategorie,
        Status: m.status,
        GemeldetVonEmail: m.gemeldetVonEmail,
        GemeldetVonName: m.gemeldetVonName,
        Notiz: m.notiz,
        AnhangName: m.anhangName,
        AnhangMime: m.anhangMime,
        AnhangUrl: m.anhangUrl
    };
    if (m.von) fields.Von = m.von;
    if (m.bis) fields.Bis = m.bis;
    if (m.gemeldetAm) fields.GemeldetAm = m.gemeldetAm;
    // Quelle nur bekannte Choice-Werte (kein „demo“)
    if (m.quelle === 'public' || m.quelle === 'ui' || m.quelle === 'import') {
        fields.Quelle = m.quelle;
    }
    if (!opts.omitOptional) {
        if (m.kenntnisLehrerEmail) fields.KenntnisLehrerEmail = m.kenntnisLehrerEmail;
        if (m.seedTag) fields.SeedTag = m.seedTag;
    }
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = buildTitle(m);
    if (m.id) fields.MeldungId = m.id;
    return fields;
}

/**
 * @param {string} [webUrl]
 */
export async function resolveKrankmeldungContext(webUrl) {
    const settings = loadKrankmeldungSpSettings();
    const url = String(webUrl || settings.siteUrl || resolveKrankmeldungSiteUrl() || '').trim();
    if (!url) {
        throw new Error('Keine SharePoint-Site hinterlegt. Bitte Listen-Assistent oder System nutzen.');
    }
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
    const ctx = await resolveKrankmeldungContext(opts.webUrl);
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
        created = await G().graphJson('POST', path, tok, { fields }, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/KenntnisLehrer|SeedTag|MeldungId|Quelle/i.test(msg)) {
            fields = mapMeldungToFields(meldung, { omitOptional: true });
            created = await G().graphJson('POST', path, tok, { fields }, 'v1.0');
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
        await G().graphJson('PATCH', path, tok, fields, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/KenntnisLehrer|SeedTag|MeldungId|Quelle/i.test(msg)) {
            fields = mapMeldungToFields(meldung, { omitOptional: true });
            await G().graphJson('PATCH', path, tok, fields, 'v1.0');
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
    await G().graphJson('DELETE', path, tok, undefined, 'v1.0');
}

export { SCOPES as KRANKMELDUNG_GRAPH_SCOPES };
