/**
 * SharePoint-CRUD für Ausgabenverwaltung (Inventar + Vorgänge).
 */
import { LIST_TITLES } from './ausgaben-schema.js';
import { normalizeGegenstand, normalizeVorgang } from './ausgaben-model.js';
import { loadAusgabenSpSettings, resolveAusgabenSiteUrl } from './ausgaben-settings.js';

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

function fieldBool(fields, key, fallbackTrue) {
    const v = fields && fields[key];
    if (v === true || v === 'true' || v === 1 || v === '1') return true;
    if (v === false || v === 'false' || v === 0 || v === '0') return false;
    return fallbackTrue !== false;
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

/**
 * @param {object} item Graph list item
 */
export function mapGegenstandFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeGegenstand({
        id: fieldStr(f, 'GegenstandId') || (item && item.id != null ? 'spo-gg-' + item.id : ''),
        typ: fieldStr(f, 'Typ') || 'sonstiges',
        typLabel: fieldStr(f, 'TypLabel'),
        kennung: fieldStr(f, 'Kennung') || fieldStr(f, 'Title'),
        bezeichnung: fieldStr(f, 'Bezeichnung') || fieldStr(f, 'Title'),
        ortZone: fieldStr(f, 'OrtZone'),
        aktiv: fieldBool(f, 'Aktiv', true),
        notiz: fieldStr(f, 'Notiz'),
        quelle: fieldStr(f, 'Quelle') || 'ui',
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/**
 * @param {object} gegenstand
 * @param {{ omitOptional?: boolean }} [opts]
 */
export function mapGegenstandToFields(gegenstand, opts = {}) {
    const g = normalizeGegenstand(gegenstand);
    /** @type {Record<string, string|boolean>} */
    const fields = {
        Title: g.bezeichnung || g.kennung || 'Gegenstand',
        GegenstandId: g.id,
        Typ: g.typ || 'sonstiges',
        Kennung: g.kennung,
        Bezeichnung: g.bezeichnung,
        OrtZone: g.ortZone,
        Aktiv: !!g.aktiv,
        Notiz: g.notiz
    };
    if (!opts.omitOptional) {
        if (g.typLabel) fields.TypLabel = g.typLabel;
        if (g.quelle) fields.Quelle = g.quelle;
    }
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = g.bezeichnung || g.kennung || 'Gegenstand';
    if (g.id) fields.GegenstandId = g.id;
    fields.Aktiv = !!g.aktiv;
    return fields;
}

/**
 * @param {object} item Graph list item
 */
export function mapVorgangFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeVorgang({
        id: fieldStr(f, 'VorgangId') || (item && item.id != null ? 'spo-vg-' + item.id : ''),
        gegenstandId: fieldStr(f, 'GegenstandId'),
        kennung: fieldStr(f, 'Kennung'),
        empfaengerTyp: fieldStr(f, 'EmpfaengerTyp') || 'lehrer',
        empfaengerName: fieldStr(f, 'EmpfaengerName') || fieldStr(f, 'Title'),
        empfaengerEmail: fieldStr(f, 'EmpfaengerEmail'),
        klasseCode: fieldStr(f, 'KlasseCode'),
        ausgegebenAm: fieldStr(f, 'AusgegebenAm'),
        faelligAm: fieldStr(f, 'FaelligAm'),
        zurueckAm: fieldStr(f, 'ZurueckAm'),
        status: fieldStr(f, 'Status') || 'ausgegeben',
        kautionHinweis: fieldStr(f, 'KautionHinweis'),
        unterschriftEmpfang: fieldBool(f, 'UnterschriftEmpfang', false),
        notiz: fieldStr(f, 'Notiz'),
        ausgegebenVon: fieldStr(f, 'AusgegebenVon'),
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/**
 * @param {object} vorgang
 * @param {{ omitOptional?: boolean }} [opts]
 */
export function mapVorgangToFields(vorgang, opts = {}) {
    const v = normalizeVorgang(vorgang);
    /** @type {Record<string, string|boolean>} */
    const fields = {
        Title: (v.kennung ? v.kennung + ' → ' : '') + (v.empfaengerName || 'Vorgang'),
        VorgangId: v.id,
        GegenstandId: v.gegenstandId,
        Kennung: v.kennung,
        EmpfaengerTyp: v.empfaengerTyp,
        EmpfaengerName: v.empfaengerName,
        EmpfaengerEmail: v.empfaengerEmail,
        Status: v.status,
        Notiz: v.notiz,
        AusgegebenVon: v.ausgegebenVon,
        UnterschriftEmpfang: !!v.unterschriftEmpfang
    };
    if (v.ausgegebenAm) fields.AusgegebenAm = v.ausgegebenAm;
    if (v.faelligAm) fields.FaelligAm = v.faelligAm;
    if (v.zurueckAm) fields.ZurueckAm = v.zurueckAm;
    if (!opts.omitOptional) {
        if (v.klasseCode) fields.KlasseCode = v.klasseCode;
        if (v.kautionHinweis) fields.KautionHinweis = v.kautionHinweis;
    }
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = (v.kennung ? v.kennung + ' → ' : '') + (v.empfaengerName || 'Vorgang');
    if (v.id) fields.VorgangId = v.id;
    fields.UnterschriftEmpfang = !!v.unterschriftEmpfang;
    return fields;
}

/**
 * @param {string} [webUrl]
 */
export async function resolveAusgabenContext(webUrl) {
    const settings = loadAusgabenSpSettings();
    const url = String(webUrl || settings.siteUrl || resolveAusgabenSiteUrl() || '').trim();
    if (!url) throw new Error('Keine SharePoint-Site hinterlegt. Bitte Listen-Assistent oder System nutzen.');
    const listGgTitle = settings.listGegenstaende || LIST_TITLES.gegenstaende;
    const listVgTitle = settings.listVorgaenge || LIST_TITLES.vorgaenge;
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const listGg = await findListByTitle(tok, siteId, listGgTitle);
    if (!listGg) {
        throw new Error(
            'Liste „' + listGgTitle + '“ fehlt. Bitte zuerst unter „Listen einrichten“ anlegen.'
        );
    }
    const listVg = await findListByTitle(tok, siteId, listVgTitle);
    if (!listVg) {
        throw new Error(
            'Liste „' + listVgTitle + '“ fehlt. Bitte zuerst unter „Listen einrichten“ anlegen.'
        );
    }
    return {
        webUrl: url,
        siteId,
        siteName: site.displayName || '',
        lists: {
            gegenstaende: {
                id: String(listGg.id),
                webUrl: listGg.webUrl || '',
                title: listGg.displayName || listGgTitle
            },
            vorgaenge: {
                id: String(listVg.id),
                webUrl: listVg.webUrl || '',
                title: listVg.displayName || listVgTitle
            }
        }
    };
}

/**
 * @param {{ webUrl?: string }} [opts]
 */
export async function loadAllAusgaben(opts = {}) {
    const ctx = await resolveAusgabenContext(opts.webUrl);
    const tok = await token();
    const ggRows = await fetchAllItems(tok, ctx.siteId, ctx.lists.gegenstaende.id);
    const vgRows = await fetchAllItems(tok, ctx.siteId, ctx.lists.vorgaenge.id);
    return {
        ctx,
        gegenstaende: ggRows.map(mapGegenstandFromItem),
        vorgaenge: vgRows.map(mapVorgangFromItem)
    };
}

/**
 * @param {object} ctx
 * @param {object} gegenstand
 */
export async function createGegenstandItem(ctx, gegenstand) {
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.lists.gegenstaende.id) +
        '/items';
    let fields = mapGegenstandToFields(gegenstand, { omitOptional: false });
    let created;
    try {
        created = await G().graphJson('POST', path, tok, { fields }, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/TypLabel|Quelle|GegenstandId/i.test(msg)) {
            fields = mapGegenstandToFields(gegenstand, { omitOptional: true });
            created = await G().graphJson('POST', path, tok, { fields }, 'v1.0');
        } else {
            throw e;
        }
    }
    const mapped = mapGegenstandFromItem(created);
    const local = normalizeGegenstand(gegenstand);
    return {
        ...local,
        ...mapped,
        id: local.id || mapped.id,
        spoItemId: mapped.spoItemId || local.spoItemId
    };
}

/**
 * @param {object} ctx
 * @param {object} gegenstand
 */
export async function updateGegenstandItem(ctx, gegenstand) {
    const spoId = String(gegenstand.spoItemId || '').trim();
    if (!spoId) throw new Error('spoItemId fehlt – Gegenstand noch nicht in SharePoint.');
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.lists.gegenstaende.id) +
        '/items/' +
        encodeURIComponent(spoId) +
        '/fields';
    let fields = mapGegenstandToFields(gegenstand, { omitOptional: false });
    try {
        await G().graphJson('PATCH', path, tok, fields, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/TypLabel|Quelle|GegenstandId/i.test(msg)) {
            fields = mapGegenstandToFields(gegenstand, { omitOptional: true });
            await G().graphJson('PATCH', path, tok, fields, 'v1.0');
        } else {
            throw e;
        }
    }
    return { ...normalizeGegenstand(gegenstand), spoItemId: spoId };
}

/**
 * @param {object} ctx
 * @param {string} spoItemId
 */
export async function deleteGegenstandItem(ctx, spoItemId) {
    const spoId = String(spoItemId || '').trim();
    if (!spoId) return;
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.lists.gegenstaende.id) +
        '/items/' +
        encodeURIComponent(spoId);
    await G().graphJson('DELETE', path, tok, undefined, 'v1.0');
}

/**
 * @param {object} ctx
 * @param {object} vorgang
 */
export async function createVorgangItem(ctx, vorgang) {
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.lists.vorgaenge.id) +
        '/items';
    let fields = mapVorgangToFields(vorgang, { omitOptional: false });
    let created;
    try {
        created = await G().graphJson('POST', path, tok, { fields }, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/KlasseCode|KautionHinweis|VorgangId/i.test(msg)) {
            fields = mapVorgangToFields(vorgang, { omitOptional: true });
            created = await G().graphJson('POST', path, tok, { fields }, 'v1.0');
        } else {
            throw e;
        }
    }
    const mapped = mapVorgangFromItem(created);
    const local = normalizeVorgang(vorgang);
    return {
        ...local,
        ...mapped,
        id: local.id || mapped.id,
        spoItemId: mapped.spoItemId || local.spoItemId
    };
}

/**
 * @param {object} ctx
 * @param {object} vorgang
 */
export async function updateVorgangItem(ctx, vorgang) {
    const spoId = String(vorgang.spoItemId || '').trim();
    if (!spoId) throw new Error('spoItemId fehlt – Vorgang noch nicht in SharePoint.');
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.lists.vorgaenge.id) +
        '/items/' +
        encodeURIComponent(spoId) +
        '/fields';
    let fields = mapVorgangToFields(vorgang, { omitOptional: false });
    try {
        await G().graphJson('PATCH', path, tok, fields, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/KlasseCode|KautionHinweis|VorgangId/i.test(msg)) {
            fields = mapVorgangToFields(vorgang, { omitOptional: true });
            await G().graphJson('PATCH', path, tok, fields, 'v1.0');
        } else {
            throw e;
        }
    }
    return { ...normalizeVorgang(vorgang), spoItemId: spoId };
}

/**
 * @param {object} ctx
 * @param {string} spoItemId
 */
export async function deleteVorgangItem(ctx, spoItemId) {
    const spoId = String(spoItemId || '').trim();
    if (!spoId) return;
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.lists.vorgaenge.id) +
        '/items/' +
        encodeURIComponent(spoId);
    await G().graphJson('DELETE', path, tok, undefined, 'v1.0');
}

export { SCOPES as AUSGABEN_GRAPH_SCOPES };
