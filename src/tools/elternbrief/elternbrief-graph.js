/**
 * SharePoint-CRUD für Elternbrief (Briefe + Lesebestätigungen).
 */
import { LIST_TITLES } from './elternbrief-schema.js';
import { normalizeBrief, normalizeEmpfaenger } from './elternbrief-model.js';
import {
    loadElternbriefSpSettings,
    resolveElternbriefSiteUrl
} from './elternbrief-settings.js';

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

function fieldBool(fields, key, fallback = true) {
    const v = fields && fields[key];
    if (v === false || v === 0 || v === '0' || v === 'false' || v === 'False') return false;
    if (v === true || v === 1 || v === '1' || v === 'true' || v === 'True') return true;
    return fallback;
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
export function mapBriefFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeBrief({
        id: fieldStr(f, 'BriefId') || (item && item.id != null ? 'spo-' + item.id : ''),
        betreff: fieldStr(f, 'Betreff') || fieldStr(f, 'Title'),
        inhalt: fieldStr(f, 'InhaltHtml'),
        vorlageId: fieldStr(f, 'VorlageId'),
        klasseCodes: fieldStr(f, 'KlasseCodes'),
        absenderModus: fieldStr(f, 'AbsenderModus'),
        absenderEmail: fieldStr(f, 'AbsenderEmail'),
        absenderName: fieldStr(f, 'AbsenderName'),
        anhangName: fieldStr(f, 'AnhangName'),
        anhangMime: fieldStr(f, 'AnhangMime'),
        anhangUrl: fieldStr(f, 'AnhangUrl'),
        faelligAm: fieldStr(f, 'FaelligAm'),
        lesebestaetigung: fieldBool(f, 'Lesebestaetigung', true),
        status: fieldStr(f, 'Status'),
        erstellerEmail: fieldStr(f, 'ErstellerEmail'),
        erstellerName: fieldStr(f, 'ErstellerName'),
        gesendetAm: fieldStr(f, 'GesendetAm'),
        seedTag: fieldStr(f, 'SeedTag'),
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/**
 * @param {object} brief
 * @param {{ omitOptional?: boolean }} [opts]
 */
export function mapBriefToFields(brief, opts = {}) {
    const b = normalizeBrief(brief);
    /** @type {Record<string, string|boolean>} */
    const fields = {
        Title: b.betreff || 'Elternbrief',
        BriefId: b.id,
        Betreff: b.betreff || 'Elternbrief',
        InhaltHtml: b.inhalt,
        VorlageId: b.vorlageId,
        KlasseCodes: b.klasseCodes,
        AbsenderModus: b.absenderModus,
        AbsenderEmail: b.absenderEmail,
        AbsenderName: b.absenderName,
        AnhangName: b.anhangName,
        AnhangMime: b.anhangMime,
        AnhangUrl: b.anhangUrl,
        Status: b.status,
        ErstellerEmail: b.erstellerEmail,
        ErstellerName: b.erstellerName,
        Lesebestaetigung: b.lesebestaetigung !== false
    };
    if (b.faelligAm) fields.FaelligAm = b.faelligAm;
    if (b.gesendetAm) fields.GesendetAm = b.gesendetAm;
    if (!opts.omitOptional) {
        if (b.seedTag) fields.SeedTag = b.seedTag;
    }
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = b.betreff || 'Elternbrief';
    if (b.id) fields.BriefId = b.id;
    fields.Lesebestaetigung = b.lesebestaetigung !== false;
    return fields;
}

/**
 * @param {object} item Graph list item
 */
export function mapEmpfaengerFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeEmpfaenger({
        id: fieldStr(f, 'LeseId') || (item && item.id != null ? 'spo-el-' + item.id : ''),
        briefId: fieldStr(f, 'BriefId'),
        schuelerName: fieldStr(f, 'SchuelerName') || fieldStr(f, 'Title'),
        klasseCode: fieldStr(f, 'KlasseCode'),
        elternEmail: fieldStr(f, 'ElternEmail'),
        status: fieldStr(f, 'Status'),
        gelesenAm: fieldStr(f, 'GelesenAm'),
        seedTag: fieldStr(f, 'SeedTag'),
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/**
 * @param {object} emp
 * @param {{ omitOptional?: boolean }} [opts]
 */
export function mapEmpfaengerToFields(emp, opts = {}) {
    const e = normalizeEmpfaenger(emp);
    /** @type {Record<string, string>} */
    const fields = {
        Title: e.schuelerName || e.id || 'Empfänger',
        LeseId: e.id,
        BriefId: e.briefId,
        SchuelerName: e.schuelerName,
        KlasseCode: e.klasseCode,
        ElternEmail: e.elternEmail,
        Status: e.status
    };
    if (e.gelesenAm) fields.GelesenAm = e.gelesenAm;
    if (!opts.omitOptional) {
        if (e.seedTag) fields.SeedTag = e.seedTag;
    }
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = e.schuelerName || e.id || 'Empfänger';
    if (e.id) fields.LeseId = e.id;
    if (e.briefId) fields.BriefId = e.briefId;
    return fields;
}

/**
 * @param {string} [webUrl]
 */
export async function resolveElternbriefContext(webUrl) {
    const settings = loadElternbriefSpSettings();
    const url = String(webUrl || settings.siteUrl || resolveElternbriefSiteUrl() || '').trim();
    if (!url) throw new Error('Keine SharePoint-Site hinterlegt. Bitte Listen-Assistent oder System nutzen.');
    const titleBriefe = settings.listBriefe || LIST_TITLES.briefe;
    const titleLese = settings.listLese || LIST_TITLES.lesebestaetigungen;
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const listBriefe = await findListByTitle(tok, siteId, titleBriefe);
    if (!listBriefe) {
        throw new Error(
            'Liste „' + titleBriefe + '“ fehlt. Bitte zuerst unter „Listen einrichten“ anlegen.'
        );
    }
    const listLese = await findListByTitle(tok, siteId, titleLese);
    if (!listLese) {
        throw new Error(
            'Liste „' + titleLese + '“ fehlt. Bitte zuerst unter „Listen einrichten“ anlegen.'
        );
    }
    return {
        webUrl: url,
        siteId,
        siteName: site.displayName || '',
        listBriefe: {
            id: String(listBriefe.id),
            webUrl: listBriefe.webUrl || '',
            title: listBriefe.displayName || titleBriefe
        },
        listLese: {
            id: String(listLese.id),
            webUrl: listLese.webUrl || '',
            title: listLese.displayName || titleLese
        }
    };
}

/**
 * @param {{ webUrl?: string }} [opts]
 */
export async function loadAllElternbrief(opts = {}) {
    const ctx = await resolveElternbriefContext(opts.webUrl);
    const tok = await token();
    const [briefRows, leseRows] = await Promise.all([
        fetchAllItems(tok, ctx.siteId, ctx.listBriefe.id),
        fetchAllItems(tok, ctx.siteId, ctx.listLese.id)
    ]);
    return {
        ctx,
        briefe: briefRows.map(mapBriefFromItem),
        empfaenger: leseRows.map(mapEmpfaengerFromItem)
    };
}

/**
 * @param {object} ctx
 * @param {object} brief
 */
export async function createBriefItem(ctx, brief) {
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listBriefe.id) +
        '/items';
    let fields = mapBriefToFields(brief, { omitOptional: false });
    let created;
    try {
        created = await G().graphJson('POST', path, tok, { fields }, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/SeedTag|BriefId|Betreff|InhaltHtml/i.test(msg)) {
            fields = mapBriefToFields(brief, { omitOptional: true });
            created = await G().graphJson('POST', path, tok, { fields }, 'v1.0');
        } else {
            throw e;
        }
    }
    const mapped = mapBriefFromItem(created);
    const local = normalizeBrief(brief);
    return {
        ...local,
        ...mapped,
        id: local.id || mapped.id,
        spoItemId: mapped.spoItemId || local.spoItemId
    };
}

/**
 * @param {object} ctx
 * @param {object} brief
 */
export async function updateBriefItem(ctx, brief) {
    const spoId = String(brief.spoItemId || '').trim();
    if (!spoId) throw new Error('spoItemId fehlt – Brief noch nicht in SharePoint.');
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listBriefe.id) +
        '/items/' +
        encodeURIComponent(spoId) +
        '/fields';
    let fields = mapBriefToFields(brief, { omitOptional: false });
    try {
        await G().graphJson('PATCH', path, tok, fields, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/SeedTag|BriefId|Betreff|InhaltHtml/i.test(msg)) {
            fields = mapBriefToFields(brief, { omitOptional: true });
            await G().graphJson('PATCH', path, tok, fields, 'v1.0');
        } else {
            throw e;
        }
    }
    return { ...normalizeBrief(brief), spoItemId: spoId };
}

/**
 * @param {object} ctx
 * @param {string} spoItemId
 */
export async function deleteBriefItem(ctx, spoItemId) {
    const spoId = String(spoItemId || '').trim();
    if (!spoId) return;
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listBriefe.id) +
        '/items/' +
        encodeURIComponent(spoId);
    await G().graphJson('DELETE', path, tok, undefined, 'v1.0');
}

/**
 * @param {object} ctx
 * @param {object} emp
 */
export async function createEmpfaengerItem(ctx, emp) {
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listLese.id) +
        '/items';
    let fields = mapEmpfaengerToFields(emp, { omitOptional: false });
    let created;
    try {
        created = await G().graphJson('POST', path, tok, { fields }, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/SeedTag|LeseId|SchuelerName/i.test(msg)) {
            fields = mapEmpfaengerToFields(emp, { omitOptional: true });
            created = await G().graphJson('POST', path, tok, { fields }, 'v1.0');
        } else {
            throw e;
        }
    }
    const mapped = mapEmpfaengerFromItem(created);
    const local = normalizeEmpfaenger(emp);
    return {
        ...local,
        ...mapped,
        id: local.id || mapped.id,
        spoItemId: mapped.spoItemId || local.spoItemId
    };
}

/**
 * @param {object} ctx
 * @param {object} emp
 */
export async function updateEmpfaengerItem(ctx, emp) {
    const spoId = String(emp.spoItemId || '').trim();
    if (!spoId) throw new Error('spoItemId fehlt – Empfänger noch nicht in SharePoint.');
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listLese.id) +
        '/items/' +
        encodeURIComponent(spoId) +
        '/fields';
    let fields = mapEmpfaengerToFields(emp, { omitOptional: false });
    try {
        await G().graphJson('PATCH', path, tok, fields, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/SeedTag|LeseId|SchuelerName/i.test(msg)) {
            fields = mapEmpfaengerToFields(emp, { omitOptional: true });
            await G().graphJson('PATCH', path, tok, fields, 'v1.0');
        } else {
            throw e;
        }
    }
    return { ...normalizeEmpfaenger(emp), spoItemId: spoId };
}

/**
 * @param {object} ctx
 * @param {string} spoItemId
 */
export async function deleteEmpfaengerItem(ctx, spoItemId) {
    const spoId = String(spoItemId || '').trim();
    if (!spoId) return;
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.listLese.id) +
        '/items/' +
        encodeURIComponent(spoId);
    await G().graphJson('DELETE', path, tok, undefined, 'v1.0');
}

export { SCOPES as ELTERNBRIEF_GRAPH_SCOPES };
