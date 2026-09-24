/**
 * SharePoint-CRUD für Schularbeiten-Planer (Graph).
 */
import { LIST_TITLES, newEntityId } from './schularbeiten-planer-schema.js';
import { toIsoDateOnly } from './schularbeiten-planer-logic.js';

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
 * @param {string} webUrl
 */
export async function resolvePlanerContext(webUrl) {
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die SharePoint-Website-URL eintragen.');
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    const regelwerk = await findListByTitle(tok, siteId, LIST_TITLES.regelwerk);
    const terminfenster = await findListByTitle(tok, siteId, LIST_TITLES.terminfenster);
    const schularbeiten = await findListByTitle(tok, siteId, LIST_TITLES.schularbeiten);
    const fachMeta = await findListByTitle(tok, siteId, LIST_TITLES.fachMeta);
    if (!regelwerk || !terminfenster || !schularbeiten) {
        throw new Error(
            'Listen unvollständig. Bitte zuerst „Schularbeiten-Listen“ anlegen (Regelwerk, Terminfenster, Schularbeiten).'
        );
    }
    return {
        webUrl: url,
        siteId,
        siteName: site.displayName || '',
        lists: {
            regelwerk: { id: String(regelwerk.id), webUrl: regelwerk.webUrl || '' },
            terminfenster: { id: String(terminfenster.id), webUrl: terminfenster.webUrl || '' },
            schularbeiten: { id: String(schularbeiten.id), webUrl: schularbeiten.webUrl || '' },
            fachMeta: fachMeta
                ? { id: String(fachMeta.id), webUrl: fachMeta.webUrl || '' }
                : null
        }
    };
}

function fieldStr(fields, key) {
    const v = fields && fields[key];
    return v == null ? '' : String(v).trim();
}

function fieldNum(fields, key, fallback) {
    const n = Number(fields && fields[key]);
    return Number.isFinite(n) ? n : fallback;
}

/**
 * @param {object} item Graph list item
 */
export function mapSchularbeitFromItem(item) {
    const f = (item && item.fields) || {};
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        schularbeitId: fieldStr(f, 'SchularbeitId'),
        thema: fieldStr(f, 'Title'),
        fachCode: fieldStr(f, 'FachCode'),
        klasseCode: fieldStr(f, 'KlasseCode'),
        lehrerCode: fieldStr(f, 'LehrerCode'),
        lehrerEmail: fieldStr(f, 'LehrerEmail').toLowerCase(),
        datum: toIsoDateOnly(f.Datum) || '',
        dauerMinuten: fieldNum(f, 'DauerMinuten', 100),
        semester: fieldStr(f, 'Semester') || 'WS',
        status: fieldStr(f, 'Status') || 'beantragt',
        notiz: fieldStr(f, 'Notiz'),
        ablehnungsGrund: fieldStr(f, 'AblehnungsGrund'),
        beantragtVon: fieldStr(f, 'BeantragtVon'),
        fixiertVon: fieldStr(f, 'FixiertVon'),
        fixiertAm: fieldStr(f, 'FixiertAm'),
        schulterminKey: fieldStr(f, 'SchulterminKey')
    };
}

/**
 * @param {object} sa
 * @param {{ includeId?: boolean }} [opts]
 */
export function mapSchularbeitToFields(sa, opts) {
    const includeId = !opts || opts.includeId !== false;
    const fields = {
        Title: String(sa.thema || '').trim() || 'Schularbeit',
        FachCode: String(sa.fachCode || '').trim(),
        KlasseCode: String(sa.klasseCode || '').trim(),
        LehrerCode: String(sa.lehrerCode || '').trim(),
        LehrerEmail: String(sa.lehrerEmail || '').trim().toLowerCase(),
        Datum: toIsoDateOnly(sa.datum),
        DauerMinuten: Number(sa.dauerMinuten) || 100,
        Semester: String(sa.semester || 'WS').toUpperCase() === 'SS' ? 'SS' : 'WS',
        Status: String(sa.status || 'beantragt').toLowerCase(),
        Notiz: String(sa.notiz || ''),
        AblehnungsGrund: String(sa.ablehnungsGrund || ''),
        BeantragtVon: String(sa.beantragtVon || ''),
        FixiertVon: String(sa.fixiertVon || '')
    };
    if (sa.fixiertAm) {
        fields.FixiertAm = sa.fixiertAm;
    }
    if (sa.schulterminKey) {
        fields.SchulterminKey = String(sa.schulterminKey);
    }
    if (includeId) {
        fields.SchularbeitId = String(sa.schularbeitId || '').trim() || newEntityId('sa');
    }
    return fields;
}

export function mapFachMetaFromItem(item) {
    const f = (item && item.fields) || {};
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        fachCode: fieldStr(f, 'FachCode'),
        name: fieldStr(f, 'Title') || fieldStr(f, 'FachCode'),
        farbe: fieldStr(f, 'Farbe'),
        hatSchularbeiten: f.HatSchularbeiten !== false && f.HatSchularbeiten !== 'false',
        proSemester: fieldNum(f, 'ProSemester', 2),
        standardDauer: fieldNum(f, 'StandardDauer', 100)
    };
}

export function mapFachMetaToFields(meta) {
    return {
        Title: String(meta.name || meta.fachCode || '').trim() || 'Fach',
        FachCode: String(meta.fachCode || '').trim(),
        Farbe: String(meta.farbe || '').trim(),
        HatSchularbeiten: meta.hatSchularbeiten !== false,
        ProSemester: Number(meta.proSemester) || 2,
        StandardDauer: Number(meta.standardDauer) || 100
    };
}

export function mapFensterFromItem(item) {
    const f = (item && item.fields) || {};
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        terminfensterId: fieldStr(f, 'TerminfensterId'),
        titel: fieldStr(f, 'Title'),
        typ: fieldStr(f, 'Typ') === 'erlaubt' ? 'erlaubt' : 'gesperrt',
        startdatum: toIsoDateOnly(f.Startdatum) || '',
        enddatum: toIsoDateOnly(f.Enddatum) || '',
        beschreibung: fieldStr(f, 'Beschreibung')
    };
}

export function mapFensterToFields(win) {
    return {
        Title: String(win.titel || '').trim() || 'Terminfenster',
        TerminfensterId: String(win.terminfensterId || '').trim() || newEntityId('tf'),
        Typ: win.typ === 'erlaubt' ? 'erlaubt' : 'gesperrt',
        Startdatum: toIsoDateOnly(win.startdatum),
        Enddatum: toIsoDateOnly(win.enddatum),
        Beschreibung: String(win.beschreibung || '')
    };
}

export function mapRegelwerkFromItem(item) {
    const f = (item && item.fields) || {};
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        regelwerkId: fieldStr(f, 'RegelwerkId'),
        name: fieldStr(f, 'Title') || 'Regelwerk',
        maxProTag: fieldNum(f, 'MaxProTag', 1),
        maxProWoche: fieldNum(f, 'MaxProWoche', 2),
        ankuendigungsfristTage: fieldNum(f, 'AnkuendigungsfristTage', 7),
        sperreVorNotenkonferenzTage: fieldNum(f, 'SperreVorNotenkonferenzTage', 7),
        aktiv: f.Aktiv !== false && f.Aktiv !== 'false'
    };
}

export function mapRegelwerkToFields(rw) {
    return {
        Title: String(rw.name || '').trim() || 'Regelwerk',
        RegelwerkId: String(rw.regelwerkId || '').trim() || newEntityId('rw'),
        MaxProTag: Number(rw.maxProTag) || 1,
        MaxProWoche: Number(rw.maxProWoche) || 2,
        AnkuendigungsfristTage: Number(rw.ankuendigungsfristTage) || 7,
        SperreVorNotenkonferenzTage: Number(rw.sperreVorNotenkonferenzTage) || 7,
        Aktiv: rw.aktiv !== false
    };
}

/**
 * @param {Awaited<ReturnType<typeof resolvePlanerContext>>} ctx
 */
export async function loadAllPlanerData(ctx) {
    const tok = await token();
    const jobs = [
        fetchAllItems(tok, ctx.siteId, ctx.lists.schularbeiten.id),
        fetchAllItems(tok, ctx.siteId, ctx.lists.terminfenster.id),
        fetchAllItems(tok, ctx.siteId, ctx.lists.regelwerk.id)
    ];
    if (ctx.lists.fachMeta && ctx.lists.fachMeta.id) {
        jobs.push(fetchAllItems(tok, ctx.siteId, ctx.lists.fachMeta.id));
    }
    const results = await Promise.all(jobs);
    const saItems = results[0];
    const tfItems = results[1];
    const rwItems = results[2];
    const metaItems = results[3] || [];

    const items = saItems.map(mapSchularbeitFromItem);
    const windows = tfItems.map(mapFensterFromItem);
    const regelwerke = rwItems.map(mapRegelwerkFromItem);
    const active = regelwerke.find((r) => r.aktiv) || regelwerke[0] || null;
    const fachMeta = metaItems.map(mapFachMetaFromItem);

    return { items, windows, rules: active, allRules: regelwerke, fachMeta };
}

export async function createSchularbeitItem(ctx, sa) {
    const tok = await token();
    const fields = mapSchularbeitToFields(sa, { includeId: true });
    const created = await G().graphJson(
        'POST',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.schularbeiten.id) +
            '/items',
        tok,
        { fields },
        'v1.0'
    );
    return mapSchularbeitFromItem(created);
}

export async function updateSchularbeitItem(ctx, itemId, saPatch) {
    const tok = await token();
    const fields = mapSchularbeitToFields(saPatch, { includeId: true });
    await G().graphJson(
        'PATCH',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.schularbeiten.id) +
            '/items/' +
            encodeURIComponent(itemId) +
            '/fields',
        tok,
        fields,
        'v1.0'
    );
}

export async function deleteSchularbeitItem(ctx, itemId) {
    const tok = await token();
    await G().graphJson(
        'DELETE',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.schularbeiten.id) +
            '/items/' +
            encodeURIComponent(itemId),
        tok,
        undefined,
        'v1.0'
    );
}

export async function createFensterItem(ctx, win) {
    const tok = await token();
    const fields = mapFensterToFields(win);
    const created = await G().graphJson(
        'POST',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.terminfenster.id) +
            '/items',
        tok,
        { fields },
        'v1.0'
    );
    return mapFensterFromItem(created);
}

export async function deleteFensterItem(ctx, itemId) {
    const tok = await token();
    await G().graphJson(
        'DELETE',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.terminfenster.id) +
            '/items/' +
            encodeURIComponent(itemId),
        tok,
        undefined,
        'v1.0'
    );
}

export async function updateRegelwerkItem(ctx, itemId, rw) {
    const tok = await token();
    const fields = mapRegelwerkToFields(rw);
    await G().graphJson(
        'PATCH',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.regelwerk.id) +
            '/items/' +
            encodeURIComponent(itemId) +
            '/fields',
        tok,
        fields,
        'v1.0'
    );
}

export async function createFachMetaItem(ctx, meta) {
    if (!ctx.lists.fachMeta || !ctx.lists.fachMeta.id) {
        throw new Error('Liste SA-FachMeta fehlt – bitte Listen-Paket erneut ausführen.');
    }
    const tok = await token();
    const fields = mapFachMetaToFields(meta);
    const created = await G().graphJson(
        'POST',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.fachMeta.id) +
            '/items',
        tok,
        { fields },
        'v1.0'
    );
    return mapFachMetaFromItem(created);
}

export async function updateFachMetaItem(ctx, itemId, meta) {
    if (!ctx.lists.fachMeta || !ctx.lists.fachMeta.id) {
        throw new Error('Liste SA-FachMeta fehlt.');
    }
    const tok = await token();
    await G().graphJson(
        'PATCH',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.fachMeta.id) +
            '/items/' +
            encodeURIComponent(itemId) +
            '/fields',
        tok,
        mapFachMetaToFields(meta),
        'v1.0'
    );
}

export async function deleteFachMetaItem(ctx, itemId) {
    if (!ctx.lists.fachMeta || !ctx.lists.fachMeta.id) {
        throw new Error('Liste SA-FachMeta fehlt.');
    }
    const tok = await token();
    await G().graphJson(
        'DELETE',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.fachMeta.id) +
            '/items/' +
            encodeURIComponent(itemId),
        tok,
        undefined,
        'v1.0'
    );
}
