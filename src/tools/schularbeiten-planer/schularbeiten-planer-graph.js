/**
 * SharePoint-CRUD für Schularbeiten-Planer (Graph).
 */
import {
    LIST_TITLES,
    LIST_KEYS,
    newEntityId,
    DEFAULT_FACH_META_STANDARD_DAUER,
    DEFAULT_FACH_META_PRO_SEMESTER
} from './schularbeiten-planer-schema.js';
import { toIsoDateOnly, composeSchularbeitListItemTitle, normalizeBeginnUhrzeit } from './schularbeiten-planer-logic.js';
import { resolvePlanerList } from './schularbeiten-planer-lists.js';
import {
    normalizeSchuljahr,
    pickActiveRegelwerk,
    filterTerminfensterForSchuljahr,
    filterFachMetaForSchuljahr
} from './schularbeiten-planer-schuljahr.js';

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

export async function findListByDisplayName(tok, siteId, listTitle) {
    const title = String(listTitle || '').trim();
    const path =
        G().graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await G().graphJson('GET', path, tok, undefined, 'v1.0');
    return ((data && data.value) || [])[0] || null;
}

async function findListByTitle(tok, siteId, listTitle) {
    return findListByDisplayName(tok, siteId, listTitle);
}

async function fetchAllItems(tok, siteId, listId) {
    let path =
        G().graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(listId) +
        '/items?$expand=fields&$top=999';
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

    const rwRes = await resolvePlanerList(tok, siteId, 'regelwerk', findListByDisplayName);
    const tfRes = await resolvePlanerList(tok, siteId, 'terminfenster', findListByDisplayName);
    const saRes = await resolvePlanerList(tok, siteId, 'schularbeiten', findListByDisplayName);
    const metaRes = await resolvePlanerList(tok, siteId, 'fachMeta', findListByDisplayName);
    const regelwerk = rwRes && rwRes.list;
    const terminfenster = tfRes && tfRes.list;
    const schularbeiten = saRes && saRes.list;
    const fachMeta = metaRes && metaRes.list;
    if (!regelwerk || !terminfenster || !schularbeiten) {
        throw new Error(
            'Listen unvollständig. Bitte zuerst „Schularbeiten-Listen“ anlegen (SAP-Regelwerk, SAP-Terminfenster, SAP-Schularbeiten).'
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
    const titelCol = fieldStr(f, 'Titel');
    const themaCol = fieldStr(f, 'Thema');
    const spTitle = fieldStr(f, 'Title');
    let titel = titelCol;
    let thema = themaCol;
    if (!titelCol && !themaCol && spTitle) {
        titel = spTitle;
    }
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        schularbeitId: fieldStr(f, 'SchularbeitId'),
        titel,
        thema,
        fachCode: fieldStr(f, 'FachCode'),
        klasseCode: fieldStr(f, 'KlasseCode'),
        lehrerCode: fieldStr(f, 'LehrerCode'),
        lehrerEmail: fieldStr(f, 'LehrerEmail').toLowerCase(),
        datum: toIsoDateOnly(f.Datum) || '',
        beginnUhrzeit: normalizeBeginnUhrzeit(fieldStr(f, 'BeginnUhrzeit')),
        dauerMinuten: fieldNum(f, 'DauerMinuten', 100),
        semester: fieldStr(f, 'Semester') || 'WS',
        status: fieldStr(f, 'Status') || 'beantragt',
        notiz: fieldStr(f, 'Notiz'),
        ablehnungsGrund: fieldStr(f, 'AblehnungsGrund'),
        beantragtVon: fieldStr(f, 'BeantragtVon'),
        fixiertVon: fieldStr(f, 'FixiertVon'),
        fixiertAm: fieldStr(f, 'FixiertAm'),
        schulterminKey: fieldStr(f, 'SchulterminKey'),
        teamsCalendarEventId: fieldStr(f, 'TeamsCalendarEventId'),
        schuljahr: normalizeSchuljahr(fieldStr(f, 'Schuljahr'))
    };
}

/**
 * @param {object} sa
 * @param {{ includeId?: boolean }} [opts]
 */
export function mapSchularbeitToFields(sa, opts) {
    const includeId = !opts || opts.includeId !== false;
    const labels = opts && opts.labels;
    const fachLabel =
        labels && labels.fach && sa.fachCode ? labels.fach[String(sa.fachCode).trim()] : '';
    const klasseLabel =
        labels && labels.klasse && sa.klasseCode ? labels.klasse[String(sa.klasseCode).trim()] : '';
    const fields = {
        Title: composeSchularbeitListItemTitle(sa, {
            fachLabel: fachLabel || undefined,
            klasseLabel: klasseLabel || undefined
        }),
        Titel: String(sa.titel || '').trim(),
        Thema: String(sa.thema || '').trim(),
        FachCode: String(sa.fachCode || '').trim(),
        KlasseCode: String(sa.klasseCode || '').trim(),
        LehrerCode: String(sa.lehrerCode || '').trim(),
        LehrerEmail: String(sa.lehrerEmail || '').trim().toLowerCase(),
        Datum: toIsoDateOnly(sa.datum),
        BeginnUhrzeit: normalizeBeginnUhrzeit(sa.beginnUhrzeit),
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
    if (sa.teamsCalendarEventId) {
        fields.TeamsCalendarEventId = String(sa.teamsCalendarEventId);
    }
    if (includeId) {
        fields.SchularbeitId = String(sa.schularbeitId || '').trim() || newEntityId('sa');
    }
    const sj = normalizeSchuljahr(sa.schuljahr);
    if (sj) fields.Schuljahr = sj;
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
        proSemester: fieldNum(f, 'ProSemester', DEFAULT_FACH_META_PRO_SEMESTER),
        standardDauer: fieldNum(f, 'StandardDauer', DEFAULT_FACH_META_STANDARD_DAUER),
        schuljahr: normalizeSchuljahr(fieldStr(f, 'Schuljahr'))
    };
}

export function mapFachMetaToFields(meta) {
    const fields = {
        Title: String(meta.name || meta.fachCode || '').trim() || 'Fach',
        FachCode: String(meta.fachCode || '').trim(),
        Farbe: String(meta.farbe || '').trim(),
        HatSchularbeiten: meta.hatSchularbeiten !== false,
        ProSemester: Number(meta.proSemester) || DEFAULT_FACH_META_PRO_SEMESTER,
        StandardDauer: Number(meta.standardDauer) || DEFAULT_FACH_META_STANDARD_DAUER
    };
    const sj = normalizeSchuljahr(meta.schuljahr);
    if (sj) fields.Schuljahr = sj;
    return fields;
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
        beschreibung: fieldStr(f, 'Beschreibung'),
        schuljahr: normalizeSchuljahr(fieldStr(f, 'Schuljahr'))
    };
}

export function mapFensterToFields(win) {
    const fields = {
        Title: String(win.titel || '').trim() || 'Terminfenster',
        TerminfensterId: String(win.terminfensterId || '').trim() || newEntityId('tf'),
        Typ: win.typ === 'erlaubt' ? 'erlaubt' : 'gesperrt',
        Startdatum: toIsoDateOnly(win.startdatum),
        Enddatum: toIsoDateOnly(win.enddatum),
        Beschreibung: String(win.beschreibung || '')
    };
    const sj = normalizeSchuljahr(win.schuljahr);
    if (sj) fields.Schuljahr = sj;
    return fields;
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
        aktiv: f.Aktiv !== false && f.Aktiv !== 'false',
        schuljahr: normalizeSchuljahr(fieldStr(f, 'Schuljahr'))
    };
}

export function mapRegelwerkToFields(rw) {
    const fields = {
        Title: String(rw.name || '').trim() || 'Regelwerk',
        RegelwerkId: String(rw.regelwerkId || '').trim() || newEntityId('rw'),
        MaxProTag: Number(rw.maxProTag) || 1,
        MaxProWoche: Number(rw.maxProWoche) || 2,
        AnkuendigungsfristTage: Number(rw.ankuendigungsfristTage) || 7,
        SperreVorNotenkonferenzTage: Number(rw.sperreVorNotenkonferenzTage) || 7,
        Aktiv: rw.aktiv !== false
    };
    const sj = normalizeSchuljahr(rw.schuljahr);
    if (sj) fields.Schuljahr = sj;
    return fields;
}

/**
 * @param {Awaited<ReturnType<typeof resolvePlanerContext>>} ctx
 */
/**
 * @param {Awaited<ReturnType<typeof resolvePlanerContext>>} ctx
 * @param {{ schuljahr?: string }} [opts]
 */
export async function loadAllPlanerData(ctx, opts) {
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
    const windowsAll = tfItems.map(mapFensterFromItem);
    const regelwerke = rwItems.map(mapRegelwerkFromItem);
    const fachMetaAll = metaItems.map(mapFachMetaFromItem);
    const schuljahr = normalizeSchuljahr(opts && opts.schuljahr);
    const windows = filterTerminfensterForSchuljahr(windowsAll, schuljahr);
    const fachMeta = filterFachMetaForSchuljahr(fachMetaAll, schuljahr);
    const active = pickActiveRegelwerk(regelwerke, schuljahr);

    return {
        items,
        windows,
        windowsAll,
        rules: active,
        allRules: regelwerke,
        fachMeta,
        fachMetaAll
    };
}

export async function createSchularbeitItem(ctx, sa, mapOpts) {
    const tok = await token();
    const fields = mapSchularbeitToFields(sa, { includeId: true, ...(mapOpts || {}) });
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

export async function updateSchularbeitItem(ctx, itemId, saPatch, mapOpts) {
    const tok = await token();
    const fields = mapSchularbeitToFields(saPatch, { includeId: true, ...(mapOpts || {}) });
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
        throw new Error('Liste SAP-FachMeta fehlt – bitte Listen-Paket erneut ausführen.');
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
        throw new Error('Liste SAP-FachMeta fehlt.');
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
        throw new Error('Liste SAP-FachMeta fehlt.');
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

async function deleteListItemById(tok, siteId, listId, itemId) {
    await G().graphJson(
        'DELETE',
        G().graphPathSite(siteId) +
            '/lists/' +
            encodeURIComponent(listId) +
            '/items/' +
            encodeURIComponent(itemId),
        tok,
        undefined,
        'v1.0'
    );
    if (typeof G().sleep === 'function') await G().sleep(50);
}

/**
 * Alle Listeneinträge in den gewählten SAP-Listen löschen (Listen bleiben erhalten).
 * @param {Awaited<ReturnType<typeof resolvePlanerContext>>} ctx
 * @param {string[]} listKeys subset of LIST_KEYS
 * @param {(msg: { listKey: string, listTitle: string, done: number, total: number, phase?: string }) => void} [onProgress]
 * @returns {Promise<Record<string, number>>} gelöschte Anzahl je listKey
 */
export async function emptyPlanerLists(ctx, listKeys, onProgress) {
    if (!ctx || !ctx.lists) throw new Error('SharePoint-Kontext fehlt – bitte Site laden.');
    const keys = (listKeys || []).filter((k) => LIST_KEYS.includes(k));
    if (!keys.length) throw new Error('Keine gültigen Listen gewählt.');

    const tok = await token();
    const deleted = {};

    for (let ki = 0; ki < keys.length; ki++) {
        const listKey = keys[ki];
        const meta = ctx.lists[listKey];
        const listTitle = (meta && meta.displayName) || LIST_TITLES[listKey] || listKey;
        if (!meta || !meta.id) {
            throw new Error('Liste nicht gefunden: ' + listTitle);
        }
        if (onProgress) onProgress({ listKey, listTitle, done: 0, total: 0, phase: 'load' });
        const items = await fetchAllItems(tok, ctx.siteId, meta.id);
        const total = items.length;
        deleted[listKey] = 0;
        for (let i = 0; i < items.length; i++) {
            const id = items[i] && items[i].id;
            if (id == null) continue;
            await deleteListItemById(tok, ctx.siteId, meta.id, id);
            deleted[listKey]++;
            if (onProgress) {
                onProgress({ listKey, listTitle, done: deleted[listKey], total, phase: 'delete' });
            }
        }
    }
    return deleted;
}
