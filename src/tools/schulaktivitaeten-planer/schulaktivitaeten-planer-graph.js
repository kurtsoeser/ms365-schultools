/**
 * SharePoint-CRUD für Schulaktivitäten-Planer.
 */
import { LIST_TITLES, newEntityId } from './schulaktivitaeten-planer-schema.js';
import { toIsoDateOnly } from './schulaktivitaeten-planer-logic.js';

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

function fieldStr(fields, key) {
    const v = fields && fields[key];
    return v == null ? '' : String(v).trim();
}

function fieldNum(fields, key, fallback) {
    const n = Number(fields && fields[key]);
    return Number.isFinite(n) ? n : fallback;
}

export async function resolveAktContext(webUrl) {
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die SharePoint-Website-URL eintragen.');
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    const aktivitaeten = await findListByTitle(tok, siteId, LIST_TITLES.aktivitaeten);
    const regelwerk = await findListByTitle(tok, siteId, LIST_TITLES.regelwerk);
    if (!aktivitaeten) {
        throw new Error(
            'Liste „Schulaktivitaeten“ fehlt. Bitte zuerst „Schulaktivitäten-Listen“ anlegen.'
        );
    }
    return {
        webUrl: url,
        siteId,
        siteName: site.displayName || '',
        lists: {
            aktivitaeten: { id: String(aktivitaeten.id), webUrl: aktivitaeten.webUrl || '' },
            regelwerk: regelwerk
                ? { id: String(regelwerk.id), webUrl: regelwerk.webUrl || '' }
                : null
        }
    };
}

export function mapAktivitaetFromItem(item) {
    const f = (item && item.fields) || {};
    const start = toIsoDateOnly(f.Startdatum) || '';
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        aktivitaetId: fieldStr(f, 'AktivitaetId'),
        titel: fieldStr(f, 'Title'),
        typ: fieldStr(f, 'Typ') || 'Exkursion',
        klasseCode: fieldStr(f, 'KlasseCode'),
        lehrerCode: fieldStr(f, 'LehrerCode'),
        lehrerEmail: fieldStr(f, 'LehrerEmail').toLowerCase(),
        begleitung: fieldStr(f, 'Begleitung'),
        ort: fieldStr(f, 'Ort'),
        startdatum: start,
        enddatum: toIsoDateOnly(f.Enddatum) || start,
        startZeit: fieldStr(f, 'StartZeit'),
        endZeit: fieldStr(f, 'EndZeit'),
        status: fieldStr(f, 'Status') || 'beantragt',
        notiz: fieldStr(f, 'Notiz'),
        ablehnungsGrund: fieldStr(f, 'AblehnungsGrund'),
        beantragtVon: fieldStr(f, 'BeantragtVon'),
        genehmigtVon: fieldStr(f, 'GenehmigtVon'),
        genehmigtAm: fieldStr(f, 'GenehmigtAm'),
        verkehrsmittel: fieldStr(f, 'Verkehrsmittel'),
        kostenHinweis: fieldStr(f, 'KostenHinweis'),
        schulterminKey: fieldStr(f, 'SchulterminKey')
    };
}

export function mapAktivitaetToFields(akt, opts) {
    const includeId = !opts || opts.includeId !== false;
    const start = toIsoDateOnly(akt.startdatum);
    const end = toIsoDateOnly(akt.enddatum) || start;
    const fields = {
        Title: String(akt.titel || '').trim() || 'Schulaktivität',
        Typ: String(akt.typ || 'Exkursion').trim(),
        KlasseCode: String(akt.klasseCode || '').trim(),
        LehrerCode: String(akt.lehrerCode || '').trim(),
        LehrerEmail: String(akt.lehrerEmail || '').trim().toLowerCase(),
        Begleitung: String(akt.begleitung || ''),
        Ort: String(akt.ort || '').trim(),
        Startdatum: start,
        Enddatum: end,
        StartZeit: String(akt.startZeit || '').trim(),
        EndZeit: String(akt.endZeit || '').trim(),
        Status: String(akt.status || 'beantragt').toLowerCase(),
        Notiz: String(akt.notiz || ''),
        AblehnungsGrund: String(akt.ablehnungsGrund || ''),
        BeantragtVon: String(akt.beantragtVon || ''),
        GenehmigtVon: String(akt.genehmigtVon || ''),
        Verkehrsmittel: String(akt.verkehrsmittel || ''),
        KostenHinweis: String(akt.kostenHinweis || '')
    };
    if (akt.genehmigtAm) fields.GenehmigtAm = akt.genehmigtAm;
    if (akt.schulterminKey) fields.SchulterminKey = String(akt.schulterminKey);
    if (includeId) {
        fields.AktivitaetId = String(akt.aktivitaetId || '').trim() || newEntityId('akt');
    }
    return fields;
}

export function mapRegelwerkFromItem(item) {
    const f = (item && item.fields) || {};
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        title: fieldStr(f, 'Title'),
        regelwerkId: fieldStr(f, 'RegelwerkId'),
        minVorlaufTage: fieldNum(f, 'MinVorlaufTage', 7),
        maxGleichzeitigProKlasse: fieldNum(f, 'MaxGleichzeitigProKlasse', 1),
        aktiv: f.Aktiv !== false && f.Aktiv !== 'false'
    };
}

export async function loadAllAktData(ctx) {
    const tok = await token();
    const rows = await fetchAllItems(tok, ctx.siteId, ctx.lists.aktivitaeten.id);
    const items = rows.map(mapAktivitaetFromItem);
    let rules = null;
    if (ctx.lists.regelwerk) {
        const rw = await fetchAllItems(tok, ctx.siteId, ctx.lists.regelwerk.id);
        const mapped = rw.map(mapRegelwerkFromItem).filter((r) => r.aktiv);
        rules = mapped[0] || (rw[0] ? mapRegelwerkFromItem(rw[0]) : null);
    }
    return { items, rules };
}

export async function createAktivitaetItem(ctx, akt) {
    const tok = await token();
    const fields = mapAktivitaetToFields({
        ...akt,
        aktivitaetId: akt.aktivitaetId || newEntityId('akt')
    });
    const created = await G().graphJson(
        'POST',
        G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(ctx.lists.aktivitaeten.id) + '/items',
        tok,
        { fields },
        'v1.0'
    );
    return mapAktivitaetFromItem(created);
}

export async function updateAktivitaetItem(ctx, itemId, akt) {
    const tok = await token();
    const fields = mapAktivitaetToFields(akt, { includeId: true });
    await G().graphJson(
        'PATCH',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.aktivitaeten.id) +
            '/items/' +
            encodeURIComponent(itemId) +
            '/fields',
        tok,
        fields,
        'v1.0'
    );
    return { ...akt, itemId: String(itemId) };
}

export async function deleteAktivitaetItem(ctx, itemId) {
    const tok = await token();
    await G().graphJson(
        'DELETE',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.aktivitaeten.id) +
            '/items/' +
            encodeURIComponent(itemId),
        tok,
        undefined,
        'v1.0'
    );
}

export async function updateRegelwerkItem(ctx, itemId, patch) {
    if (!ctx.lists.regelwerk) throw new Error('Regelwerk-Liste fehlt.');
    const tok = await token();
    const fields = {
        MinVorlaufTage: Number(patch.minVorlaufTage) || 7,
        MaxGleichzeitigProKlasse: Number(patch.maxGleichzeitigProKlasse) || 1,
        Aktiv: patch.aktiv !== false
    };
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
