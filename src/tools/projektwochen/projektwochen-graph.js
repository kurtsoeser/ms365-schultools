/**
 * SharePoint-CRUD für Projektwochen (Graph).
 */
import { LIST_TITLES, newEntityId } from './projektwochen-schema.js';
import { toIsoDateOnly, toDatetimeLocalValue, weekdayLabelDeFromIso } from './projektwochen-logic.js';

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

async function listColumnNames(tok, siteId, listId) {
    const path =
        G().graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(listId) +
        '/columns?$select=name&$top=200';
    const data = await G().graphJson('GET', path, tok, undefined, 'v1.0');
    return new Set(
        ((data && data.value) || [])
            .map((c) => String((c && c.name) || '').trim())
            .filter(Boolean)
    );
}

/**
 * Interne Spalte für Mo–Fr: neue Listen „Wochentag“, Legacy „Tag“.
 * @param {Set<string>|string[]} colNames
 */
export function resolveWeekdayFieldName(colNames) {
    const set = colNames instanceof Set ? colNames : new Set(colNames || []);
    if (set.has('Wochentag')) return 'Wochentag';
    if (set.has('Tag')) return 'Tag';
    return 'Wochentag';
}

/**
 * SPO-Felder an Listen-Schema anpassen (Wochentag ↔ Tag).
 * @param {Record<string, unknown>} fields
 * @param {string} [weekdayField]
 */
export function adaptAngebotFieldsForList(fields, weekdayField) {
    const f = Object.assign({}, fields || {});
    const day = f.Wochentag != null && f.Wochentag !== '' ? f.Wochentag : f.Tag;
    delete f.Wochentag;
    delete f.Tag;
    const key = weekdayField === 'Tag' ? 'Tag' : 'Wochentag';
    if (day != null && day !== '') f[key] = day;
    return f;
}

/**
 * @param {string} webUrl
 */
export async function resolvePwContext(webUrl) {
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die SharePoint-Website-URL eintragen.');
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    const aktionen = await findListByTitle(tok, siteId, LIST_TITLES.aktionen);
    const angebote = await findListByTitle(tok, siteId, LIST_TITLES.angebote);
    if (!aktionen || !angebote) {
        throw new Error(
            'Listen unvollständig. Bitte zuerst unter Einstellungen die Projektwochen-Listen anlegen (PW-Aktionen, PW-Angebote).'
        );
    }
    const angebotCols = await listColumnNames(tok, siteId, angebote.id);
    const weekdayField = resolveWeekdayFieldName(angebotCols);
    if (!angebotCols.has('Wochentag') && !angebotCols.has('Tag')) {
        throw new Error(
            'In PW-Angebote fehlt die Spalte Wochentag (bzw. Legacy „Tag“). Bitte unter Einstellungen „Listen anlegen / aktualisieren“ ausführen.'
        );
    }
    return {
        webUrl: url,
        siteId,
        siteName: site.displayName || '',
        weekdayField,
        lists: {
            aktionen: { id: String(aktionen.id), webUrl: aktionen.webUrl || '' },
            angebote: { id: String(angebote.id), webUrl: angebote.webUrl || '' }
        }
    };
}

export function mapAktionFromItem(item) {
    const f = (item && item.fields) || {};
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        title: fieldStr(f, 'Title'),
        aktionId: fieldStr(f, 'AktionId'),
        startdatum: toIsoDateOnly(f.Startdatum) || '',
        enddatum: toIsoDateOnly(f.Enddatum) || '',
        buchungAbDefault: fieldStr(f, 'BuchungAbDefault'),
        bookingsBusinessId: fieldStr(f, 'BookingsBusinessId'),
        bookingsBusinessName: fieldStr(f, 'BookingsBusinessName'),
        status: fieldStr(f, 'Status') || 'entwurf',
        beschreibung: fieldStr(f, 'Beschreibung')
    };
}

export function mapAktionToFields(obj) {
    return {
        Title: String(obj.title || '').trim() || 'Projektwoche',
        AktionId: String(obj.aktionId || '').trim() || newEntityId('pw'),
        Startdatum: toIsoDateOnly(obj.startdatum) || null,
        Enddatum: toIsoDateOnly(obj.enddatum) || null,
        BuchungAbDefault: String(obj.buchungAbDefault || '').trim() || null,
        BookingsBusinessId: String(obj.bookingsBusinessId || '').trim() || null,
        BookingsBusinessName: String(obj.bookingsBusinessName || '').trim() || null,
        Status: String(obj.status || 'entwurf'),
        Beschreibung: String(obj.beschreibung || '')
    };
}

export function mapAngebotFromItem(item) {
    const f = (item && item.fields) || {};
    const datum = toIsoDateOnly(f.Datum) || '';
    const tag = fieldStr(f, 'Wochentag') || fieldStr(f, 'Tag') || weekdayLabelDeFromIso(datum);
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        title: fieldStr(f, 'Title'),
        angebotId: fieldStr(f, 'AngebotId'),
        aktionId: fieldStr(f, 'AktionId'),
        beschreibung: fieldStr(f, 'Beschreibung'),
        hinweisEltern: fieldStr(f, 'HinweisEltern'),
        ort: fieldStr(f, 'Ort'),
        treffpunkt: fieldStr(f, 'Treffpunkt'),
        tag,
        datum,
        slot: fieldStr(f, 'Slot') || 'vormittag',
        startzeit: fieldStr(f, 'Startzeit'),
        endzeit: fieldStr(f, 'Endzeit'),
        kapazitaet: fieldNum(f, 'Kapazitaet', 20),
        preisEuro: fieldNum(f, 'PreisEuro', 0),
        kostenHinweis: fieldStr(f, 'KostenHinweis'),
        zielklassen: fieldStr(f, 'Zielklassen') || 'alle',
        lehrerCode: fieldStr(f, 'LehrerCode'),
        lehrerEmail: fieldStr(f, 'LehrerEmail').toLowerCase(),
        begleitung: fieldStr(f, 'Begleitung'),
        kategorie: fieldStr(f, 'Kategorie') || 'sonstiges',
        status: fieldStr(f, 'Status') || 'beantragt',
        buchungAb: fieldStr(f, 'BuchungAb'),
        ablehnungsGrund: fieldStr(f, 'AblehnungsGrund'),
        bookingsServiceId: fieldStr(f, 'BookingsServiceId'),
        bookingsBookingUrl: fieldStr(f, 'BookingsBookingUrl'),
        syncStatus: fieldStr(f, 'SyncStatus'),
        syncFehler: fieldStr(f, 'SyncFehler'),
        syncAm: fieldStr(f, 'SyncAm'),
        beantragtVon: fieldStr(f, 'BeantragtVon'),
        freigegebenVon: fieldStr(f, 'FreigegebenVon'),
        freigegebenAm: fieldStr(f, 'FreigegebenAm'),
        notizIntern: fieldStr(f, 'NotizIntern')
    };
}

/**
 * @param {object} obj
 * @param {{ includeAdminFields?: boolean, weekdayField?: string }} [opts]
 */
export function mapAngebotToFields(obj, opts) {
    const includeAdmin = !opts || opts.includeAdminFields !== false;
    const weekdayField = opts && opts.weekdayField === 'Tag' ? 'Tag' : 'Wochentag';
    const datum = toIsoDateOnly(obj.datum) || null;
    const tag = String(obj.tag || '').trim() || weekdayLabelDeFromIso(datum) || null;
    /** @type {Record<string, unknown>} */
    const fields = {
        Title: String(obj.title || '').trim() || 'Angebot',
        AngebotId: String(obj.angebotId || '').trim() || newEntityId('ang'),
        AktionId: String(obj.aktionId || '').trim() || null,
        Beschreibung: String(obj.beschreibung || ''),
        HinweisEltern: String(obj.hinweisEltern || ''),
        Ort: String(obj.ort || ''),
        Treffpunkt: String(obj.treffpunkt || ''),
        Datum: datum,
        Slot: String(obj.slot || 'vormittag'),
        Startzeit: String(obj.startzeit || ''),
        Endzeit: String(obj.endzeit || ''),
        Kapazitaet: Number(obj.kapazitaet) || 0,
        PreisEuro: Number(obj.preisEuro) || 0,
        KostenHinweis: String(obj.kostenHinweis || ''),
        Zielklassen: String(obj.zielklassen || 'alle'),
        LehrerCode: String(obj.lehrerCode || ''),
        LehrerEmail: String(obj.lehrerEmail || '').trim().toLowerCase(),
        Begleitung: String(obj.begleitung || ''),
        Kategorie: String(obj.kategorie || 'sonstiges'),
        Status: String(obj.status || 'beantragt'),
        BeantragtVon: String(obj.beantragtVon || '')
    };
    fields[weekdayField] = tag;
    if (includeAdmin) {
        fields.BuchungAb = String(obj.buchungAb || '').trim() || null;
        fields.AblehnungsGrund = String(obj.ablehnungsGrund || '');
        fields.BookingsServiceId = String(obj.bookingsServiceId || '') || null;
        fields.BookingsBookingUrl = String(obj.bookingsBookingUrl || '') || null;
        fields.SyncStatus = String(obj.syncStatus || '') || null;
        fields.SyncFehler = String(obj.syncFehler || '') || null;
        fields.SyncAm = String(obj.syncAm || '') || null;
        fields.FreigegebenVon = String(obj.freigegebenVon || '') || null;
        fields.FreigegebenAm = String(obj.freigegebenAm || '') || null;
        fields.NotizIntern = String(obj.notizIntern || '');
    }
    return fields;
}

/**
 * @param {object} ctx
 */
export async function loadAllPwData(ctx) {
    const tok = await token();
    const siteId = ctx.siteId;
    const [aktionItems, angebotItems] = await Promise.all([
        fetchAllItems(tok, siteId, ctx.lists.aktionen.id),
        fetchAllItems(tok, siteId, ctx.lists.angebote.id)
    ]);
    return {
        aktionen: aktionItems.map(mapAktionFromItem),
        angebote: angebotItems.map(mapAngebotFromItem)
    };
}

export async function createAngebotItem(ctx, angebot) {
    const tok = await token();
    const fields = mapAngebotToFields(angebot, { weekdayField: ctx && ctx.weekdayField });
    const created = await G().graphJson(
        'POST',
        G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(ctx.lists.angebote.id) + '/items',
        tok,
        { fields },
        'v1.0'
    );
    return mapAngebotFromItem(created);
}

export async function updateAngebotItem(ctx, itemId, angebot) {
    const tok = await token();
    const fields = mapAngebotToFields(angebot, { weekdayField: ctx && ctx.weekdayField });
    await G().graphJson(
        'PATCH',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.angebote.id) +
            '/items/' +
            encodeURIComponent(itemId) +
            '/fields',
        tok,
        fields,
        'v1.0'
    );
    return { ...angebot, itemId: String(itemId) };
}

export async function deleteAngebotItem(ctx, itemId) {
    const tok = await token();
    await G().graphJson(
        'DELETE',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.angebote.id) +
            '/items/' +
            encodeURIComponent(itemId),
        tok,
        undefined,
        'v1.0'
    );
}

export async function updateAktionItem(ctx, itemId, aktion) {
    const tok = await token();
    const fields = mapAktionToFields(aktion);
    await G().graphJson(
        'PATCH',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.lists.aktionen.id) +
            '/items/' +
            encodeURIComponent(itemId) +
            '/fields',
        tok,
        fields,
        'v1.0'
    );
    return { ...aktion, itemId: String(itemId) };
}

export { toDatetimeLocalValue };
