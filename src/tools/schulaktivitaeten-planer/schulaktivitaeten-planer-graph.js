/**
 * SharePoint-CRUD für Schulaktivitäten-Planer.
 */
import {
    LIST_TITLES,
    AKTIVITAETEN_COLUMNS,
    REGELWERK_COLUMNS,
    newEntityId,
    toGraphColumnBody
} from './schulaktivitaeten-planer-schema.js';
import {
    toIsoDateOnly,
    joinCodeList,
    parseCodeList,
    klasseCodesOf,
    normalizeTeilnahme,
    resolveAktOwnership,
    assertAktOwnership
} from './schulaktivitaeten-planer-logic.js';
import { parseUeNumber, applyUeToDraft } from './schulaktivitaeten-planer-zeitraster.js';

/** Spalten, ohne die Meine-Aktivitäten nicht funktioniert. */
const OWNERSHIP_COLUMN_NAMES = ['BeantragtVon', 'LehrerEmail', 'LehrerCode'];

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
        // createdBy darf nicht $expand-ed werden; $select liefert die Identity.
        '/items?$select=id,createdBy&$expand=fields&$top=100';
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
 * SharePoint-REST-Fallback (wie Browser-UI), wenn Graph 0 Zeilen / 403 liefert.
 * @param {string} siteWebUrl
 * @param {string} listTitle
 */
async function fetchAllItemsViaSpoRest(siteWebUrl, listTitle) {
    const api = G();
    const url = String(siteWebUrl || '').trim().replace(/\/+$/, '');
    const title = String(listTitle || '').trim();
    if (!url || !title) return [];
    if (typeof api.getSpoRequestDigest !== 'function' || typeof api.spoRestFetch !== 'function') {
        return [];
    }
    let host = '';
    try {
        host = new URL(url).hostname;
    } catch {
        return [];
    }
    const spoScopes = [
        ['https://' + host + '/AllSites.Read', 'https://graph.microsoft.com/User.Read'],
        ['https://' + host + '/AllSites.Write', 'https://graph.microsoft.com/User.Read']
    ];
    let spoToken = '';
    for (let i = 0; i < spoScopes.length; i++) {
        try {
            if (typeof window !== 'undefined' && typeof window.ms365AuthAcquireTokenSilent === 'function') {
                try {
                    spoToken = await window.ms365AuthAcquireTokenSilent(spoScopes[i]);
                } catch {
                    spoToken = await api.getGraphToken(spoScopes[i]);
                }
            } else {
                spoToken = await api.getGraphToken(spoScopes[i]);
            }
            if (spoToken) break;
        } catch {
            /* next scope */
        }
    }
    if (!spoToken) return [];
    const digest = await api.getSpoRequestDigest(url, spoToken);
    const esc = title.replace(/'/g, "''");
    const select = [
        'Id',
        'Title',
        'AktivitaetId',
        'Typ',
        'KlasseCode',
        'KlasseCodes',
        'TeilnahmeModus',
        'GruppenText',
        'TeilnehmendeText',
        'NichtTeilnehmende',
        'LehrerCode',
        'LehrerEmail',
        'Begleitung',
        'BegleitungCodes',
        'ExterneVortragende',
        'Ort',
        'Ablauf',
        'Unterrichtsbezug',
        'Startdatum',
        'Enddatum',
        'StartUe',
        'EndUe',
        'StartZeit',
        'EndZeit',
        'Status',
        'Notiz',
        'AblehnungsGrund',
        'RueckfrageText',
        'BeantragtVon',
        'GenehmigtVon',
        'GenehmigtAm',
        'UntisErledigt',
        'Verkehrsmittel',
        'KostenHinweis',
        'Author/EMail',
        'Author/Title'
    ].join(',');
    /** @type {object[]} */
    const out = [];
    let apiPath =
        "/_api/web/lists/getbytitle('" +
        esc +
        "')/items?$top=100&$select=" +
        select +
        '&$expand=Author';
    while (apiPath) {
        const res = await api.spoRestFetch(url, spoToken, digest, 'GET', apiPath);
        if (!res.ok) {
            throw new Error(
                'SPO „' + title + '": HTTP ' + res.status + ' ' + String(res.text || '').slice(0, 180)
            );
        }
        const rows = (res.data && res.data.value) || [];
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i] || {};
            const authorMail =
                (row.Author && (row.Author.EMail || row.Author.Email || row.Author.UserPrincipalName)) ||
                '';
            out.push({
                id: row.Id != null ? row.Id : row.id,
                fields: row,
                createdBy: authorMail
                    ? { user: { email: String(authorMail).toLowerCase() } }
                    : undefined
            });
        }
        apiPath = String(
            (res.data && (res.data['@odata.nextLink'] || res.data['odata.nextLink'])) || ''
        ).trim();
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

function fieldBool(fields, key) {
    const v = fields && fields[key];
    return v === true || v === 'true' || v === 1 || v === '1';
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
    const ctx = {
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
    // Fehlende Spalten (z. B. BegleitungCodes) nachziehen, damit Writes nicht still scheitern.
    try {
        await ensureAktListColumns(ctx, tok);
    } catch {
        /* Setup-Hinweis bleibt dem Health-Check / Listen-Setup vorbehalten */
    }
    return ctx;
}

/**
 * Stellt sicher, dass Schema-Spalten auf vorhandenen Listen existieren.
 * @param {{ siteId: string, lists: { aktivitaeten?: { id: string }, regelwerk?: { id: string }|null } }} ctx
 * @param {string} [tok]
 */
export async function ensureAktListColumns(ctx, tok) {
    const accessToken = tok || (await token());
    const ensure = async (listId, defs) => {
        if (!listId) return { added: 0 };
        const colsPath =
            G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(listId) +
            '/columns?$top=200';
        const colsData = await G().graphJson('GET', colsPath, accessToken, undefined, 'v1.0');
        const byName = new Set(
            ((colsData && colsData.value) || [])
                .map((c) => String((c && c.name) || '').trim())
                .filter(Boolean)
        );
        const base =
            G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
        let added = 0;
        for (const def of defs || []) {
            if (!def || !def.name || byName.has(def.name)) continue;
            await G().graphJson('POST', base, accessToken, toGraphColumnBody(def), 'v1.0');
            byName.add(def.name);
            added += 1;
        }
        return { added };
    };
    const a = await ensure(ctx.lists && ctx.lists.aktivitaeten && ctx.lists.aktivitaeten.id, AKTIVITAETEN_COLUMNS);
    const r = await ensure(ctx.lists && ctx.lists.regelwerk && ctx.lists.regelwerk.id, REGELWERK_COLUMNS);
    return { aktivitaeten: a, regelwerk: r };
}

function authorEmailFromItem(item) {
    const created =
        (item && item.createdBy && item.createdBy.user && item.createdBy.user.email) ||
        (item && item.createdBy && item.createdBy.user && item.createdBy.user.userPrincipalName) ||
        '';
    if (created) return String(created).trim().toLowerCase();
    const f = (item && item.fields) || {};
    if (f.Author && typeof f.Author === 'object') {
        const a = f.Author;
        return String(a.Email || a.EMail || a.UserPrincipalName || a.LookupValue || '')
            .trim()
            .toLowerCase();
    }
    return String(f.AuthorEmail || f.Author0 || '')
        .trim()
        .toLowerCase();
}

export function mapAktivitaetFromItem(item) {
    const f = (item && item.fields) || {};
    const start = toIsoDateOnly(f.Startdatum) || '';
    const teilnahme = normalizeTeilnahme({
        klasseCode: fieldStr(f, 'KlasseCode'),
        klasseCodes: fieldStr(f, 'KlasseCodes'),
        teilnahmeModus: fieldStr(f, 'TeilnahmeModus'),
        gruppenText: fieldStr(f, 'GruppenText'),
        teilnehmendeText: fieldStr(f, 'TeilnehmendeText'),
        nichtTeilnehmende: fieldStr(f, 'NichtTeilnehmende')
    });
    const authorEmail = authorEmailFromItem(item);
    const beantragtVon = fieldStr(f, 'BeantragtVon').toLowerCase() || authorEmail;
    const lehrerEmail = fieldStr(f, 'LehrerEmail').toLowerCase() || beantragtVon || authorEmail;
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        aktivitaetId: fieldStr(f, 'AktivitaetId'),
        titel: fieldStr(f, 'Title'),
        typ: fieldStr(f, 'Typ') || 'Exkursion',
        klasseCode: teilnahme.klasseCode,
        klasseCodes: teilnahme.klasseCodes,
        teilnahmeModus: teilnahme.teilnahmeModus,
        gruppenText: teilnahme.gruppenText,
        teilnehmendeText: teilnahme.teilnehmendeText,
        nichtTeilnehmende: teilnahme.nichtTeilnehmende,
        lehrerCode: fieldStr(f, 'LehrerCode'),
        lehrerEmail,
        authorEmail,
        begleitung: fieldStr(f, 'Begleitung'),
        begleitungCodes: joinCodeList(
            fieldStr(f, 'BegleitungCodes') ||
                // Legacy: wenn nur Begleitung Codes enthielt
                (looksLikeCodeList(fieldStr(f, 'Begleitung')) ? fieldStr(f, 'Begleitung') : '')
        ),
        externeVortragende: fieldStr(f, 'ExterneVortragende'),
        ort: fieldStr(f, 'Ort'),
        ablauf: fieldStr(f, 'Ablauf'),
        unterrichtsbezug: fieldStr(f, 'Unterrichtsbezug'),
        startdatum: start,
        enddatum: toIsoDateOnly(f.Enddatum) || start,
        startUe: parseUeNumber(f.StartUe) || '',
        endUe: parseUeNumber(f.EndUe) || '',
        startZeit: fieldStr(f, 'StartZeit'),
        endZeit: fieldStr(f, 'EndZeit'),
        status: fieldStr(f, 'Status') || 'beantragt',
        notiz: fieldStr(f, 'Notiz'),
        ablehnungsGrund: fieldStr(f, 'AblehnungsGrund'),
        rueckfrageText: fieldStr(f, 'RueckfrageText'),
        kvStellungnahme: fieldStr(f, 'KvStellungnahme'),
        kvStellungnahmeVon: fieldStr(f, 'KvStellungnahmeVon').toLowerCase(),
        kvStellungnahmeAm: fieldStr(f, 'KvStellungnahmeAm'),
        aenderungsLog: fieldStr(f, 'AenderungsLog'),
        beantragtVon,
        genehmigtVon: fieldStr(f, 'GenehmigtVon').toLowerCase(),
        genehmigtAm: fieldStr(f, 'GenehmigtAm'),
        untisErledigt: fieldBool(f, 'UntisErledigt'),
        untisErledigtVon: fieldStr(f, 'UntisErledigtVon'),
        untisErledigtAm: fieldStr(f, 'UntisErledigtAm'),
        verkehrsmittel: fieldStr(f, 'Verkehrsmittel'),
        kostenHinweis: fieldStr(f, 'KostenHinweis'),
        schulterminKey: fieldStr(f, 'SchulterminKey')
    };
}

function looksLikeCodeList(raw) {
    const parts = parseCodeList(raw);
    if (!parts.length) return false;
    return parts.every((p) => /^[A-Za-z0-9_-]{1,12}$/.test(p) && !/\s/.test(p));
}

/**
 * @param {object} akt
 * @param {{ includeId?: boolean, teacherLabels?: Record<string,string> }} [opts]
 */
/**
 * @param {object} akt
 * @param {{ includeId?: boolean, teacherLabels?: Record<string,string>, accountEmail?: string, teachers?: object[], requireOwnership?: boolean }} [opts]
 */
export function mapAktivitaetToFields(akt, opts) {
    const includeId = !opts || opts.includeId !== false;
    const teacherLabels = (opts && opts.teacherLabels) || {};
    const ownership = resolveAktOwnership(akt, {
        accountEmail: opts && opts.accountEmail,
        teachers: opts && opts.teachers
    });
    if (!opts || opts.requireOwnership !== false) {
        assertAktOwnership(ownership);
    }
    const start = toIsoDateOnly(akt.startdatum);
    const end = toIsoDateOnly(akt.enddatum) || start;
    const teilnahme = normalizeTeilnahme(akt);
    const begleitungCodes = joinCodeList(akt.begleitungCodes || (looksLikeCodeList(akt.begleitung) ? akt.begleitung : ''));
    const begleitungLabel =
        String(akt.begleitung || '').trim() && !looksLikeCodeList(akt.begleitung)
            ? String(akt.begleitung).trim()
            : parseCodeList(begleitungCodes)
                  .map((c) => teacherLabels[c] || c)
                  .join('; ');
    const withUe = applyUeToDraft(akt).draft;
    const fields = {
        Title: String(akt.titel || '').trim() || 'Schulaktivität',
        Typ: String(akt.typ || 'Exkursion').trim(),
        KlasseCode: teilnahme.klasseCode || klasseCodesOf(akt)[0] || '',
        KlasseCodes: teilnahme.klasseCodes,
        TeilnahmeModus: teilnahme.teilnahmeModus,
        GruppenText: teilnahme.gruppenText,
        TeilnehmendeText: teilnahme.teilnehmendeText,
        NichtTeilnehmende: teilnahme.nichtTeilnehmende,
        LehrerCode: ownership.lehrerCode,
        LehrerEmail: ownership.lehrerEmail,
        Begleitung: begleitungLabel || begleitungCodes,
        BegleitungCodes: begleitungCodes,
        ExterneVortragende: String(akt.externeVortragende || '').trim(),
        Ort: String(akt.ort || '').trim(),
        Ablauf: String(akt.ablauf || ''),
        Unterrichtsbezug: String(akt.unterrichtsbezug || '').trim(),
        Startdatum: start,
        Enddatum: end,
        StartZeit: String(withUe.startZeit || '').trim(),
        EndZeit: String(withUe.endZeit || '').trim(),
        Status: String(akt.status || 'beantragt').toLowerCase(),
        ...(parseUeNumber(withUe.startUe) != null ? { StartUe: parseUeNumber(withUe.startUe) } : {}),
        ...(parseUeNumber(withUe.endUe) != null ? { EndUe: parseUeNumber(withUe.endUe) } : {}),
        Notiz: String(akt.notiz || ''),
        AblehnungsGrund: String(akt.ablehnungsGrund || ''),
        RueckfrageText: String(akt.rueckfrageText || ''),
        KvStellungnahme: String(akt.kvStellungnahme || ''),
        KvStellungnahmeVon: String(akt.kvStellungnahmeVon || '').trim().toLowerCase(),
        AenderungsLog: String(akt.aenderungsLog || ''),
        BeantragtVon: ownership.beantragtVon,
        GenehmigtVon: String(akt.genehmigtVon || '').trim().toLowerCase(),
        UntisErledigt: !!akt.untisErledigt,
        UntisErledigtVon: String(akt.untisErledigtVon || ''),
        Verkehrsmittel: String(akt.verkehrsmittel || ''),
        KostenHinweis: String(akt.kostenHinweis || '')
    };
    if (akt.genehmigtAm) fields.GenehmigtAm = akt.genehmigtAm;
    if (akt.kvStellungnahmeAm) fields.KvStellungnahmeAm = akt.kvStellungnahmeAm;
    if (akt.untisErledigtAm) fields.UntisErledigtAm = akt.untisErledigtAm;
    if (akt.schulterminKey) fields.SchulterminKey = String(akt.schulterminKey);
    if (includeId) {
        fields.AktivitaetId = String(akt.aktivitaetId || '').trim() || newEntityId('akt');
    }
    return fields;
}

/**
 * Ownership-Spalten müssen existieren – sonst stilles Leerschreiben.
 * @param {{ siteId: string, lists: { aktivitaeten?: { id: string } } }} ctx
 * @param {string} tok
 */
async function ensureOwnershipColumnsOrThrow(ctx, tok) {
    const listId = ctx && ctx.lists && ctx.lists.aktivitaeten && ctx.lists.aktivitaeten.id;
    if (!listId) throw new Error('Aktivitäten-Liste fehlt.');
    try {
        await ensureAktListColumns(ctx, tok);
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e);
        throw new Error(
            'Spalten für BeantragtVon/LehrerEmail/LehrerCode konnten nicht angelegt werden. IT: Listen-Assistent erneut ausführen. Details: ' +
                msg
        );
    }
    const colsPath =
        G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(listId) + '/columns?$top=200';
    const colsData = await G().graphJson('GET', colsPath, tok, undefined, 'v1.0');
    const names = new Set(
        ((colsData && colsData.value) || [])
            .map((c) => String((c && c.name) || '').trim().toLowerCase())
            .filter(Boolean)
    );
    const missing = OWNERSHIP_COLUMN_NAMES.filter((n) => !names.has(n.toLowerCase()));
    if (missing.length) {
        throw new Error(
            'In der SharePoint-Liste „Schulaktivitaeten“ fehlen Spalten: ' +
                missing.join(', ') +
                '. Bitte als IT unter „Schulaktivitäten-Listen“ das Paket erneut anlegen/prüfen.'
        );
    }
}

function parseKostenLimitsJson(raw) {
    const s = String(raw || '').trim();
    if (!s) return [];
    try {
        const parsed = JSON.parse(s);
        if (!Array.isArray(parsed)) return [];
        return parsed
            .map((r) => ({
                typ: String((r && (r.typ || r.Typ)) || '').trim(),
                okBis: Number(r && (r.okBis != null ? r.okBis : r.OkBis)) || 0,
                warnBis: Number(r && (r.warnBis != null ? r.warnBis : r.WarnBis)) || 0
            }))
            .filter((r) => r.typ);
    } catch {
        return [];
    }
}

export function mapRegelwerkFromItem(item) {
    const f = (item && item.fields) || {};
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        title: fieldStr(f, 'Title'),
        regelwerkId: fieldStr(f, 'RegelwerkId'),
        minVorlaufTage: fieldNum(f, 'MinVorlaufTage', 7),
        maxGleichzeitigProKlasse: fieldNum(f, 'MaxGleichzeitigProKlasse', 1),
        genehmigungPfad: fieldStr(f, 'GenehmigungPfad') || 'direktion',
        vorlaufModus: fieldStr(f, 'VorlaufModus') || 'sperre',
        kostenLimits: parseKostenLimitsJson(fieldStr(f, 'KostenLimitsJson')),
        aktiv: f.Aktiv !== false && f.Aktiv !== 'false'
    };
}

export async function loadAllAktData(ctx) {
    const tok = await token();
    /** @type {object[]} */
    let rows = [];
    let source = 'graph';
    try {
        rows = await fetchAllItems(tok, ctx.siteId, ctx.lists.aktivitaeten.id);
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e);
        if (!/403|accessdenied|forbidden/i.test(msg) && !(e && e.status === 403)) throw e;
        rows = [];
        source = 'graph-error';
    }
    if (!rows.length && ctx.webUrl) {
        try {
            const spoRows = await fetchAllItemsViaSpoRest(ctx.webUrl, LIST_TITLES.aktivitaeten);
            if (spoRows.length) {
                rows = spoRows;
                source = 'spo-rest';
            }
        } catch {
            /* Graph-Ergebnis (leer) behalten */
        }
    }
    const items = rows.map(mapAktivitaetFromItem);
    let rules = null;
    if (ctx.lists.regelwerk) {
        try {
            let rw = await fetchAllItems(tok, ctx.siteId, ctx.lists.regelwerk.id);
            if (!rw.length && ctx.webUrl) {
                try {
                    rw = await fetchAllItemsViaSpoRest(ctx.webUrl, LIST_TITLES.regelwerk);
                } catch {
                    /* ignore */
                }
            }
            const mapped = rw.map(mapRegelwerkFromItem).filter((r) => r.aktiv);
            rules = mapped[0] || (rw[0] ? mapRegelwerkFromItem(rw[0]) : null);
        } catch {
            rules = null;
        }
    }
    return { items, rules, source };
}

/**
 * Graph-Create liefert oft unvollständige fields – Ownership aus Draft nachziehen.
 * @param {object} mapped
 * @param {object} akt
 */
function mergeCreatedAktivitaet(mapped, akt) {
    const a = akt || {};
    const m = mapped || {};
    return {
        ...a,
        ...m,
        itemId: m.itemId || '',
        aktivitaetId: m.aktivitaetId || a.aktivitaetId || '',
        titel: m.titel || a.titel || '',
        status: m.status || a.status || 'beantragt',
        beantragtVon: m.beantragtVon || String(a.beantragtVon || '').toLowerCase(),
        lehrerEmail: m.lehrerEmail || String(a.lehrerEmail || '').toLowerCase(),
        lehrerCode: m.lehrerCode || String(a.lehrerCode || ''),
        begleitungCodes: m.begleitungCodes || a.begleitungCodes || '',
        begleitung: m.begleitung || a.begleitung || '',
        klasseCodes: m.klasseCodes || a.klasseCodes || '',
        klasseCode: m.klasseCode || a.klasseCode || '',
        startdatum: m.startdatum || a.startdatum || '',
        enddatum: m.enddatum || a.enddatum || '',
        authorEmail: m.authorEmail || String(a.beantragtVon || a.lehrerEmail || '').toLowerCase()
    };
}

export async function createAktivitaetItem(ctx, akt, opts) {
    const tok = await token();
    await ensureOwnershipColumnsOrThrow(ctx, tok);
    const ownership = resolveAktOwnership(akt, {
        accountEmail: opts && opts.accountEmail,
        teachers: opts && opts.teachers
    });
    assertAktOwnership(ownership);
    const aktWithOwn = { ...akt, ...ownership, aktivitaetId: akt.aktivitaetId || newEntityId('akt') };
    const fields = mapAktivitaetToFields(aktWithOwn, {
        teacherLabels: opts && opts.teacherLabels,
        accountEmail: opts && opts.accountEmail,
        teachers: opts && opts.teachers,
        requireOwnership: true
    });
    const created = await G().graphJson(
        'POST',
        G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(ctx.lists.aktivitaeten.id) + '/items',
        tok,
        { fields },
        'v1.0'
    );
    const itemId = created && created.id != null ? String(created.id) : '';
    // Zweiter Schreibvorgang nur Ownership – falls Create Felder still verworfen hat
    if (itemId) {
        try {
            await G().graphJson(
                'PATCH',
                G().graphPathSite(ctx.siteId) +
                    '/lists/' +
                    encodeURIComponent(ctx.lists.aktivitaeten.id) +
                    '/items/' +
                    encodeURIComponent(itemId) +
                    '/fields',
                tok,
                {
                    BeantragtVon: ownership.beantragtVon,
                    LehrerEmail: ownership.lehrerEmail,
                    LehrerCode: ownership.lehrerCode
                },
                'v1.0'
            );
        } catch (e) {
            const msg = e && e.message ? String(e.message) : String(e);
            throw new Error(
                'Aktivität angelegt, aber BeantragtVon/LehrerEmail/LehrerCode konnten nicht geschrieben werden. IT: Spalten in „Schulaktivitaeten“ prüfen. Details: ' +
                    msg
            );
        }
        try {
            const full = await G().graphJson(
                'GET',
                G().graphPathSite(ctx.siteId) +
                    '/lists/' +
                    encodeURIComponent(ctx.lists.aktivitaeten.id) +
                    '/items/' +
                    encodeURIComponent(itemId) +
                    '?$select=id,createdBy&$expand=fields',
                tok,
                undefined,
                'v1.0'
            );
            const mapped = mapAktivitaetFromItem(full);
            if (
                !String(mapped.beantragtVon || '').trim() ||
                !String(mapped.lehrerEmail || '').trim() ||
                !String(mapped.lehrerCode || '').trim()
            ) {
                throw new Error(
                    'Ownership-Felder sind nach dem Speichern in SharePoint leer. Bitte als IT die Spalten BeantragtVon, LehrerEmail und LehrerCode in „Schulaktivitaeten“ anlegen (Text) und erneut einreichen.'
                );
            }
            return mergeCreatedAktivitaet(mapped, aktWithOwn);
        } catch (e) {
            if (e && /Ownership-Felder|Spalten/.test(String(e.message || e))) throw e;
            return mergeCreatedAktivitaet(mapAktivitaetFromItem(created), aktWithOwn);
        }
    }
    return mergeCreatedAktivitaet(mapAktivitaetFromItem(created), aktWithOwn);
}

export async function updateAktivitaetItem(ctx, itemId, akt, opts) {
    const tok = await token();
    await ensureOwnershipColumnsOrThrow(ctx, tok);
    const ownership = resolveAktOwnership(akt, {
        accountEmail: opts && opts.accountEmail,
        teachers: opts && opts.teachers
    });
    assertAktOwnership(ownership);
    const aktWithOwn = { ...akt, ...ownership };
    const fields = mapAktivitaetToFields(aktWithOwn, {
        includeId: true,
        teacherLabels: opts && opts.teacherLabels,
        accountEmail: opts && opts.accountEmail,
        teachers: opts && opts.teachers,
        requireOwnership: true
    });
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
    return { ...aktWithOwn, itemId: String(itemId) };
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
        MinVorlaufTage: Number(patch.minVorlaufTage) || 0,
        MaxGleichzeitigProKlasse: Number(patch.maxGleichzeitigProKlasse) || 1,
        GenehmigungPfad: String(patch.genehmigungPfad || 'direktion'),
        VorlaufModus: String(patch.vorlaufModus || 'sperre'),
        Aktiv: patch.aktiv !== false
    };
    if (Object.prototype.hasOwnProperty.call(patch, 'kostenLimits')) {
        const list = Array.isArray(patch.kostenLimits) ? patch.kostenLimits : [];
        fields.KostenLimitsJson = JSON.stringify(
            list.map((r) => ({
                typ: String((r && r.typ) || '').trim(),
                okBis: Number(r && r.okBis) || 0,
                warnBis: Number(r && r.warnBis) || 0
            }))
        );
    }
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
