/**
 * SharePoint-CRUD für sRDP-Anmeldungen.
 */
import { listTitleCandidates, listTitleForYear } from './srdp-anmeldung-schema.js';
import {
    buildItemTitle,
    buildPruefplanKurz,
    validateAnmeldungDraft,
    resolveSrdpOwnership,
    assertSrdpOwnership
} from './srdp-anmeldung-logic.js';

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

function fieldBool(fields, key) {
    const v = fields && fields[key];
    return v === true || v === 'true' || v === 1 || v === '1' || v === 'Ja' || v === 'ja';
}

export function mapAnmeldungFromItem(item) {
    const f = (item && item.fields) || {};
    const createdBy =
        (item && item.createdBy && item.createdBy.user && item.createdBy.user.email) ||
        fieldStr(f, 'Author') ||
        '';
    const schuelerEmail =
        fieldStr(f, 'SchuelerEmail') || String(createdBy || '').trim().toLowerCase();
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        title: fieldStr(f, 'Title'),
        klasse: fieldStr(f, 'Klasse'),
        nachname: fieldStr(f, 'Nachname'),
        vorname: fieldStr(f, 'Vorname'),
        titelDiplomarbeit: fieldStr(f, 'TitelDiplomarbeit'),
        titelABA: fieldStr(f, 'TitelABA'),
        betreuungslehrerDA: fieldStr(f, 'BetreuungslehrerDA'),
        betreuungslehrerABA: fieldStr(f, 'BetreuungslehrerABA'),
        variante: fieldStr(f, 'Variante'),
        pruefplanKurz: fieldStr(f, 'PruefplanKurz'),
        lfs: fieldStr(f, 'LFS'),
        lehrerLFS: fieldStr(f, 'LehrerLFS'),
        lehrerBKO: fieldStr(f, 'LehrerBKO'),
        wahlfachMuendlich: fieldStr(f, 'WahlfachMuendlich'),
        seminar: fieldStr(f, 'Seminar'),
        lehrerWahlfach: fieldStr(f, 'LehrerWahlfach'),
        lehrerSchwerpunkt: fieldStr(f, 'LehrerSchwerpunkt'),
        lehrerFachkolloquium: fieldStr(f, 'LehrerFachkolloquium'),
        lehrerFachtheorie: fieldStr(f, 'LehrerFachtheorie'),
        lehrerVierteKlausur: fieldStr(f, 'LehrerVierteKlausur'),
        lehrerMuendlich: fieldStr(f, 'LehrerMuendlich'),
        bestaetigung: fieldBool(f, 'Bestaetigung'),
        nacharbeitErledigt: fieldBool(f, 'NacharbeitErledigt'),
        schuelerEmail: String(schuelerEmail || '').trim().toLowerCase(),
        beantragtVon: String(schuelerEmail || createdBy || '')
            .trim()
            .toLowerCase(),
        schulform: fieldStr(f, 'Schulform'),
        terminJahr: fieldStr(f, 'TerminJahr'),
        createdBy: String(createdBy || '').trim().toLowerCase(),
        created: (item && item.createdDateTime) || fieldStr(f, 'Created') || '',
        webUrl: (item && item.webUrl) || ''
    };
}

/**
 * @param {string} webUrl
 * @param {{ profileId?: string, jahr?: string|number }} [opts]
 */
export async function resolveSrdpContext(webUrl, opts) {
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die SharePoint-Website-URL eintragen.');
    const o = opts || {};
    const jahr = String(o.jahr || new Date().getFullYear()).trim();
    const profileId = o.profileId || 'hak';
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    const candidates = listTitleCandidates(jahr, profileId);
    let list = null;
    for (let i = 0; i < candidates.length; i++) {
        list = await findListByTitle(tok, siteId, candidates[i]);
        if (list) break;
    }
    if (!list) {
        const expected = listTitleForYear(jahr, profileId);
        throw new Error(
            'Liste „' +
                expected +
                '“ fehlt. Bitte zuerst unter Einrichtung die Liste anlegen.'
        );
    }
    return {
        webUrl: url,
        siteId,
        siteName: site.displayName || '',
        list: { id: String(list.id), webUrl: list.webUrl || '', title: list.displayName || '' }
    };
}

export async function loadAllAnmeldungen(ctx) {
    const tok = await token();
    const rows = await fetchAllItems(tok, ctx.siteId, ctx.list.id);
    return rows.map(mapAnmeldungFromItem);
}

/**
 * @param {object} ctx
 * @param {object} draft
 * @param {{ profileId?: string, jahr?: string, accountEmail?: string }} [meta]
 */
export async function createAnmeldungItem(ctx, draft, meta) {
    const m = meta || {};
    const profileId = m.profileId || 'hak';
    const ownership = resolveSrdpOwnership(draft, {
        accountEmail: m.accountEmail,
        students: m.students
    });
    assertSrdpOwnership(ownership);
    const validation = validateAnmeldungDraft({
        profileId,
        variante: draft.variante,
        nachname: draft.nachname || ownership.nachname,
        vorname: draft.vorname || ownership.vorname,
        wahlfach: draft.wahlfachMuendlich || draft.wahlfach,
        seminar: draft.seminar,
        lfs: draft.lfs,
        lehrerLfs: draft.lehrerLFS,
        bestaetigung: draft.bestaetigung
    });
    if (!validation.ok) {
        const err = new Error(validation.errors.join(' '));
        err.validation = validation;
        throw err;
    }
    const title = validation.title || buildItemTitle(draft.nachname, draft.vorname);
    const pruefplan = validation.pruefplanKurz || buildPruefplanKurz(draft.variante, profileId);
    const fields = {
        Title: title,
        Klasse: String(draft.klasse || ownership.klasse || '').trim(),
        Nachname: String(draft.nachname || ownership.nachname || '').trim(),
        Vorname: String(draft.vorname || ownership.vorname || '').trim(),
        TitelDiplomarbeit: String(draft.titelDiplomarbeit || '').trim(),
        BetreuungslehrerDA: String(draft.betreuungslehrerDA || '').trim(),
        Variante: String(draft.variante || '').trim(),
        PruefplanKurz: pruefplan,
        LFS: String(draft.lfs || '').trim(),
        LehrerLFS: String(draft.lehrerLFS || '').trim(),
        LehrerBKO: String(draft.lehrerBKO || '').trim(),
        WahlfachMuendlich: String(draft.wahlfachMuendlich || '').trim(),
        Seminar: String(draft.seminar || '').trim(),
        LehrerWahlfach: String(draft.lehrerWahlfach || '').trim(),
        Bestaetigung: true,
        SchuelerEmail: ownership.schuelerEmail,
        NacharbeitErledigt: false,
        Schulform: String(profileId || 'hak').toUpperCase(),
        TerminJahr: String(m.jahr || '').trim()
    };
    // leere Choice-Felder weglassen
    Object.keys(fields).forEach((k) => {
        if (fields[k] === '' || fields[k] == null) delete fields[k];
    });
    fields.Title = title;
    fields.Bestaetigung = true;
    fields.SchuelerEmail = ownership.schuelerEmail;
    fields.NacharbeitErledigt = false;

    const tok = await token();
    const path = G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(ctx.list.id) + '/items';
    let created;
    try {
        created = await G().graphJson('POST', path, tok, { fields }, 'v1.0');
    } catch (e) {
        // ältere Listen ohne Ownership-/Nacharbeit-Spalten: Fallback ohne diese Felder
        const msg = String((e && e.message) || e || '');
        if (/SchuelerEmail|NacharbeitErledigt/i.test(msg)) {
            delete fields.SchuelerEmail;
            delete fields.NacharbeitErledigt;
            created = await G().graphJson('POST', path, tok, { fields }, 'v1.0');
        } else {
            throw e;
        }
    }
    const mapped = mapAnmeldungFromItem(created);
    mapped.schuelerEmail = ownership.schuelerEmail;
    mapped.beantragtVon = ownership.beantragtVon;
    mapped.createdBy = ownership.schuelerEmail;
    mapped.nacharbeitErledigt = false;
    return mapped;
}

/**
 * Nacharbeit (Prüfungsliste) in SharePoint aktualisieren.
 * @param {object} ctx
 * @param {string} itemId SPO list item id
 * @param {boolean} nacharbeitErledigt
 */
export async function updateAnmeldungNacharbeit(ctx, itemId, nacharbeitErledigt) {
    const spoId = String(itemId || '').trim();
    if (!spoId) throw new Error('itemId fehlt – Anmeldung noch nicht in SharePoint.');
    if (/^demo-/i.test(spoId)) {
        throw new Error('Demo-Items werden nicht nach SharePoint geschrieben.');
    }
    const tok = await token();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.list.id) +
        '/items/' +
        encodeURIComponent(spoId) +
        '/fields';
    const fields = { NacharbeitErledigt: !!nacharbeitErledigt };
    try {
        await G().graphJson('PATCH', path, tok, fields, 'v1.0');
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        if (/NacharbeitErledigt/i.test(msg)) {
            throw new Error(
                'Spalte „NacharbeitErledigt“ fehlt in der Liste. Bitte Setup (fehlende Spalten) ausführen.'
            );
        }
        throw e;
    }
    return { itemId: spoId, nacharbeitErledigt: !!nacharbeitErledigt };
}
