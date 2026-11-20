/**
 * SharePoint-CRUD für sRDP-Anmeldungen.
 */
import { listTitleCandidates, listTitleForYear } from './srdp-anmeldung-schema.js';
import { bfkIsoFromColumn } from './srdp-pruefungstermine.js';
import {
    buildItemTitle,
    buildPruefplanKurz,
    validateAnmeldungDraft,
    resolveSrdpOwnership,
    assertSrdpOwnership,
    preferredSrdpTerminJahr
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

const freshTokenByScope = new Map();

/**
 * Beim Laden einmal ein frisches Token holen, damit neue Gruppenmitgliedschaften
 * ohne Abmelden gelten. Danach derselbe Token für diese Berechtigungen.
 * @param {string[]} [scopes]
 */
async function token(scopes) {
    const sc = Array.isArray(scopes) && scopes.length ? scopes : SCOPES;
    const key = sc.join(' ');
    if (freshTokenByScope.has(key)) return freshTokenByScope.get(key);
    let tok = '';
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthAcquireTokenFresh === 'function') {
            tok = await window.ms365AuthAcquireTokenFresh(sc);
        }
    } catch {
        tok = '';
    }
    if (!tok) tok = await G().getGraphToken(sc);
    freshTokenByScope.set(key, tok);
    return tok;
}

function listNameKey(s) {
    return String(s || '')
        .trim()
        .toLowerCase()
        .replace(/[\s_\-]+/g, '');
}

/**
 * Liste finden – $filter allein reicht oft nicht (Schüler mit Unique Permissions /
 * Graph-Filter auf displayName liefert leeres value). Fallback: Direktzugriff + Enumeration.
 * @param {string} tok
 * @param {string} siteId
 * @param {string} listTitle
 * @returns {Promise<object|null>}
 */
async function findListByTitle(tok, siteId, listTitle) {
    const title = String(listTitle || '').trim();
    if (!title) return null;
    const want = title.toLowerCase();
    const wantKey = listNameKey(title);

    try {
        const path =
            G().graphPathSite(siteId) +
            '/lists?$filter=' +
            encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
            '&$select=id,displayName,webUrl,name';
        const data = await G().graphJson('GET', path, tok, undefined, 'v1.0');
        const hit = ((data && data.value) || [])[0] || null;
        if (hit && hit.id) return hit;
    } catch {
        /* Filter oft leer/unsupported – weiter mit Fallbacks */
    }

    try {
        const direct = await G().graphJson(
            'GET',
            G().graphPathSite(siteId) +
                '/lists/' +
                encodeURIComponent(title) +
                '?$select=id,displayName,webUrl,name',
            tok,
            undefined,
            'v1.0'
        );
        if (direct && direct.id) return direct;
    } catch {
        /* DisplayName-Pfad nicht lesbar */
    }

    try {
        let path = G().graphPathSite(siteId) + '/lists?$select=id,displayName,webUrl,name&$top=200';
        while (path) {
            const data = await G().graphJson(
                'GET',
                path.indexOf('http') === 0 ? path : path,
                tok,
                undefined,
                'v1.0'
            );
            const batch = (data && data.value) || [];
            for (let i = 0; i < batch.length; i++) {
                const row = batch[i];
                const dn = String(row.displayName || '').trim().toLowerCase();
                const nm = String(row.name || '').trim().toLowerCase();
                if (dn === want || nm === want) return row;
                if (listNameKey(row.displayName) === wantKey || listNameKey(row.name) === wantKey) {
                    return row;
                }
            }
            path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
        }
    } catch {
        /* Enumeration für Kandidat:innen oft leer oder 403 – SPO-getbytitle übernimmt */
    }
    return null;
}

/**
 * Direkt per SharePoint-REST. Graph listet Listen mit eigener Berechtigung für Schüler oft nicht.
 * @param {string} webUrl
 * @param {string[]} titles
 * @returns {Promise<{ list: object|null, status: number }>}
 */
async function findListViaSpoRest(webUrl, titles) {
    const url = String(webUrl || '').trim().replace(/\/$/, '');
    const names = (Array.isArray(titles) ? titles : []).map((t) => String(t || '').trim()).filter(Boolean);
    if (!url || !names.length) return { list: null, status: 0 };
    let host = '';
    try {
        host = new URL(url).hostname;
    } catch {
        return { list: null, status: 0 };
    }
    let spoTok = '';
    try {
        spoTok = await token([
            'https://' + host + '/AllSites.Read',
            'https://graph.microsoft.com/User.Read'
        ]);
    } catch {
        return { list: null, status: 0 };
    }
    let digest = '';
    try {
        digest = await G().getSpoRequestDigest(url, spoTok);
    } catch {
        digest = '';
    }
    let lastStatus = 0;
    for (let i = 0; i < names.length; i++) {
        const title = names[i];
        const api =
            "/_api/web/lists/getbytitle('" +
            title.replace(/'/g, "''") +
            "')?$select=Id,Title,DefaultViewUrl";
        let res;
        try {
            res = await G().spoRestFetch(url, spoTok, digest, 'GET', api);
        } catch {
            continue;
        }
        lastStatus = res && res.status ? res.status : 0;
        if (!res || !res.ok) continue;
        const d = (res.data && res.data.d) || res.data || {};
        const id = d.Id || d.id || d.ID;
        if (!id) continue;
        return {
            list: {
                id: String(id),
                displayName: d.Title || title,
                webUrl: d.DefaultViewUrl || ''
            },
            status: lastStatus
        };
    }
    return { list: null, status: lastStatus };
}

/**
 * Klassen-Choices über SharePoint-REST, wenn Graph die Spalten nicht liefert.
 * @param {string} webUrl
 * @param {string} listId
 * @returns {Promise<string[]>}
 */
async function spoTokenForSite(webUrl) {
    const url = String(webUrl || '').trim().replace(/\/$/, '');
    let host = '';
    try {
        host = new URL(url).hostname;
    } catch {
        return { url: '', spoTok: '', digest: '' };
    }
    if (!url || !host) return { url: '', spoTok: '', digest: '' };
    let spoTok = '';
    try {
        spoTok = await token([
            'https://' + host + '/AllSites.Read',
            'https://graph.microsoft.com/User.Read'
        ]);
    } catch {
        return { url, spoTok: '', digest: '' };
    }
    let digest = '';
    try {
        digest = await G().getSpoRequestDigest(url, spoTok);
    } catch {
        digest = '';
    }
    return { url, spoTok, digest };
}

function choiceValuesFromSpoField(data) {
    const d = (data && data.d) || data || {};
    const ch = d.Choices;
    const raw = Array.isArray(ch) ? ch : ch && Array.isArray(ch.results) ? ch.results : [];
    return raw.map((x) => String(x || '').trim()).filter((x) => x && x !== '—' && x !== '–');
}

async function loadSrdpFieldChoicesViaSpo(webUrl, listId, internalName) {
    const id = String(listId || '').trim();
    const field = String(internalName || '').trim();
    const ctx = await spoTokenForSite(webUrl);
    if (!ctx.url || !ctx.spoTok || !id || !field) return [];
    const api =
        "/_api/web/lists(guid'" +
        id.replace(/'/g, "''") +
        "')/fields/getbyinternalnameortitle('" +
        field.replace(/'/g, "''") +
        "')?$select=InternalName,Choices";
    let res;
    try {
        res = await G().spoRestFetch(ctx.url, ctx.spoTok, ctx.digest, 'GET', api);
    } catch {
        return [];
    }
    if (!res || !res.ok) return [];
    return choiceValuesFromSpoField(res.data);
}

/**
 * Klassen-Choices über SharePoint-REST, wenn Graph die Spalten nicht liefert.
 * @param {string} webUrl
 * @param {string} listId
 * @returns {Promise<string[]>}
 */
export async function loadSrdpKlasseChoicesViaSpo(webUrl, listId) {
    return loadSrdpFieldChoicesViaSpo(webUrl, listId, 'Klasse');
}

/**
 * Lehrkraft-Namen aus den Choice-Spalten der Anmeldeliste (Schüler dürfen diese Liste lesen).
 * @param {string} webUrl
 * @param {string} listId
 * @returns {Promise<string[]>}
 */
export async function loadSrdpLehrerChoicesViaSpo(webUrl, listId) {
    const fields = ['BetreuungslehrerDA', 'LehrerLFS', 'LehrerBKO', 'LehrerWahlfach'];
    const out = [];
    const seen = new Set();
    for (let i = 0; i < fields.length; i++) {
        const rows = await loadSrdpFieldChoicesViaSpo(webUrl, listId, fields[i]);
        rows.forEach((label) => {
            const key = label.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            out.push(label);
        });
    }
    out.sort((a, b) => a.localeCompare(b, 'de'));
    return out;
}

async function getListById(tok, siteId, listId) {
    const id = String(listId || '').trim();
    if (!id) return null;
    try {
        const list = await G().graphJson(
            'GET',
            G().graphPathSite(siteId) +
                '/lists/' +
                encodeURIComponent(id) +
                '?$select=id,displayName,webUrl,name',
            tok,
            undefined,
            'v1.0'
        );
        return list && list.id ? list : null;
    } catch {
        return null;
    }
}

/**
 * Choice-Werte aus der Live-Liste (Klasse / LFS / Wahlfach / Seminar) –
 * zuverlässig für Schüler ohne IT-Wizard in localStorage.
 * @param {string} tok
 * @param {string} siteId
 * @param {string} listId
 */
export async function loadSrdpListChoiceOptions(tok, siteId, listId) {
    const sid = String(siteId || '').trim();
    const lid = String(listId || '').trim();
    if (!sid || !lid) return {};
    const path =
        G().graphPathSite(sid) +
        '/lists/' +
        encodeURIComponent(lid) +
        '/columns?$select=name,displayName,choice,description,defaultValue&$top=200';
    let data;
    try {
        data = await G().graphJson('GET', path, tok, undefined, 'v1.0');
    } catch {
        data = await G().graphJson(
            'GET',
            path.replace(',description,defaultValue', ''),
            tok,
            undefined,
            'v1.0'
        );
    }
    const cols = (data && data.value) || [];
    const pick = (internalName) => {
        const col = cols.find((c) => String(c.name || '').trim() === internalName);
        const choices = col && col.choice && Array.isArray(col.choice.choices) ? col.choice.choices : [];
        return choices.map((x) => String(x || '').trim()).filter((x) => x && x !== '—' && x !== '–');
    };
    const teacherFields = [
        'BetreuungslehrerDA',
        'BetreuungslehrerABA',
        'LehrerLFS',
        'LehrerBKO',
        'LehrerWahlfach',
        'LehrerSchwerpunkt',
        'LehrerFachkolloquium',
        'LehrerFachtheorie',
        'LehrerVierteKlausur',
        'LehrerMuendlich'
    ];
    const lehrer = [];
    const seenLehrer = new Set();
    teacherFields.forEach((name) => {
        pick(name).forEach((label) => {
            const key = label.toLowerCase();
            if (seenLehrer.has(key)) return;
            seenLehrer.add(key);
            lehrer.push(label);
        });
    });
    lehrer.sort((a, b) => a.localeCompare(b, 'de'));
    return {
        klassen: pick('Klasse'),
        lfs: pick('LFS'),
        wahlfaecher: pick('WahlfachMuendlich'),
        seminare: pick('Seminar'),
        lehrer,
        bfkDatum: bfkIsoFromColumn(cols.find((c) => String(c.name || '').trim() === 'BfkSchriftlich')),
        bfkKnown: cols.some((c) => String(c.name || '').trim() === 'BfkSchriftlich')
    };
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

function yearOfListTitle(title) {
    const m = String(title || '').match(/\b(20\d{2})\b/);
    return m ? m[1] : '';
}

/**
 * @param {string} tok
 * @param {string} siteId
 * @param {string} webUrl
 * @param {string} jahr
 * @param {string} profileId
 */
async function findListForYear(tok, siteId, webUrl, jahr, profileId) {
    const candidates = listTitleCandidates(jahr, profileId);
    for (let i = 0; i < candidates.length; i++) {
        const hit = await findListByTitle(tok, siteId, candidates[i]);
        if (hit) return { list: hit, spoStatus: 0 };
    }
    const spo = await findListViaSpoRest(webUrl, candidates);
    return {
        list: spo && spo.list ? spo.list : null,
        spoStatus: spo && spo.status ? spo.status : 0
    };
}

/**
 * @param {string} webUrl
 * @param {{ profileId?: string, jahr?: string|number, listId?: string }} [opts]
 */
export async function resolveSrdpContext(webUrl, opts) {
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die SharePoint-Website-URL eintragen.');
    const o = opts || {};
    const savedJahr = String(o.jahr || '').trim();
    const preferred = preferredSrdpTerminJahr(new Date(), savedJahr);
    const profileId = o.profileId || 'hak';
    const hintId = String(o.listId || '').trim();
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    const hinted = hintId ? await getListById(tok, siteId, hintId) : null;
    const hintedYear = yearOfListTitle(hinted && (hinted.displayName || hinted.title));
    let list = hinted && hintedYear === preferred ? hinted : null;
    let spoStatus = 0;
    if (!list) {
        const found = await findListForYear(tok, siteId, url, preferred, profileId);
        list = found.list;
        spoStatus = found.spoStatus;
    }
    if (!list) {
        const expected = listTitleForYear(preferred, profileId);
        const hidden = new Error(
            'Liste „' +
                expected +
                '“ ist für dieses Konto noch nicht sichtbar. Zugriff haben die Kandidaten-Gruppen (Klassen), Lehrkräfte und die Direktion. Seite neu laden – die Mitgliedschaft wird dabei neu geholt, ohne Abmelden.'
        );
        hidden.code = spoStatus === 403 || spoStatus === 401 ? 'srdp-list-denied' : 'srdp-list-hidden';
        throw hidden;
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
 * @param {{ profileId?: string, jahr?: string, accountEmail?: string, accountName?: string, students?: object[] }} [meta]
 */
export async function createAnmeldungItem(ctx, draft, meta) {
    const m = meta || {};
    const profileId = m.profileId || 'hak';
    const ownership = resolveSrdpOwnership(draft, {
        accountEmail: m.accountEmail,
        accountName: m.accountName,
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

/**
 * Gruppen, in denen das angemeldete Konto Mitglied ist.
 * Reicht für „DEMO Klasse 5A“, ohne die Schülerliste im Browser des Schülers.
 * @returns {Promise<{ id: string, displayName: string, mail: string }[]>}
 */
export async function fetchMyMemberGroups() {
    let tok = '';
    try {
        tok = await token(['https://graph.microsoft.com/User.Read']);
    } catch {
        return [];
    }
    if (!tok) return [];
    try {
        const data = await G().graphJson(
            'GET',
            '/me/memberOf/microsoft.graph.group?$select=id,displayName,mail&$top=999',
            tok,
            undefined,
            'v1.0'
        );
        return ((data && data.value) || [])
            .map((g) => ({
                id: String((g && g.id) || '').trim(),
                displayName: String((g && g.displayName) || '').trim(),
                mail: String((g && g.mail) || '')
                    .trim()
                    .toLowerCase()
            }))
            .filter((g) => g.id || g.displayName);
    } catch {
        return [];
    }
}
