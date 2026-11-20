/**
 * SharePoint-CRUD: Lehrer-Freistellungen.
 */
import { LIST_TITLE_DEFAULT, newAntragId, STATUS_CHOICES } from './lfr-schema.js';
import {
    toIsoDateOnly,
    toSharePointDateTime,
    normalizeStatus
} from './lfr-logic.js';
import { loadSetupCfg } from './lfr-state.js';

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
    // Nur fields expandieren. $select=createdBy lässt die Spalten oft leer –
    // dann filtert die App eigene Anträge weg, obwohl die Liste sie enthält.
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

/** SharePoint-Antwort (REST, GetItems, RenderListDataAsStream) → Zeilen. */
export function extractSpoItemRows(data) {
    if (!data) return [];
    if (Array.isArray(data)) return data;
    if (Array.isArray(data.Row)) return data.Row;
    const d = data.d || data;
    if (Array.isArray(d)) return d;
    if (Array.isArray(d.Row)) return d.Row;
    if (Array.isArray(d.results)) return d.results;
    if (Array.isArray(d.value)) return d.value;
    return [];
}

/**
 * SharePoint-Zeile → Graph-ähnliches Item.
 * callerVisible: REST läuft mit dem Benutzer-Token und „nur eigene Elemente“.
 * @param {object} row
 */
export function mapRestListItem(row) {
    const r = row || {};
    const author = r.Author || {};
    const email =
        r.LehrerEmail ||
        r.EMailLehrkraft ||
        author.EMail ||
        author.Email ||
        author.UserPrincipalName ||
        '';
    return {
        id: r.Id != null ? r.Id : r.ID != null ? r.ID : r.id,
        callerVisible: true,
        createdBy: {
            user: {
                email: author.EMail || author.Email || email || '',
                userPrincipalName: author.UserPrincipalName || '',
                displayName: author.Title || r.LehrerName || ''
            }
        },
        fields: {
            Title: r.Title || r.Titel,
            Beginn: r.Beginn,
            Ende: r.Ende,
            Status: r.Status,
            Kategorie: r.Kategorie,
            Beschreibung: r.Beschreibung,
            LehrerName: r.LehrerName || r.Lehrkraft,
            LehrerEmail: email,
            AntragId: r.AntragId || r.AntragID,
            GenehmigtVonDirektion: r.GenehmigtVonDirektion,
            GenehmigtAmDirektion: r.GenehmigtAmDirektion,
            AbgelehntVon: r.AbgelehntVon,
            AbgelehntAm: r.AbgelehntAm,
            BemerkungDirektion: r.BemerkungDirektion
        }
    };
}

function listApiPath(listId, listTitle) {
    const id = String(listId || '').trim();
    if (id) return "/_api/web/lists(guid'" + id.replace(/'/g, "''") + "')";
    const title = String(listTitle || 'Lehrer-Freistellungen').trim() || 'Lehrer-Freistellungen';
    return "/_api/web/lists/getbytitle('" + title.replace(/'/g, "''") + "')";
}

/**
 * SPO-REST: dieselbe Sicht wie die SharePoint-Liste (auch bei „nur eigene Elemente“).
 * Graph liefert in dem Fall oft 0 Zeilen.
 * @param {string} webUrl
 * @param {string} listId
 * @param {string} [listTitle]
 */
async function fetchAllItemsViaSpoRest(webUrl, listId, listTitle) {
    const url = String(webUrl || '')
        .trim()
        .replace(/\/+$/, '');
    if (!url) return [];
    let host = '';
    try {
        host = new URL(url).hostname;
    } catch {
        return [];
    }
    const spoTok = await G().getGraphToken([
        'https://' + host + '/AllSites.Read',
        'https://graph.microsoft.com/User.Read'
    ]);
    const digest = await G().getSpoRequestDigest(url, spoTok);
    const paths = [];
    const byId = listApiPath(listId, '');
    const byTitle = listApiPath('', listTitle || 'Lehrer-Freistellungen');
    if (listId) paths.push(byId);
    if (byTitle !== byId) paths.push(byTitle);
    /** @type {string[]} */
    const errors = [];
    for (let i = 0; i < paths.length; i++) {
        const base = paths[i];
        const attempts = [
            {
                method: 'POST',
                api: base + '/RenderListDataAsStream',
                body: {
                    parameters: {
                        RenderOptions: 2,
                        ViewXml: '<View><RowLimit>300</RowLimit></View>'
                    }
                }
            },
            {
                method: 'GET',
                api: base + '/items?$expand=FieldValuesAsText&$top=200',
                body: undefined
            }
        ];
        for (let a = 0; a < attempts.length; a++) {
            const attempt = attempts[a];
            let res;
            try {
                res = await G().spoRestFetch(
                    url,
                    spoTok,
                    digest,
                    attempt.method,
                    attempt.api,
                    attempt.body
                );
            } catch (e) {
                errors.push(e && e.message ? e.message : String(e));
                continue;
            }
            if (!res || !res.ok) {
                errors.push(
                    attempt.method +
                        ' ' +
                        (res && res.status ? res.status : '?') +
                        ' ' +
                        String((res && res.text) || '').slice(0, 180)
                );
                continue;
            }
            const rows = extractSpoItemRows(res.data);
            if (rows.length) {
                return rows.map((row) => {
                    const text = row && (row.FieldValuesAsText || row.fieldValuesAsText);
                    if (text && typeof text === 'object') {
                        return mapRestListItem(
                            Object.assign({}, text, {
                                Id: row.Id || row.ID || text.Id || text.ID
                            })
                        );
                    }
                    return mapRestListItem(row);
                });
            }
        }
    }
    if (errors.length) {
        const err = new Error(errors[0]);
        err.details = errors;
        throw err;
    }
    return [];
}

function fieldStr(fields, key) {
    const v = fields && fields[key];
    if (v == null) return '';
    if (typeof v === 'object') {
        const mail = v.Email || v.EMail || v.email || v.UserPrincipalName || '';
        return String(mail || '').trim();
    }
    const s = String(v).trim();
    const m = /^\/Date\((-?\d+)\)\/$/.exec(s);
    if (m) {
        const d = new Date(Number(m[1]));
        if (!isNaN(d.getTime())) return d.toISOString();
    }
    const de = /^(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{2}):(\d{2}))?/.exec(s);
    if (de) {
        return de[3] + '-' + de[2] + '-' + de[1] + (de[4] ? 'T' + de[4] + ':' + de[5] : '');
    }
    return s;
}

/** Case-insensitive Feldzugriff (Graph liefert manchmal abweichende Keys). */
function fieldStrCI(fields, keys) {
    const list = Array.isArray(keys) ? keys : [keys];
    for (let i = 0; i < list.length; i++) {
        const hit = fieldStr(fields, list[i]);
        if (hit) return hit;
    }
    const map = {};
    Object.keys(fields || {}).forEach((k) => {
        map[String(k).toLowerCase()] = k;
    });
    for (let i = 0; i < list.length; i++) {
        const real = map[String(list[i] || '').toLowerCase()];
        if (real) {
            const hit = fieldStr(fields, real);
            if (hit) return hit;
        }
    }
    return '';
}

export function authorEmailFromLfrItem(item) {
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

/**
 * @param {string} webUrl
 * @param {{ listName?: string, listId?: string }} [opts]
 */
export async function resolveLfrContext(webUrl, opts) {
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die SharePoint-Website-URL eintragen.');
    const setup = loadSetupCfg();
    const options = opts || {};
    const listName =
        String(options.listName || setup.listName || LIST_TITLE_DEFAULT).trim() ||
        LIST_TITLE_DEFAULT;
    const hintId = String(options.listId || setup.listId || '').trim();

    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    let list = null;
    if (hintId) {
        try {
            list = await G().graphJson(
                'GET',
                G().graphPathSite(siteId) +
                    '/lists/' +
                    encodeURIComponent(hintId) +
                    '?$select=id,displayName,webUrl',
                tok,
                undefined,
                'v1.0'
            );
        } catch {
            list = null;
        }
    }
    if (!list || !list.id) list = await findListByTitle(tok, siteId, listName);
    if (!list || !list.id) {
        throw new Error(
            'Liste „' +
                listName +
                '“ fehlt. Unter „Einrichtung“ anlegen oder IT-Setup (demnächst) nutzen.'
        );
    }
    return {
        webUrl: url.replace(/\/$/, ''),
        siteId,
        siteName: site.displayName || '',
        list: { id: String(list.id), title: list.displayName || listName, webUrl: list.webUrl || '' }
    };
}

export function mapItemFromSp(item) {
    const f = (item && item.fields) || {};
    const authorEmail = authorEmailFromLfrItem(item);
    const lehrerEmail =
        fieldStrCI(f, ['LehrerEmail', 'EMailLehrkraft', 'EmailLehrkraft']).toLowerCase() ||
        authorEmail;
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        antragId: fieldStrCI(f, ['AntragId', 'AntragID']),
        titel: fieldStrCI(f, ['Title', 'Titel']),
        beginn: fieldStrCI(f, ['Beginn', 'Start']) || fieldStr(f, 'Start'),
        ende: fieldStrCI(f, ['Ende']),
        status: normalizeStatus(fieldStrCI(f, ['Status']) || 'Ausstehend'),
        kategorie: fieldStrCI(f, ['Kategorie']),
        beschreibung: fieldStrCI(f, ['Beschreibung']),
        lehrerName: fieldStrCI(f, ['LehrerName', 'Lehrkraft']) || '',
        lehrerEmail,
        authorEmail,
        callerVisible: !!(item && item.callerVisible),
        genehmigtVon: fieldStrCI(f, ['GenehmigtVonDirektion']),
        genehmigtAm: fieldStrCI(f, ['GenehmigtAmDirektion']),
        abgelehntVon: fieldStrCI(f, ['AbgelehntVon']),
        abgelehntAm: fieldStrCI(f, ['AbgelehntAm']),
        bemerkungDirektion: fieldStrCI(f, ['BemerkungDirektion'])
    };
}

export function mapItemToFields(row, opts) {
    const includeId = !opts || opts.includeId !== false;
    const fields = {
        Title: String(row.titel || '').trim() || 'Freistellung Lehrkraft',
        Beginn: toSharePointDateTime(row.beginn),
        Ende: toSharePointDateTime(row.ende),
        Status: normalizeStatus(row.status || 'Ausstehend'),
        Kategorie: String(row.kategorie || '').trim(),
        Beschreibung: String(row.beschreibung || ''),
        LehrerName: String(row.lehrerName || '').trim(),
        LehrerEmail: String(row.lehrerEmail || '').trim().toLowerCase(),
        BemerkungDirektion: String(row.bemerkungDirektion || '')
    };
    if (row.genehmigtVon) fields.GenehmigtVonDirektion = row.genehmigtVon;
    if (row.genehmigtAm) fields.GenehmigtAmDirektion = toIsoDateOnly(row.genehmigtAm);
    if (row.abgelehntVon) fields.AbgelehntVon = row.abgelehntVon;
    if (row.abgelehntAm) fields.AbgelehntAm = toIsoDateOnly(row.abgelehntAm);
    if (includeId) {
        fields.AntragId = String(row.antragId || '').trim() || newAntragId();
    }
    return fields;
}

function rowHasPayload(row) {
    const f = (row && row.fields) || {};
    return !!(
        fieldStr(f, 'Title') ||
        fieldStr(f, 'LehrerEmail') ||
        fieldStr(f, 'LehrerName') ||
        fieldStr(f, 'Beginn')
    );
}

export async function loadAllItems(ctx) {
    // Zuerst die SharePoint-Ansicht des angemeldeten Kontos.
    // Nach „nur eigene Elemente“ ist das die Liste, die die Lehrkraft im Browser sieht.
    // Graph (Sites.ReadWrite.All) liefert dann oft 0 Zeilen oder zusätzlich fremde.
    let restError = '';
    if (ctx && ctx.webUrl) {
        try {
            const restRows = await fetchAllItemsViaSpoRest(
                ctx.webUrl,
                ctx.list && ctx.list.id,
                ctx.list && ctx.list.title
            );
            if (restRows.length) return restRows.map(mapItemFromSp);
        } catch (e) {
            restError = e && e.message ? e.message : String(e);
        }
    }
    const tok = await token();
    let graphRows = [];
    try {
        graphRows = await fetchAllItems(tok, ctx.siteId, ctx.list.id);
    } catch (e) {
        if (!restError) restError = e && e.message ? e.message : String(e);
        graphRows = [];
    }
    const graphOk = graphRows.filter(rowHasPayload);
    if (graphOk.length) return graphOk.map(mapItemFromSp);
    if (restError) throw new Error(restError);
    return [];
}

export async function createItem(ctx, row) {
    const tok = await token();
    const fields = mapItemToFields({
        ...row,
        antragId: row.antragId || newAntragId(),
        status: row.status || 'Ausstehend'
    });
    const created = await G().graphJson(
        'POST',
        G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(ctx.list.id) + '/items',
        tok,
        { fields },
        'v1.0'
    );
    return mapItemFromSp(created);
}

export async function updateItem(ctx, itemId, row) {
    const tok = await token();
    const fields = mapItemToFields(row, { includeId: true });
    await G().graphJson(
        'PATCH',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.list.id) +
            '/items/' +
            encodeURIComponent(itemId) +
            '/fields',
        tok,
        fields,
        'v1.0'
    );
    return row;
}

/** Spalten anlegen (Einrichtung). */
export async function ensureListColumns(ctx, columnDefs, write) {
    const tok = await token();
    const { toGraphColumnBody } = await import('../freistellung-planer/freistellung-planer-schema.js');
    const toBody = toGraphColumnBody || ((col) => col);
    const listId = ctx.list.id;
    const existing = await G().graphJson(
        'GET',
        G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(listId) + '/columns',
        tok,
        undefined,
        'v1.0'
    );
    const names = new Set(
        ((existing && existing.value) || []).map((c) => String(c.name || '').toLowerCase())
    );
    for (const col of columnDefs || []) {
        const n = String(col.name || '').toLowerCase();
        if (!n || names.has(n)) continue;
        if (write) write('Spalte: ' + col.displayName);
        await G().graphJson(
            'POST',
            G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(listId) + '/columns',
            tok,
            toBody(col),
            'v1.0'
        );
        names.add(n);
    }
}

export async function createListIfMissing(ctx, listTitle, write) {
    const tok = await token();
    const title = String(listTitle || LIST_TITLE_DEFAULT).trim();
    let list = await findListByTitle(tok, ctx.siteId, title);
    if (list && list.id) {
        const next = { ...ctx, list: { id: String(list.id), title, webUrl: list.webUrl || '' } };
        const { LFR_COLUMNS } = await import('./lfr-schema.js');
        await ensureListColumns(next, LFR_COLUMNS, write);
        return next;
    }
    if (write) write('Liste anlegen: ' + title);
    const created = await G().graphJson(
        'POST',
        G().graphPathSite(ctx.siteId) + '/lists',
        tok,
        {
            displayName: title,
            list: { template: 'genericList' }
        },
        'v1.0'
    );
    const id = created && created.id ? String(created.id) : '';
    if (!id) throw new Error('Liste konnte nicht angelegt werden.');
    const next = {
        ...ctx,
        list: { id, title, webUrl: created.webUrl || '' }
    };
    const { LFR_COLUMNS } = await import('./lfr-schema.js');
    await ensureListColumns(next, LFR_COLUMNS, write);
    return next;
}

export { STATUS_CHOICES };
