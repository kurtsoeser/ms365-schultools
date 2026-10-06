/**
 * SharePoint-CRUD für Freistellungs-Planer.
 */
import { LIST_TITLE_DEFAULT, KATEGORIE_CHOICES } from './freistellung-planer-schema.js';
import { parseNachweiseField } from './freistellung-planer-nachweise.js';
import { mergeKategorieChoices } from './freistellung-planer-kategorien.js';
import { toIsoDateOnly, approvalPath, inclusiveDayCount } from './freistellung-planer-logic.js';
import { buildAntragTitle } from './freistellung-planer-state.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

const SCOPES_READ = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
];

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

async function token() {
    return await G().getGraphToken(SCOPES);
}

/** Lesen (Schüler/KV oft nur Sites.Read + Listen-contribute). */
async function tokenRead() {
    const scopes = [...SCOPES_READ, 'https://graph.microsoft.com/Sites.ReadWrite.All'];
    return await G().getGraphToken(scopes);
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

/**
 * Liste auf einer Site finden (ID, Filter oder Enumeration – für Schüler mit nur Listenrecht).
 * @param {string} webUrl
 * @param {{ listName?: string, listId?: string }} [opts]
 */
export async function findFreistellungListOnSite(webUrl, opts) {
    const url = String(webUrl || '').trim().replace(/\/$/, '');
    if (!url) return null;
    const options = opts || {};
    const listName = String(options.listName || LIST_TITLE_DEFAULT).trim() || LIST_TITLE_DEFAULT;
    const hintId = String(options.listId || '').trim();
    const tok = await tokenRead();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) return null;

    if (hintId) {
        try {
            const list = await G().graphJson(
                'GET',
                G().graphPathSite(siteId) +
                    '/lists/' +
                    encodeURIComponent(hintId) +
                    '?$select=id,displayName,webUrl',
                tok,
                undefined,
                'v1.0'
            );
            if (list && list.id) return list;
        } catch {
            /* direkte ID nicht lesbar */
        }
    }

    let list = await findListByTitle(tok, siteId, listName);
    if (list && list.id) return list;

    let path = G().graphPathSite(siteId) + '/lists?$select=id,displayName,webUrl&$top=200';
    const want = listName.toLowerCase();
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
            if (String(row.displayName || '').trim().toLowerCase() === want) return row;
        }
        path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
    }
    return null;
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

/**
 * Person-Feld aus Graph fields lesen.
 * @param {object} fields
 * @param {string} key
 */
export function readPersonField(fields, key) {
    const raw = fields && fields[key];
    if (!raw) {
        const lookupEmail = fieldStr(fields, key + 'Email');
        if (lookupEmail) return { email: lookupEmail.toLowerCase(), name: '', lookupId: '' };
        return { email: '', name: '', lookupId: '' };
    }
    const entry = Array.isArray(raw) ? raw[0] : raw;
    if (!entry || typeof entry !== 'object') {
        return { email: '', name: String(entry || '').trim(), lookupId: '' };
    }
    return {
        email: String(entry.Email || entry.email || '')
            .trim()
            .toLowerCase(),
        name: String(entry.LookupValue || entry.DisplayName || entry.Title || '').trim(),
        lookupId:
            entry.LookupId != null
                ? String(entry.LookupId)
                : entry.id != null
                  ? String(entry.id)
                  : ''
    };
}

/**
 * SharePoint User Information List → LookupId für Person-Spalte.
 * @param {string} tok
 * @param {string} siteId
 * @param {string} email
 */
export async function resolvePersonLookupId(tok, siteId, email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em || !em.includes('@')) return '';

    let userList = null;
    if (typeof G().findSiteUserInformationList === 'function') {
        userList = await G().findSiteUserInformationList(tok, siteId);
    } else {
        const listsPath =
            G().graphPathSite(siteId) + '/lists?$select=id,displayName,name,system&$top=999';
        const listsData = await G().graphJson('GET', listsPath, tok, undefined, 'v1.0');
        const lists = (listsData && listsData.value) || [];
        userList =
            lists.find((l) => String(l.name || '').trim().toLowerCase() === 'users') ||
            lists.find((l) => String(l.displayName || '') === 'User Information List') ||
            lists.find((l) => /user information/i.test(String(l.displayName || ''))) ||
            lists.find((l) => /benutzerinformationsliste/i.test(String(l.displayName || ''))) ||
            null;
    }
    if (!userList || !userList.id) {
        throw new Error(
            'User Information List nicht gefunden – Klassenvorstand kann nicht als Person gesetzt werden.'
        );
    }

    // EMail-Feld je nach Tenant-Locale unterschiedlich → breit filtern und clientseitig matchen
    let path =
        G().graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(userList.id) +
        '/items?$expand=fields&$top=200';
    while (path) {
        const data = await G().graphJson(
            'GET',
            path.indexOf('http') === 0 ? path : path,
            tok,
            undefined,
            'v1.0'
        );
        const rows = (data && data.value) || [];
        for (let i = 0; i < rows.length; i++) {
            const f = (rows[i] && rows[i].fields) || {};
            const candidates = [
                f.EMail,
                f.Email,
                f.UserName,
                f.SipAddress,
                f.Name
            ]
                .map((x) => String(x || '').trim().toLowerCase())
                .filter(Boolean);
            if (candidates.some((c) => c === em || c.endsWith('|' + em) || c.includes(em))) {
                return String(rows[i].id);
            }
        }
        path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
    }

    // Fallback: Graph User → versuchen, über ensure via display name match
    try {
        const user = await G().graphJson(
            'GET',
            "/users?$filter=" +
                encodeURIComponent("mail eq '" + em.replace(/'/g, "''") + "' or userPrincipalName eq '" + em.replace(/'/g, "''") + "'") +
                '&$select=id,displayName,mail,userPrincipalName&$top=1',
            tok,
            undefined,
            'v1.0'
        );
        const u = ((user && user.value) || [])[0];
        if (u && u.displayName) {
            // zweiter Durchlauf: Name
            // no-op – ohne ensureUser keine Garantie
        }
    } catch {
        /* ignore */
    }

    throw new Error(
        'Klassenvorstand „' +
            em +
            '“ ist auf der SharePoint-Site noch nicht bekannt. Bitte einmal die Site öffnen oder den KV manuell in SharePoint hinzufügen, dann erneut versuchen.'
    );
}

/**
 * @param {string} webUrl
 * @param {{ listName?: string, listId?: string }} [opts]
 */
export async function resolveFrContext(webUrl, opts) {
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die SharePoint-Website-URL eintragen.');
    const tok = await tokenRead();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    const listName = String((opts && opts.listName) || LIST_TITLE_DEFAULT).trim() || LIST_TITLE_DEFAULT;
    const list = await findFreistellungListOnSite(url, {
        listName,
        listId: opts && opts.listId ? String(opts.listId) : ''
    });
    if (!list) {
        throw new Error(
            'Liste „' +
                listName +
                '“ fehlt. Bitte zuerst unter Freistellungen-Setup anlegen.'
        );
    }
    return {
        webUrl: url,
        siteId,
        siteName: site.displayName || '',
        list: { id: String(list.id), webUrl: list.webUrl || '', name: list.displayName || listName }
    };
}

/**
 * Listen-ID für Remote-Sync (ohne Fehler, wenn Liste noch nicht erreichbar).
 * @param {string} webUrl
 * @param {{ listName?: string, listId?: string }} [opts]
 */
export async function tryResolveFrListId(webUrl, opts) {
    try {
        const ctx = await resolveFrContext(webUrl, opts);
        return ctx && ctx.list && ctx.list.id ? String(ctx.list.id) : '';
    } catch {
        return '';
    }
}

/**
 * Prüft, ob das Konto die Freistellungsliste lesen darf (ohne alle Items zu laden).
 * @param {{ siteId: string, list: { id: string } }} ctx
 */
export async function probeFreistellungListRead(ctx) {
    if (!ctx || !ctx.siteId || !ctx.list || !ctx.list.id) return false;
    const tok = await tokenRead();
    const path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.list.id) +
        '/items?$select=id&$top=1';
    try {
        await G().graphJson('GET', path, tok, undefined, 'v1.0');
        return true;
    } catch {
        return false;
    }
}

export function mapFreistellungFromItem(item) {
    const f = (item && item.fields) || {};
    const beginn = toIsoDateOnly(f.Beginn) || '';
    const ende = toIsoDateOnly(f.Ende) || beginn;
    const kv = readPersonField(f, 'Klassenvorstand');
    const authorPerson = readPersonField(f, 'Author');
    const createdByUser = (item && item.createdBy && item.createdBy.user) || {};
    const authorEmail = String(
        createdByUser.email ||
            createdByUser.userPrincipalName ||
            authorPerson.email ||
            f.AuthorEmail ||
            ''
    )
        .trim()
        .toLowerCase();
    const authorName = String(
        createdByUser.displayName || authorPerson.name || ''
    ).trim();
    const title = fieldStr(f, 'Title');
    const path = approvalPath(beginn, ende);
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        titel: title,
        schuelerName: title.replace(/^(GENEHMIGT|ABGELEHNT):\s*/i, '').replace(/\s*\([^)]*\)\s*$/, '').trim() || title,
        klasse: fieldStr(f, 'Klasse'),
        beginn,
        ende,
        status: fieldStr(f, 'Status') || 'Ausstehend',
        kategorie: fieldStr(f, 'Kategorie'),
        beschreibung: fieldStr(f, 'Beschreibung'),
        bemerkungen: fieldStr(f, 'Bemerkungen'),
        kvEmail: kv.email,
        kvName: kv.name,
        kvLookupId: kv.lookupId,
        authorEmail,
        authorName,
        beantragtVon: authorEmail,
        dayCount: inclusiveDayCount(beginn, ende),
        multiDay: path.multiDay,
        approvalLabel: path.label,
        genehmigtVonKv: fieldStr(f, 'GenehmigtVonKV'),
        genehmigtAmKv: toIsoDateOnly(f.GenehmigtAmKV) || '',
        genehmigtVonDirektion: fieldStr(f, 'GenehmigtVonDirektion'),
        genehmigtAmDirektion: toIsoDateOnly(f.GenehmigtAmDirektion) || '',
        abgelehntVon: fieldStr(f, 'AbgelehntVon'),
        abgelehntAm: toIsoDateOnly(f.AbgelehntAm) || '',
        nachweise: parseNachweiseField(f.Nachweise)
    };
}

/**
 * @param {object} draft
 * @param {{ kvLookupId?: string }} [extra]
 */
export function mapFreistellungToFields(draft, extra) {
    const beginn = toIsoDateOnly(draft.beginn);
    const ende = toIsoDateOnly(draft.ende) || beginn;
    const fields = {
        Title: buildAntragTitle(draft),
        Beginn: beginn,
        Ende: ende,
        Status: String(draft.status || 'Ausstehend').trim() || 'Ausstehend',
        Klasse: String(draft.klasse || '').trim(),
        Kategorie: String(draft.kategorie || '').trim(),
        Beschreibung: String(draft.beschreibung || ''),
        Bemerkungen: String(draft.bemerkungen || '')
    };
    const lookupId = (extra && extra.kvLookupId) || draft.kvLookupId || '';
    if (lookupId) {
        fields.KlassenvorstandLookupId = String(lookupId);
    }
    return fields;
}

export async function loadAllFreistellungen(ctx) {
    const tok = await tokenRead();
    // Nur fields expandieren – createdBy ist bei listItem keine Navigation Property
    // (kommt oft ohnehin im Default-Payload; sonst Author aus fields).
    let path =
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.list.id) +
        '/items?$expand=fields&$top=100';
    const rows = [];
    while (path) {
        const data = await G().graphJson(
            'GET',
            path.indexOf('http') === 0 ? path : path,
            tok,
            undefined,
            'v1.0'
        );
        const batch = (data && data.value) || [];
        for (let i = 0; i < batch.length; i++) rows.push(batch[i]);
        path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
    }
    return rows.map(mapFreistellungFromItem);
}

/**
 * @param {string} siteId
 * @param {string} listId
 * @param {string[]} choices
 */
/**
 * Auswahlfeld „Klasse“ der Freistellungsliste (für Schüler-Dropdown).
 * @param {{ siteId: string, list: { id: string } }} ctx
 */
function normalizeKlasseChoiceRows(raw) {
    const out = [];
    const seen = new Set();
    (raw || []).forEach((entry) => {
        const code = String(entry || '').trim();
        if (!code || seen.has(code.toLowerCase())) return;
        seen.add(code.toLowerCase());
        out.push({ code, name: code });
    });
    return out;
}

function choicesFromSpoFieldPayload(data) {
    const d = data && data.d ? data.d : data;
    if (!d) return [];
    const ch = d.Choices;
    if (Array.isArray(ch)) return ch;
    if (ch && Array.isArray(ch.results)) return ch.results;
    return [];
}

async function fetchFreistellungKlasseChoicesViaSpoRest(webUrl, listId) {
    const url = String(webUrl || '').trim().replace(/\/$/, '');
    const id = String(listId || '').trim();
    if (!url || !id) return [];
    let host = '';
    try {
        host = new URL(url).hostname;
    } catch {
        return [];
    }
    const spoScope = 'https://' + host + '/AllSites.Read';
    const spoTok = await G().getGraphToken([spoScope, 'https://graph.microsoft.com/User.Read']);
    const digest = await G().getSpoRequestDigest(url, spoTok);
    const api =
        "/_api/web/lists(guid'" +
        id.replace(/'/g, "''") +
        "')/fields/getbytitle('Klasse')?$select=Title,Choices,Choice";
    const res = await G().spoRestFetch(url, spoTok, digest, 'GET', api);
    if (!res || !res.ok) return [];
    return normalizeKlasseChoiceRows(choicesFromSpoFieldPayload(res.data));
}

export async function fetchFreistellungKlasseColumnChoices(ctx) {
    if (!ctx || !ctx.list || !ctx.list.id) return [];
    const listId = String(ctx.list.id);
    const webUrl = String(ctx.webUrl || '').trim();
    let fromGraph = [];
    if (ctx.siteId) {
        try {
            const tok = await tokenRead();
            const base =
                G().graphPathSite(ctx.siteId) +
                '/lists/' +
                encodeURIComponent(listId) +
                '/columns';
            const data = await G().graphJson(
                'GET',
                base + '?$select=id,name,choice,displayName&$top=200',
                tok,
                undefined,
                'v1.0'
            );
            const col = ((data && data.value) || []).find((c) => {
                const n = String(c.name || '').toLowerCase();
                const dn = String(c.displayName || '').toLowerCase();
                return n === 'klasse' || dn === 'klasse';
            });
            const raw =
                col && col.choice && Array.isArray(col.choice.choices) ? col.choice.choices : [];
            fromGraph = normalizeKlasseChoiceRows(raw);
        } catch {
            fromGraph = [];
        }
    }
    if (fromGraph.length) return fromGraph;
    try {
        return await fetchFreistellungKlasseChoicesViaSpoRest(webUrl, listId);
    } catch {
        return [];
    }
}

/**
 * @param {string} siteId
 * @param {string} listId
 * @param {string[]} choices
 */
/**
 * Alte Schema-Defaults (nicht Stammdaten) – nicht im Schüler-Dropdown zeigen.
 * @param {string} code
 */
export function isLegacyFreistellungKlasseChoice(code) {
    return /^[1-5]AHW$/i.test(String(code || '').trim());
}

/**
 * @param {{ code?: string, name?: string }[]} rows
 */
export function filterLegacyFreistellungKlasseChoices(rows) {
    return (rows || []).filter((r) => !isLegacyFreistellungKlasseChoice(r && (r.code || r.name)));
}

function normalizeKlasseChoiceCodes(choices) {
    const merged = [];
    const seen = new Set();
    (choices || []).forEach((entry) => {
        const code = String(entry || '').trim();
        if (!code || seen.has(code.toLowerCase()) || isLegacyFreistellungKlasseChoice(code)) return;
        seen.add(code.toLowerCase());
        merged.push(code);
    });
    return merged;
}

function choiceSetsMatch(expected, actual) {
    const a = new Set((expected || []).map((x) => String(x).toLowerCase()));
    const b = new Set((actual || []).map((x) => String(x).toLowerCase()));
    if (!a.size || a.size !== b.size) return false;
    for (const x of a) if (!b.has(x)) return false;
    return true;
}

function escapeXmlAttr(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function buildKlasseChoiceSchemaXml(internalName, displayName, choices) {
    const name = String(internalName || 'Klasse').replace(/[^A-Za-z0-9_]/g, '') || 'Klasse';
    const dn = escapeXmlAttr(displayName || 'Klasse');
    const choiceXml = (choices || [])
        .map((c) => '<CHOICE>' + escapeXmlAttr(c) + '</CHOICE>')
        .join('');
    return (
        '<Field Type="Choice" DisplayName="' +
        dn +
        '" Name="' +
        name +
        '" StaticName="' +
        name +
        '" Format="Dropdown" FillInChoice="FALSE" Required="FALSE">' +
        '<CHOICES>' +
        choiceXml +
        '</CHOICES>' +
        '</Field>'
    );
}

async function acquireSpoWriteToken(webUrl) {
    const origin = String(webUrl || '').trim().replace(/\/$/, '');
    let host = '';
    try {
        host = new URL(origin).hostname;
    } catch {
        throw new Error('Ungültige Site-URL.');
    }
    const graphExtras = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.ReadWrite.All'
    ];
    let spoTok = null;
    const spoScopeSets = [
        ['https://' + host + '/AllSites.FullControl'],
        ['https://' + host + '/AllSites.Write'],
        ['https://' + host + '/.default']
    ];
    let lastErr = null;
    for (let i = 0; i < spoScopeSets.length; i++) {
        try {
            spoTok = await G().getGraphToken(spoScopeSets[i].concat(graphExtras));
            if (spoTok) break;
        } catch (e) {
            lastErr = e;
        }
    }
    if (!spoTok) {
        throw new Error(
            'Kein SharePoint-Schreibrecht (AllSites.Write). ' +
                (lastErr && lastErr.message ? lastErr.message : '')
        );
    }
    const digest = await G().getSpoRequestDigest(origin, spoTok);
    return { origin, spoTok, digest };
}

async function spoMergeField(origin, spoTok, digest, listId, fieldApi, body, useVerbose) {
    const url =
        origin +
        "/_api/web/lists(guid'" +
        String(listId).replace(/'/g, "''") +
        "')/fields/" +
        fieldApi;
    const headers = {
        Authorization: 'Bearer ' + spoTok,
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
        Accept: useVerbose ? 'application/json;odata=verbose' : 'application/json;odata=nometadata',
        'Content-Type': useVerbose
            ? 'application/json;odata=verbose;charset=utf-8'
            : 'application/json;odata=nometadata;charset=utf-8'
    };
    if (digest) headers['X-RequestDigest'] = digest;
    const res = await fetch(url, {
        method: 'POST',
        headers,
        body: JSON.stringify(body)
    });
    const text = await res.text();
    return { ok: res.ok || res.status === 204, status: res.status, text };
}

async function readKlasseChoicesViaSpoRest(webUrl, listId) {
    const rows = await fetchFreistellungKlasseChoicesViaSpoRest(webUrl, listId);
    return rows.map((r) => r.code);
}

async function readKlasseChoicesViaGraph(siteId, listId) {
    const tok = await token();
    const base =
        G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
    const data = await G().graphJson(
        'GET',
        base + '?$select=id,name,displayName,choice&$top=200',
        tok,
        undefined,
        'v1.0'
    );
    const col = ((data && data.value) || []).find((c) => {
        const n = String(c.name || '').toLowerCase();
        const dn = String(c.displayName || '').toLowerCase();
        return n === 'klasse' || dn === 'klasse';
    });
    if (!col) {
        return { codes: [], columnId: '', internalName: 'Klasse', displayName: 'Klasse', allCodes: [] };
    }
    const raw = col.choice && Array.isArray(col.choice.choices) ? col.choice.choices : [];
    const allCodes = (raw || []).map((x) => String(x || '').trim()).filter(Boolean);
    return {
        codes: normalizeKlasseChoiceCodes(allCodes),
        columnId: String(col.id || ''),
        internalName: String(col.name || 'Klasse'),
        displayName: String(col.displayName || 'Klasse'),
        allCodes
    };
}

function replaceChoicesInSchemaXml(schemaXml, choices) {
    const xml = String(schemaXml || '');
    if (!xml) return '';
    const choiceXml = (choices || [])
        .map((c) => '<CHOICE>' + escapeXmlAttr(c) + '</CHOICE>')
        .join('');
    const block = '<CHOICES>' + choiceXml + '</CHOICES>';
    if (/<CHOICES>[\s\S]*?<\/CHOICES>/i.test(xml)) {
        return xml.replace(/<CHOICES>[\s\S]*?<\/CHOICES>/i, block);
    }
    if (/<\/Field>/i.test(xml)) {
        return xml.replace(/<\/Field>/i, block + '</Field>');
    }
    return xml;
}

async function spoGetFieldMeta(origin, spoTok, digest, listId, fieldApi) {
    const url =
        origin +
        "/_api/web/lists(guid'" +
        String(listId).replace(/'/g, "''") +
        "')/fields/" +
        fieldApi +
        '?$select=Id,InternalName,Title,TypeAsString,SchemaXml,Choices';
    const headers = {
        Authorization: 'Bearer ' + spoTok,
        Accept: 'application/json;odata=verbose'
    };
    if (digest) headers['X-RequestDigest'] = digest;
    const res = await fetch(url, { method: 'GET', headers });
    const text = await res.text();
    let data = null;
    try {
        data = text ? JSON.parse(text) : null;
    } catch {
        data = null;
    }
    const d = data && data.d ? data.d : data;
    if (!res.ok || !d) return null;
    return {
        id: String(d.Id || ''),
        internalName: String(d.InternalName || ''),
        title: String(d.Title || ''),
        typeAsString: String(d.TypeAsString || ''),
        schemaXml: String(d.SchemaXml || ''),
        choices: choicesFromSpoFieldPayload(data)
    };
}

async function patchFreistellungKlasseColumnViaSpoRest(webUrl, listId, choices, fieldMeta) {
    const { origin, spoTok, digest } = await acquireSpoWriteToken(webUrl);
    const id = String(listId || '').trim();
    const internal = String((fieldMeta && fieldMeta.internalName) || 'Klasse').trim() || 'Klasse';
    const display = String((fieldMeta && fieldMeta.displayName) || 'Klasse').trim() || 'Klasse';
    const fieldApis = [
        "getbyinternalnameortitle('" + internal.replace(/'/g, "''") + "')",
        "getbytitle('Klasse')"
    ];

    const attempts = [];
    for (let fi = 0; fi < fieldApis.length; fi++) {
        const fieldApi = fieldApis[fi];
        const meta = await spoGetFieldMeta(origin, spoTok, digest, id, fieldApi);

        // 1) Choices-Collection ersetzen
        attempts.push(
            await spoMergeField(
                origin,
                spoTok,
                digest,
                id,
                fieldApi,
                {
                    __metadata: { type: 'SP.FieldChoice' },
                    FillInChoice: false,
                    Choices: {
                        __metadata: { type: 'Collection(Edm.String)' },
                        results: choices
                    }
                },
                true
            )
        );

        let verified = await readKlasseChoicesViaSpoRest(origin, id);
        if (choiceSetsMatch(choices, verified)) {
            return { ok: true, count: choices.length, via: 'spo-choices', verified };
        }

        // 2) SchemaXml mit bestehender Feld-Definition (zuverlässig)
        if (meta && meta.schemaXml) {
            const patchedXml = replaceChoicesInSchemaXml(meta.schemaXml, choices);
            attempts.push(
                await spoMergeField(
                    origin,
                    spoTok,
                    digest,
                    id,
                    fieldApi,
                    {
                        __metadata: { type: 'SP.Field' },
                        SchemaXml: patchedXml
                    },
                    true
                )
            );
            verified = await readKlasseChoicesViaSpoRest(origin, id);
            if (choiceSetsMatch(choices, verified)) {
                return { ok: true, count: choices.length, via: 'spo-schemaxml', verified };
            }
        } else {
            attempts.push(
                await spoMergeField(
                    origin,
                    spoTok,
                    digest,
                    id,
                    fieldApi,
                    {
                        __metadata: { type: 'SP.Field' },
                        SchemaXml: buildKlasseChoiceSchemaXml(internal, display, choices)
                    },
                    true
                )
            );
        }

        // 3) nometadata-Fallback
        attempts.push(
            await spoMergeField(
                origin,
                spoTok,
                digest,
                id,
                fieldApi,
                { FillInChoice: false, Choices: choices },
                false
            )
        );
        verified = await readKlasseChoicesViaSpoRest(origin, id);
        if (choiceSetsMatch(choices, verified)) {
            return { ok: true, count: choices.length, via: 'spo-nometadata', verified };
        }
    }
    const last = attempts.filter((a) => !a.ok).pop() || attempts[attempts.length - 1];
    return {
        ok: false,
        reason: 'spo-rest-verify',
        status: last && last.status,
        detail: last && last.text ? String(last.text).slice(0, 280) : '',
        verified: await readKlasseChoicesViaSpoRest(origin, id)
    };
}

/**
 * Stammdaten-Klassen in SharePoint-Spalte „Klasse“ schreiben und gegenlesen.
 * @param {string} siteId
 * @param {string} listId
 * @param {string[]} choices
 * @param {{ webUrl?: string }} [opts]
 */
export async function patchFreistellungKlasseColumn(siteId, listId, choices, opts) {
    const merged = normalizeKlasseChoiceCodes(choices);
    if (!merged.length) return { ok: false, reason: 'no-choices' };

    const webUrl = String((opts && opts.webUrl) || '').trim();
    let graphMeta = { codes: [], columnId: '', internalName: 'Klasse', displayName: 'Klasse', allCodes: [] };
    try {
        graphMeta = await readKlasseChoicesViaGraph(siteId, listId);
    } catch {
        /* SPO only */
    }

    if (graphMeta.columnId) {
        try {
            const tok = await token();
            const base =
                G().graphPathSite(siteId) +
                '/lists/' +
                encodeURIComponent(listId) +
                '/columns/' +
                encodeURIComponent(graphMeta.columnId);
            await G().graphJson(
                'PATCH',
                base,
                tok,
                {
                    choice: {
                        allowTextEntry: false,
                        choices: merged
                    }
                },
                'v1.0'
            );
            const afterGraph = await readKlasseChoicesViaGraph(siteId, listId);
            if (choiceSetsMatch(merged, afterGraph.allCodes || afterGraph.codes)) {
                return {
                    ok: true,
                    count: merged.length,
                    via: 'graph',
                    verified: afterGraph.allCodes || afterGraph.codes
                };
            }
        } catch {
            /* SPO-Fallback */
        }
    }

    if (webUrl) {
        const viaRest = await patchFreistellungKlasseColumnViaSpoRest(webUrl, listId, merged, {
            internalName: graphMeta.internalName || 'Klasse',
            displayName: graphMeta.displayName || 'Klasse'
        });
        if (viaRest && viaRest.ok) return viaRest;

        const still = viaRest && viaRest.verified ? viaRest.verified : [];
        throw new Error(
            'SharePoint-Spalte „Klasse“ wurde nicht überschrieben. Aktuell noch: ' +
                (still.length ? still.join(', ') : 'unbekannt') +
                '. Erwartet: ' +
                merged.join(', ') +
                (viaRest && viaRest.detail ? ' (' + viaRest.detail + ')' : '') +
                '. Bitte als Site-Besitzer speichern und die richtige Liste (Freistellungen) prüfen.'
        );
    }

    return { ok: false, reason: 'no-weburl' };
}

export async function patchFreistellungKategorieColumn(siteId, listId, choices) {
    const tok = await token();
    const base =
        G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
    const data = await G().graphJson('GET', base + '?$select=id,name&$top=200', tok, undefined, 'v1.0');
    const col = ((data && data.value) || []).find(
        (c) => String(c.name || '').toLowerCase() === 'kategorie'
    );
    if (!col || !col.id) return { ok: false, reason: 'no-column' };
    const merged = mergeKategorieChoices(choices);
    await G().graphJson(
        'PATCH',
        base + '/' + encodeURIComponent(col.id),
        tok,
        {
            choice: {
                allowTextEntry: true,
                choices: merged.length ? merged : [...KATEGORIE_CHOICES]
            }
        },
        'v1.0'
    );
    return { ok: true, count: merged.length };
}

export async function createFreistellungItem(ctx, draft) {
    const tok = await token();
    const kvEmail = String(draft.kvEmail || '')
        .trim()
        .toLowerCase();
    const kvLookupId = await resolvePersonLookupId(tok, ctx.siteId, kvEmail);
    const fields = mapFreistellungToFields(
        { ...draft, status: draft.status || 'Ausstehend' },
        { kvLookupId }
    );
    const created = await G().graphJson(
        'POST',
        G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(ctx.list.id) + '/items',
        tok,
        { fields },
        'v1.0'
    );
    const itemId =
        created && created.id != null
            ? String(created.id)
            : created && created.itemId != null
              ? String(created.itemId)
              : '';
    let mapped = mapFreistellungFromItem(created);
    if (itemId && !mapped.itemId) {
        mapped = { ...mapped, itemId };
    }
    if (!mapped.itemId) {
        throw new Error('Antrag wurde angelegt, aber die Listen-Item-ID fehlt in der Graph-Antwort.');
    }
    return mapped;
}

export async function updateFreistellungStatus(ctx, itemId, status, bemerkungen) {
    const tok = await token();
    const fields = {
        Status: String(status || '').trim()
    };
    if (bemerkungen != null) fields.Bemerkungen = String(bemerkungen);
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
}

export async function deleteFreistellungItem(ctx, itemId) {
    const tok = await token();
    await G().graphJson(
        'DELETE',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.list.id) +
            '/items/' +
            encodeURIComponent(itemId),
        tok,
        undefined,
        'v1.0'
    );
}

// fetchAllItems unused externally but kept for potential health checks
export { fetchAllItems };
