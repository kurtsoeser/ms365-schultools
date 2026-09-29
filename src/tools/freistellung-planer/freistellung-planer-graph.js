/**
 * SharePoint-CRUD für Freistellungs-Planer.
 */
import { LIST_TITLE_DEFAULT } from './freistellung-planer-schema.js';
import { toIsoDateOnly, approvalPath, inclusiveDayCount } from './freistellung-planer-logic.js';
import { buildAntragTitle } from './freistellung-planer-state.js';

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

    const listsPath =
        G().graphPathSite(siteId) + '/lists?$select=id,displayName,system&$top=200';
    const listsData = await G().graphJson('GET', listsPath, tok, undefined, 'v1.0');
    const lists = (listsData && listsData.value) || [];
    const userList =
        lists.find((l) => String(l.displayName || '') === 'User Information List') ||
        lists.find((l) => /user information/i.test(String(l.displayName || ''))) ||
        null;
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
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    const listName = String((opts && opts.listName) || LIST_TITLE_DEFAULT).trim() || LIST_TITLE_DEFAULT;
    let list = null;
    if (opts && opts.listId) {
        try {
            list = await G().graphJson(
                'GET',
                G().graphPathSite(siteId) +
                    '/lists/' +
                    encodeURIComponent(opts.listId) +
                    '?$select=id,displayName,webUrl',
                tok,
                undefined,
                'v1.0'
            );
        } catch {
            list = null;
        }
    }
    if (!list) list = await findListByTitle(tok, siteId, listName);
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
        approvalLabel: path.label
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
    const tok = await token();
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
    return mapFreistellungFromItem(created);
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
