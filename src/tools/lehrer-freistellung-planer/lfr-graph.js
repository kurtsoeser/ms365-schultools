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
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        antragId: fieldStr(f, 'AntragId'),
        titel: fieldStr(f, 'Title'),
        beginn: fieldStr(f, 'Beginn') || fieldStr(f, 'Start'),
        ende: fieldStr(f, 'Ende'),
        status: normalizeStatus(fieldStr(f, 'Status') || 'Ausstehend'),
        kategorie: fieldStr(f, 'Kategorie'),
        beschreibung: fieldStr(f, 'Beschreibung'),
        lehrerName: fieldStr(f, 'LehrerName'),
        lehrerEmail: fieldStr(f, 'LehrerEmail').toLowerCase(),
        genehmigtVon: fieldStr(f, 'GenehmigtVonDirektion'),
        genehmigtAm: fieldStr(f, 'GenehmigtAmDirektion'),
        abgelehntVon: fieldStr(f, 'AbgelehntVon'),
        abgelehntAm: fieldStr(f, 'AbgelehntAm'),
        bemerkungDirektion: fieldStr(f, 'BemerkungDirektion')
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

export async function loadAllItems(ctx) {
    const tok = await token();
    const rows = await fetchAllItems(tok, ctx.siteId, ctx.list.id);
    return rows.map(mapItemFromSp);
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
