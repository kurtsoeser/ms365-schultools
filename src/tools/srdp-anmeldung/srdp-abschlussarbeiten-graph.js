/**
 * SharePoint-CRUD für ABA-Portal-Abschlussarbeiten.
 */
import {
    abschlussarbeitenListTitleCandidates,
    abschlussarbeitenListTitleForYear,
    buildAbschlussarbeitenColumns
} from './srdp-abschlussarbeiten-schema.js';
import { toGraphColumnBody } from './srdp-anmeldung-schema.js';

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
    return G().getGraphToken(SCOPES);
}

function fieldStr(fields, name) {
    const f = fields || {};
    const v = f[name];
    if (v == null) return '';
    if (typeof v === 'object' && v.DisplayName) return String(v.DisplayName).trim();
    return String(v).trim();
}

/**
 * @param {object} item
 */
export function mapAbschlussarbeitFromItem(item) {
    const f = (item && item.fields) || {};
    return {
        itemId: item && item.id != null ? String(item.id) : '',
        title: fieldStr(f, 'Title'),
        titel: fieldStr(f, 'Titel') || fieldStr(f, 'Title'),
        schuelerName: fieldStr(f, 'SchuelerName'),
        klasseRoh: fieldStr(f, 'KlasseRoh'),
        klasseCode: fieldStr(f, 'KlasseCode'),
        schuljahr: fieldStr(f, 'Schuljahr'),
        betreuer: fieldStr(f, 'Betreuer'),
        status: fieldStr(f, 'Status'),
        schule: fieldStr(f, 'Schule'),
        quelle: fieldStr(f, 'Quelle'),
        importBatch: fieldStr(f, 'ImportBatch'),
        terminJahr: fieldStr(f, 'TerminJahr'),
        schulform: fieldStr(f, 'Schulform'),
        webUrl: (item && item.webUrl) || ''
    };
}

async function findListByTitle(tok, siteId, listTitle) {
    const title = String(listTitle || '').trim();
    if (!title) return null;
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
        /* ignore */
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
        /* ignore */
    }
    return null;
}

async function fetchAllItems(tok, siteId, listId) {
    const out = [];
    let path =
        G().graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(listId) +
        '/items?$expand=fields&$top=100';
    while (path) {
        const data = await G().graphJson(
            'GET',
            path.indexOf('http') === 0 ? path : path,
            tok,
            undefined,
            'v1.0'
        );
        (data && data.value ? data.value : []).forEach((row) => out.push(row));
        path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
    }
    return out;
}

async function addMissingColumns(tok, siteId, listId, defs) {
    const existing = await G().graphJson(
        'GET',
        G().graphPathSite(siteId) +
            '/lists/' +
            encodeURIComponent(listId) +
            '/columns?$select=name,displayName&$top=200',
        tok,
        undefined,
        'v1.0'
    );
    const have = new Set(
        ((existing && existing.value) || []).map((c) => String(c.name || '').toLowerCase())
    );
    let added = 0;
    for (let i = 0; i < defs.length; i++) {
        const def = defs[i];
        if (have.has(String(def.name || '').toLowerCase())) continue;
        await G().graphJson(
            'POST',
            G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns',
            tok,
            toGraphColumnBody(def),
            'v1.0'
        );
        added += 1;
    }
    return added;
}

/**
 * @param {string} webUrl
 * @param {{ profileId?: string, jahr?: string|number }} opts
 */
export async function ensureAbschlussarbeitenList(webUrl, opts) {
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('SharePoint-URL fehlt.');
    const o = opts || {};
    const jahr = String(o.jahr || '').trim();
    const profileId = o.profileId || 'hak';
    const title = abschlussarbeitenListTitleForYear(jahr, profileId);
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    const candidates = abschlussarbeitenListTitleCandidates(jahr, profileId);
    let list = null;
    for (let i = 0; i < candidates.length; i++) {
        list = await findListByTitle(tok, siteId, candidates[i]);
        if (list) break;
    }
    if (!list) {
        const created = await G().graphJson(
            'POST',
            G().graphPathSite(siteId) + '/lists',
            tok,
            {
                displayName: title,
                description:
                    'Abschlussarbeiten aus ABA-Portal-Export für ' +
                    profileId.toUpperCase() +
                    ' ' +
                    jahr +
                    '.',
                list: { template: 'genericList' }
            },
            'v1.0'
        );
        list = {
            id: String(created.id),
            displayName: title,
            webUrl: created.webUrl || ''
        };
    }
    await addMissingColumns(tok, siteId, list.id, buildAbschlussarbeitenColumns());
    return {
        webUrl: url,
        siteId,
        list: {
            id: String(list.id),
            webUrl: list.webUrl || '',
            title: list.displayName || title
        }
    };
}

/**
 * @param {{ siteId: string, list: { id: string } }} ctx
 */
export async function loadAbschlussarbeiten(ctx) {
    const tok = await token();
    const rows = await fetchAllItems(tok, ctx.siteId, ctx.list.id);
    return rows.map(mapAbschlussarbeitFromItem);
}

/**
 * Ersetzt alle Items des Terminjahrs (bzw. der Liste) durch neue Import-Zeilen.
 * @param {{ siteId: string, list: { id: string } }} ctx
 * @param {Array<object>} rows parseAbaPortalRows().rows
 * @param {{ jahr?: string, profileId?: string, batchId?: string }} meta
 */
export async function replaceAbschlussarbeiten(ctx, rows, meta) {
    const m = meta || {};
    const tok = await token();
    const existing = await fetchAllItems(tok, ctx.siteId, ctx.list.id);
    for (let i = 0; i < existing.length; i++) {
        const id = existing[i] && existing[i].id;
        if (!id) continue;
        await G().graphJson(
            'DELETE',
            G().graphPathSite(ctx.siteId) +
                '/lists/' +
                encodeURIComponent(ctx.list.id) +
                '/items/' +
                encodeURIComponent(String(id)),
            tok,
            undefined,
            'v1.0'
        );
    }
    const batch = String(m.batchId || new Date().toISOString()).slice(0, 80);
    const created = [];
    for (let i = 0; i < (rows || []).length; i++) {
        const r = rows[i];
        const titel = String((r && r.titel) || '').trim();
        const schuelerName = String((r && r.schuelerName) || '').trim();
        const fields = {
            Title: schuelerName || titel || 'Abschlussarbeit',
            Titel: titel,
            SchuelerName: schuelerName,
            KlasseRoh: String((r && r.klasseRoh) || '').trim(),
            KlasseCode: String((r && r.klasseCode) || '').trim(),
            Schuljahr: String((r && r.schuljahr) || '').trim(),
            Betreuer: String((r && r.betreuer) || '').trim(),
            Status: String((r && r.status) || '').trim(),
            Schule: String((r && r.schule) || '').trim(),
            Quelle: 'ABA-Portal',
            ImportBatch: batch,
            TerminJahr: String(m.jahr || '').trim(),
            Schulform: String(m.profileId || 'hak')
                .trim()
                .toUpperCase()
        };
        Object.keys(fields).forEach((k) => {
            if (fields[k] === '' || fields[k] == null) delete fields[k];
        });
        fields.Title = schuelerName || titel || 'Abschlussarbeit';
        const item = await G().graphJson(
            'POST',
            G().graphPathSite(ctx.siteId) +
                '/lists/' +
                encodeURIComponent(ctx.list.id) +
                '/items',
            tok,
            { fields },
            'v1.0'
        );
        created.push(mapAbschlussarbeitFromItem(item));
    }
    return { count: created.length, items: created, batchId: batch };
}
