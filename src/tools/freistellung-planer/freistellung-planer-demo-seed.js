/**
 * Demo Freistellungen: SharePoint upserten / gezielt zurücksetzen (Seed-Tag).
 */
import { LIST_TITLE_DEFAULT } from './freistellung-planer-schema.js';
import { resolvePersonLookupId } from './freistellung-planer-graph.js';
import {
    getDemoSeedPackage,
    parseDemoImportJson,
    DEMO_SEED_TAG,
    DEMO_SITE_DEFAULT,
    extractDemoId,
    isDemoFreistellungFields,
    toSharePointFields
} from './freistellung-planer-demo.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen.');
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
        const data = await G().graphJson('GET', path.indexOf('http') === 0 ? path : path, tok, undefined, 'v1.0');
        const rows = (data && data.value) || [];
        for (let i = 0; i < rows.length; i++) out.push(rows[i]);
        path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
    }
    return out;
}

async function createItem(tok, siteId, listId, fields) {
    await G().graphJson(
        'POST',
        G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/items',
        tok,
        { fields },
        'v1.0'
    );
    if (typeof G().sleep === 'function') await G().sleep(100);
}

async function patchItemFields(tok, siteId, listId, itemId, fields) {
    await G().graphJson(
        'PATCH',
        G().graphPathSite(siteId) +
            '/lists/' +
            encodeURIComponent(listId) +
            '/items/' +
            encodeURIComponent(itemId) +
            '/fields',
        tok,
        fields,
        'v1.0'
    );
    if (typeof G().sleep === 'function') await G().sleep(100);
}

async function deleteItem(tok, siteId, listId, itemId) {
    await G().graphJson(
        'DELETE',
        G().graphPathSite(siteId) +
            '/lists/' +
            encodeURIComponent(listId) +
            '/items/' +
            encodeURIComponent(itemId),
        tok,
        undefined,
        'v1.0'
    );
    if (typeof G().sleep === 'function') await G().sleep(100);
}

/**
 * @param {object} stammdaten
 */
export function applyDemoStammdatenLocal(stammdaten) {
    if (!stammdaten || typeof window === 'undefined') return false;
    if (
        typeof window.ms365TenantSettingsLoad !== 'function' ||
        typeof window.ms365TenantSettingsSave !== 'function'
    ) {
        return false;
    }
    const cur = window.ms365TenantSettingsLoad() || {};
    const data = (cur && cur.data) || cur || {};
    const next = {
        ...cur,
        schoolName: stammdaten.schoolName || data.schoolName || cur.schoolName,
        domain: stammdaten.domain || data.domain || cur.domain,
        classes:
            Array.isArray(stammdaten.classes) && stammdaten.classes.length
                ? stammdaten.classes
                : data.classes || cur.classes,
        teachers:
            Array.isArray(stammdaten.teachers) && stammdaten.teachers.length
                ? stammdaten.teachers
                : data.teachers || cur.teachers
    };
    window.ms365TenantSettingsSave(next);
    return true;
}

/**
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 * @param {{ pack?: object, listName?: string, listId?: string }} [opts]
 */
export async function seedDemoFreistellungen(webUrl, logFn, opts) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('SharePoint-Website-URL fehlt.');
    const pack = parseDemoImportJson((opts && opts.pack) || getDemoSeedPackage());
    const listName = String((opts && opts.listName) || LIST_TITLE_DEFAULT).trim() || LIST_TITLE_DEFAULT;

    const tok = await token();
    write('Löse Site auf …');
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    write('Site: ' + (site.displayName || siteId));

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
    if (!list || !list.id) {
        throw new Error('Liste „' + listName + '“ fehlt. Bitte zuerst Freistellungen-Setup ausführen.');
    }
    write('OK Liste: ' + (list.displayName || listName));

    write('Löse Klassenvorstände (Person-Felder) …');
    const kvCache = new Map();
    const uniqueKv = [
        ...new Set(
            pack.freistellungen
                .map((r) => String(r._kvEmail || '').toLowerCase())
                .filter((e) => e.includes('@'))
        )
    ];
    for (let i = 0; i < uniqueKv.length; i++) {
        const em = uniqueKv[i];
        try {
            const id = await resolvePersonLookupId(tok, siteId, em);
            kvCache.set(em, id);
            write('  KV ok: ' + em);
        } catch (e) {
            write('  KV übersprungen (' + em + '): ' + String((e && e.message) || e));
        }
    }

    write('Freistellungen (' + pack.freistellungen.length + ') …');
    const existing = await fetchAllItems(tok, siteId, list.id);
    const byDemoId = new Map();
    existing.forEach((it) => {
        const id = extractDemoId(it.fields && it.fields.Beschreibung);
        if (id) byDemoId.set(id, it);
    });

    let created = 0;
    let updated = 0;
    for (let i = 0; i < pack.freistellungen.length; i++) {
        const row = pack.freistellungen[i];
        const demoId = row._demoId || extractDemoId(row.Beschreibung);
        const fields = toSharePointFields(row);
        const lookup = kvCache.get(String(row._kvEmail || '').toLowerCase());
        if (lookup) fields.KlassenvorstandLookupId = lookup;

        const existingItem = demoId ? byDemoId.get(demoId) : null;
        if (existingItem) {
            await patchItemFields(tok, siteId, list.id, existingItem.id, fields);
            updated++;
        } else {
            await createItem(tok, siteId, list.id, fields);
            created++;
        }
        if ((created + updated) % 8 === 0) {
            write('  … ' + (created + updated) + '/' + pack.freistellungen.length);
        }
    }
    write('  Fertig +' + created + ' / ~' + updated);

    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
        window.ms365ActionLog.append({
            tool: 'freistellung-planer',
            action: 'seed-freistellung-demo',
            target: url,
            summary:
                'Demo ' +
                (pack.schoolYear || '') +
                ': +' +
                created +
                ' / ~' +
                updated +
                ' (' +
                pack.seedTag +
                ')'
        });
    }

    return { created, updated, total: pack.freistellungen.length, seedTag: pack.seedTag };
}

/**
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 * @param {{ listName?: string, listId?: string, seedTag?: string }} [opts]
 */
export async function resetDemoFreistellungen(webUrl, logFn, opts) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('SharePoint-Website-URL fehlt.');
    const seedTag = String((opts && opts.seedTag) || DEMO_SEED_TAG);
    const listName = String((opts && opts.listName) || LIST_TITLE_DEFAULT).trim() || LIST_TITLE_DEFAULT;

    const tok = await token();
    write('Löse Site auf …');
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

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
    if (!list || !list.id) throw new Error('Liste „' + listName + '“ fehlt.');

    const existing = await fetchAllItems(tok, siteId, list.id);
    const demoRows = existing.filter((it) => isDemoFreistellungFields(it.fields, seedTag));
    write('Lösche ' + demoRows.length + ' Demo-Einträge …');
    let deleted = 0;
    for (let i = 0; i < demoRows.length; i++) {
        await deleteItem(tok, siteId, list.id, demoRows[i].id);
        deleted++;
    }
    write('  Gelöscht: ' + deleted);

    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
        window.ms365ActionLog.append({
            tool: 'freistellung-planer',
            action: 'reset-freistellung-demo',
            target: url,
            summary: 'Demo-Reset ' + seedTag + ': ' + deleted + ' gelöscht'
        });
    }

    return { deleted, seedTag };
}

export { DEMO_SITE_DEFAULT, DEMO_SEED_TAG };
