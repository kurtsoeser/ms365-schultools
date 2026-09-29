/**
 * Demo-Daten Schulaktivitäten: lokal laden, auf SharePoint upserten, Demo zurücksetzen.
 */
import { LIST_TITLES } from './schulaktivitaeten-planer-schema.js';
import {
    getDemoSeedPackage,
    buildLocalDemoState,
    DEMO_SEED_TAG,
    DEMO_SITE_DEFAULT
} from './schulaktivitaeten-planer-demo-data.js';

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
        const data = await G().graphJson('GET', path, tok, undefined, 'v1.0');
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
    if (typeof G().sleep === 'function') await G().sleep(80);
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
    if (typeof G().sleep === 'function') await G().sleep(80);
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
    if (typeof G().sleep === 'function') await G().sleep(80);
}

function cleanFields(fields) {
    const out = { ...fields };
    Object.keys(out).forEach((k) => {
        if (out[k] === undefined || out[k] === null) delete out[k];
    });
    return out;
}

function isDemoRow(fields, seedTag) {
    const tag = String(seedTag || DEMO_SEED_TAG);
    const id = String((fields && fields.AktivitaetId) || '');
    const notiz = String((fields && fields.Notiz) || '');
    return id.indexOf('akt-demo-') === 0 || notiz.indexOf(tag) !== -1;
}

/**
 * @param {string|object} raw
 */
export function parseDemoImportJson(raw) {
    let data = raw;
    if (typeof raw === 'string') {
        try {
            data = JSON.parse(raw);
        } catch {
            throw new Error('Keine gültige JSON-Datei.');
        }
    }
    if (!data || typeof data !== 'object') throw new Error('Leeres Demo-Paket.');
    if (!Array.isArray(data.aktivitaeten) || !data.aktivitaeten.length) {
        throw new Error('Feld „aktivitaeten“ fehlt oder ist leer.');
    }
    if (!data.regelwerk || typeof data.regelwerk !== 'object') {
        data.regelwerk = getDemoSeedPackage().regelwerk;
    }
    if (!data.stammdaten) data.stammdaten = getDemoSeedPackage().stammdaten;
    if (!data.seedTag) data.seedTag = DEMO_SEED_TAG;
    if (!data.counts) {
        data.counts = {
            aktivitaeten: data.aktivitaeten.length,
            genehmigt: data.aktivitaeten.filter((a) => a.Status === 'genehmigt').length,
            beantragt: data.aktivitaeten.filter((a) => a.Status === 'beantragt').length,
            abgelehnt: data.aktivitaeten.filter((a) => a.Status === 'abgelehnt').length
        };
    }
    return data;
}

/**
 * @param {object} stammdaten
 */
export function applyDemoStammdatenLocal(stammdaten) {
    if (!stammdaten || typeof window === 'undefined') return false;
    if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
        return false;
    }
    const cur = window.ms365TenantSettingsLoad() || {};
    const next = {
        ...cur,
        schoolName: stammdaten.schoolName || cur.schoolName,
        domain: stammdaten.domain || cur.domain,
        classes:
            Array.isArray(stammdaten.classes) && stammdaten.classes.length
                ? stammdaten.classes
                : cur.classes,
        teachers:
            Array.isArray(stammdaten.teachers) && stammdaten.teachers.length
                ? stammdaten.teachers
                : cur.teachers
    };
    window.ms365TenantSettingsSave(next);
    return true;
}

/**
 * Demo auf SharePoint schreiben (Upsert über AktivitaetId).
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 * @param {{ pack?: object }} [opts]
 */
export async function seedDemoSchulaktivitaeten(webUrl, logFn, opts) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('SharePoint-Website-URL fehlt.');
    const pack = parseDemoImportJson((opts && opts.pack) || getDemoSeedPackage());

    const tok = await token();
    write('Löse Site auf …');
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    write('Site: ' + (site.displayName || siteId));

    const listAkt = await findListByTitle(tok, siteId, LIST_TITLES.aktivitaeten);
    if (!listAkt || !listAkt.id) {
        throw new Error(
            'Liste „' + LIST_TITLES.aktivitaeten + '“ fehlt. Bitte zuerst Listen anlegen.'
        );
    }
    write('OK Liste: ' + LIST_TITLES.aktivitaeten);

    let listRw = await findListByTitle(tok, siteId, LIST_TITLES.regelwerk);
    if (listRw && listRw.id) {
        write('OK Liste: ' + LIST_TITLES.regelwerk);
        write('Regelwerk …');
        const rwItems = await fetchAllItems(tok, siteId, listRw.id);
        const rwDemo = rwItems.find(
            (it) =>
                String((it.fields && it.fields.RegelwerkId) || '') === pack.regelwerk.RegelwerkId
        );
        if (rwDemo) {
            await patchItemFields(tok, siteId, listRw.id, rwDemo.id, cleanFields(pack.regelwerk));
            write('  Regelwerk aktualisiert.');
        } else {
            await createItem(tok, siteId, listRw.id, cleanFields(pack.regelwerk));
            write('  Regelwerk angelegt.');
        }
    } else {
        write('Hinweis: Regelwerk-Liste fehlt – nur Aktivitäten werden geschrieben.');
    }

    write('Aktivitäten (' + pack.aktivitaeten.length + ') …');
    const existing = await fetchAllItems(tok, siteId, listAkt.id);
    const byId = new Map();
    existing.forEach((it) => {
        const id = String((it.fields && it.fields.AktivitaetId) || '').trim();
        if (id) byId.set(id, it);
    });

    let created = 0;
    let updated = 0;
    for (let i = 0; i < pack.aktivitaeten.length; i++) {
        const row = cleanFields(pack.aktivitaeten[i]);
        const existingItem = byId.get(row.AktivitaetId);
        if (existingItem) {
            await patchItemFields(tok, siteId, listAkt.id, existingItem.id, row);
            updated++;
        } else {
            await createItem(tok, siteId, listAkt.id, row);
            created++;
        }
        if ((created + updated) % 10 === 0) {
            write('  … ' + (created + updated) + '/' + pack.aktivitaeten.length);
        }
    }
    write('  Aktivitäten +' + created + ' / ~' + updated);

    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
        window.ms365ActionLog.append({
            tool: 'sharepoint',
            action: 'seed-schulaktivitaeten-demo',
            target: url,
            summary:
                'Demo ' +
                (pack.schoolYear || '') +
                ' (' +
                (pack.seedTag || DEMO_SEED_TAG) +
                '): ' +
                pack.aktivitaeten.length +
                ' Aktivitäten'
        });
    }

    write('Fertig.');
    return {
        ok: true,
        counts: pack.counts,
        created: { aktivitaeten: created },
        updated: { aktivitaeten: updated }
    };
}

/**
 * Alle Demo-Zeilen (Seed-Tag / akt-demo-*) aus der SharePoint-Liste löschen.
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 * @param {{ seedTag?: string }} [opts]
 */
export async function resetDemoSchulaktivitaeten(webUrl, logFn, opts) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('SharePoint-Website-URL fehlt.');
    const seedTag = (opts && opts.seedTag) || DEMO_SEED_TAG;

    const tok = await token();
    write('Löse Site auf …');
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    const listAkt = await findListByTitle(tok, siteId, LIST_TITLES.aktivitaeten);
    if (!listAkt || !listAkt.id) {
        throw new Error('Liste „' + LIST_TITLES.aktivitaeten + '“ fehlt.');
    }

    const items = await fetchAllItems(tok, siteId, listAkt.id);
    const demoItems = items.filter((it) => isDemoRow(it.fields, seedTag));
    write('Lösche ' + demoItems.length + ' Demo-Einträge (Tag ' + seedTag + ') …');

    let deleted = 0;
    for (let i = 0; i < demoItems.length; i++) {
        await deleteItem(tok, siteId, listAkt.id, demoItems[i].id);
        deleted++;
        if (deleted % 10 === 0) write('  … ' + deleted + '/' + demoItems.length);
    }

    // Demo-Regelwerk optional belassen (kann für echte Nutzung bleiben) – nur wenn Idempotent-ID
    const listRw = await findListByTitle(tok, siteId, LIST_TITLES.regelwerk);
    if (listRw && listRw.id) {
        const rwItems = await fetchAllItems(tok, siteId, listRw.id);
        const pack = getDemoSeedPackage();
        const rwDemo = rwItems.find(
            (it) => String((it.fields && it.fields.RegelwerkId) || '') === pack.regelwerk.RegelwerkId
        );
        if (rwDemo) {
            await deleteItem(tok, siteId, listRw.id, rwDemo.id);
            write('  Demo-Regelwerk entfernt.');
        }
    }

    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
        window.ms365ActionLog.append({
            tool: 'sharepoint',
            action: 'reset-schulaktivitaeten-demo',
            target: url,
            summary: 'Demo zurückgesetzt: ' + deleted + ' Aktivitäten gelöscht'
        });
    }

    write('Zurückgesetzt (' + deleted + ').');
    return { ok: true, deleted };
}

export {
    getDemoSeedPackage,
    buildLocalDemoState,
    DEMO_SEED_TAG,
    DEMO_SITE_DEFAULT
};
