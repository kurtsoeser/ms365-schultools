/**
 * Demo-Paket Projektwochen: JSON parsen, Stammdaten lokal, SharePoint upsert.
 */
import {
    getDemoSeedPackage,
    buildLocalDemoState,
    DEMO_SEED_TAG,
    DEMO_SITE_DEFAULT
} from './projektwochen-demo-data.js';
import {
    resolvePwContext,
    loadAllPwData,
    adaptAngebotFieldsForList
} from './projektwochen-graph.js';

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

function cleanFields(fields) {
    const out = { ...fields };
    Object.keys(out).forEach((k) => {
        if (out[k] === undefined || out[k] === null || out[k] === '') delete out[k];
    });
    return out;
}

function fieldStr(fields, key) {
    const v = fields && fields[key];
    return v == null ? '' : String(v).trim();
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
    if (!data.aktion || typeof data.aktion !== 'object') {
        throw new Error('Feld „aktion“ fehlt.');
    }
    if (!Array.isArray(data.angebote) || !data.angebote.length) {
        throw new Error('Feld „angebote“ fehlt oder ist leer.');
    }
    if (!data.stammdaten) data.stammdaten = { teachers: [], classes: [], students: [] };
    if (!data.seedTag) data.seedTag = DEMO_SEED_TAG;
    if (!data.siteDefault) data.siteDefault = DEMO_SITE_DEFAULT;
    if (!data.counts) {
        data.counts = {
            angebote: data.angebote.length,
            freigegeben: data.angebote.filter((a) => a && a.Status === 'freigegeben').length,
            beantragt: data.angebote.filter((a) => a && a.Status === 'beantragt').length,
            teachers: (data.stammdaten.teachers || []).length,
            classes: (data.stammdaten.classes || []).length
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
        subjects: Array.isArray(stammdaten.subjects) && stammdaten.subjects.length ? stammdaten.subjects : cur.subjects,
        classes: Array.isArray(stammdaten.classes) && stammdaten.classes.length ? stammdaten.classes : cur.classes,
        teachers: Array.isArray(stammdaten.teachers) && stammdaten.teachers.length ? stammdaten.teachers : cur.teachers,
        students:
            Array.isArray(stammdaten.students) && stammdaten.students.length
                ? stammdaten.students
                : cur.students
    };
    window.ms365TenantSettingsSave(next);
    return true;
}

/**
 * Demo auf SharePoint schreiben (Upsert über AktionId / AngebotId).
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 * @param {{ pack?: object }} [opts]
 */
export async function seedDemoProjektwochen(webUrl, logFn, opts) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('SharePoint-Website-URL fehlt.');
    const pack = parseDemoImportJson((opts && opts.pack) || getDemoSeedPackage());

    write('Löse Site / Listen auf …');
    const ctx = await resolvePwContext(url);
    const tok = await token();
    const weekdayField = ctx.weekdayField || 'Wochentag';
    if (weekdayField === 'Tag') {
        write('Hinweis: Legacy-Spalte „Tag“ wird verwendet (kein „Wochentag“).');
    }

    const aktionItems = await fetchAllItems(tok, ctx.siteId, ctx.lists.aktionen.id);
    const angebotItems = await fetchAllItems(tok, ctx.siteId, ctx.lists.angebote.id);

    const aktionFields = cleanFields(pack.aktion);
    const aktionId = String(aktionFields.AktionId || '').trim();
    const existingAktion = aktionItems.find((it) => fieldStr(it.fields, 'AktionId') === aktionId);
    if (existingAktion && existingAktion.id != null) {
        write('Aktualisiere Aktion ' + aktionId + ' …');
        await patchItemFields(tok, ctx.siteId, ctx.lists.aktionen.id, existingAktion.id, aktionFields);
    } else {
        write('Lege Aktion ' + aktionId + ' an …');
        await createItem(tok, ctx.siteId, ctx.lists.aktionen.id, aktionFields);
    }

    // Leere Setup-Seed-Aktion („pw-demo“) schließen, damit sie nicht die Ansicht verdeckt
    for (let i = 0; i < aktionItems.length; i++) {
        const it = aktionItems[i];
        const oid = fieldStr(it.fields, 'AktionId');
        if (!oid || oid === aktionId) continue;
        if (oid === 'pw-demo' && fieldStr(it.fields, 'Status') === 'offen' && it.id != null) {
            write('Schließe leere Setup-Aktion „pw-demo“ …');
            await patchItemFields(tok, ctx.siteId, ctx.lists.aktionen.id, it.id, { Status: 'geschlossen' });
        }
    }

    let created = 0;
    let updated = 0;
    for (let i = 0; i < pack.angebote.length; i++) {
        const raw = pack.angebote[i];
        const fields = cleanFields(adaptAngebotFieldsForList(raw, weekdayField));
        const aid = String(fields.AngebotId || '').trim();
        if (!aid) continue;
        const match = angebotItems.find((it) => fieldStr(it.fields, 'AngebotId') === aid);
        if (match && match.id != null) {
            await patchItemFields(tok, ctx.siteId, ctx.lists.angebote.id, match.id, fields);
            updated += 1;
        } else {
            await createItem(tok, ctx.siteId, ctx.lists.angebote.id, fields);
            created += 1;
        }
        if ((i + 1) % 5 === 0) write('  … ' + (i + 1) + '/' + pack.angebote.length);
    }

    write(
        'Fertig. Aktion upsert · Angebote neu ' +
            created +
            ' / aktualisiert ' +
            updated +
            ' (Seed ' +
            (pack.seedTag || DEMO_SEED_TAG) +
            ').'
    );

    try {
        if (typeof window !== 'undefined' && typeof window.ms365LogAction === 'function') {
            window.ms365LogAction({
                action: 'seed-projektwochen-demo',
                detail:
                    'Demo ' +
                    (pack.seedTag || DEMO_SEED_TAG) +
                    ' · ' +
                    pack.angebote.length +
                    ' Angebote'
            });
        }
    } catch {
        /* ignore */
    }

    return { ctx, pack, created, updated };
}

/**
 * Nach Seed neu laden.
 * @param {string} webUrl
 */
export async function reloadAfterSeed(webUrl) {
    const ctx = await resolvePwContext(webUrl);
    const data = await loadAllPwData(ctx);
    return { ctx, data };
}

export {
    getDemoSeedPackage,
    buildLocalDemoState,
    DEMO_SITE_DEFAULT,
    DEMO_SEED_TAG
};
