/**
 * Demo-Daten 2026/27 auf SharePoint-Listen schreiben (idempotent per Seed-Tag / IDs).
 */
import { LIST_TITLES, newEntityId } from './schularbeiten-planer-schema.js';
import { getDemoSeedPackage, DEMO_SEED_TAG } from './schularbeiten-planer-demo-data.js';

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
    await G().sleep(80);
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
    await G().sleep(80);
}

function cleanFields(fields) {
    const out = { ...fields };
    Object.keys(out).forEach((k) => {
        if (out[k] === undefined || out[k] === null) delete out[k];
    });
    return out;
}

/**
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 * @param {{ replaceDemo?: boolean, pack?: object }} [opts]
 */
export async function seedDemoSchularbeiten(webUrl, logFn, opts) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('SharePoint-Website-URL fehlt.');
    const replaceDemo = !opts || opts.replaceDemo !== false;
    const pack = (opts && opts.pack) || getDemoSeedPackage();
    if (!pack || !Array.isArray(pack.schularbeiten)) {
        throw new Error('Ungültiges Demo-Paket (schularbeiten fehlt).');
    }

    const tok = await token();
    write('Löse Site auf …');
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    write('Site: ' + (site.displayName || siteId));

    const need = [
        LIST_TITLES.regelwerk,
        LIST_TITLES.terminfenster,
        LIST_TITLES.schularbeiten,
        LIST_TITLES.fachMeta
    ];
    /** @type {Record<string, { id: string }>} */
    const lists = {};
    for (let i = 0; i < need.length; i++) {
        const title = need[i];
        const list = await findListByTitle(tok, siteId, title);
        if (!list || !list.id) {
            throw new Error(
                'Liste „' + title + '“ fehlt. Bitte zuerst „Paket anlegen“ ausführen.'
            );
        }
        lists[title] = { id: String(list.id) };
        write('OK Liste: ' + title);
    }

    // Regelwerk
    write('Regelwerk …');
    const rwItems = await fetchAllItems(tok, siteId, lists[LIST_TITLES.regelwerk].id);
    const rwDemo = rwItems.find(
        (it) => String((it.fields && it.fields.RegelwerkId) || '') === pack.regelwerk.RegelwerkId
    );
    if (rwDemo) {
        await patchItemFields(tok, siteId, lists[LIST_TITLES.regelwerk].id, rwDemo.id, cleanFields(pack.regelwerk));
        write('  Regelwerk aktualisiert.');
    } else {
        await createItem(tok, siteId, lists[LIST_TITLES.regelwerk].id, cleanFields(pack.regelwerk));
        write('  Regelwerk angelegt.');
    }

    // FachMeta
    write('SA-FachMeta (' + pack.fachMeta.length + ') …');
    const metaItems = await fetchAllItems(tok, siteId, lists[LIST_TITLES.fachMeta].id);
    const metaByCode = new Map();
    metaItems.forEach((it) => {
        const code = String((it.fields && it.fields.FachCode) || '').trim();
        if (code) metaByCode.set(code, it);
    });
    let metaCreated = 0;
    let metaUpdated = 0;
    for (let i = 0; i < pack.fachMeta.length; i++) {
        const row = pack.fachMeta[i];
        const existing = metaByCode.get(row.FachCode);
        if (existing) {
            await patchItemFields(tok, siteId, lists[LIST_TITLES.fachMeta].id, existing.id, cleanFields(row));
            metaUpdated++;
        } else {
            await createItem(tok, siteId, lists[LIST_TITLES.fachMeta].id, cleanFields(row));
            metaCreated++;
        }
    }
    write('  FachMeta +' + metaCreated + ' / ~' + metaUpdated);

    // Terminfenster
    write('Terminfenster (' + pack.terminfenster.length + ') …');
    const tfItems = await fetchAllItems(tok, siteId, lists[LIST_TITLES.terminfenster].id);
    const tfById = new Map();
    tfItems.forEach((it) => {
        const id = String((it.fields && it.fields.TerminfensterId) || '').trim();
        if (id) tfById.set(id, it);
    });
    let tfC = 0;
    let tfU = 0;
    for (let i = 0; i < pack.terminfenster.length; i++) {
        const row = pack.terminfenster[i];
        const existing = tfById.get(row.TerminfensterId);
        if (existing) {
            await patchItemFields(tok, siteId, lists[LIST_TITLES.terminfenster].id, existing.id, cleanFields(row));
            tfU++;
        } else {
            await createItem(tok, siteId, lists[LIST_TITLES.terminfenster].id, cleanFields(row));
            tfC++;
        }
    }
    write('  Terminfenster +' + tfC + ' / ~' + tfU);

    // Schularbeiten
    write('Schularbeiten (' + pack.schularbeiten.length + ') …');
    const saItems = await fetchAllItems(tok, siteId, lists[LIST_TITLES.schularbeiten].id);
    const saById = new Map();
    saItems.forEach((it) => {
        const id = String((it.fields && it.fields.SchularbeitId) || '').trim();
        if (id) saById.set(id, it);
    });

    if (replaceDemo) {
        // Alte Demo-Zeilen ohne stabile ID (nur Seed-Tag in Notiz) belassen – Upsert über SchularbeitId
        write('  Upsert nach SchularbeitId (Seed-Tag ' + DEMO_SEED_TAG + ') …');
    }

    let saC = 0;
    let saU = 0;
    for (let i = 0; i < pack.schularbeiten.length; i++) {
        const row = cleanFields(pack.schularbeiten[i]);
        if (!row.SchularbeitId) row.SchularbeitId = newEntityId('sa');
        const existing = saById.get(row.SchularbeitId);
        if (existing) {
            await patchItemFields(tok, siteId, lists[LIST_TITLES.schularbeiten].id, existing.id, row);
            saU++;
        } else {
            await createItem(tok, siteId, lists[LIST_TITLES.schularbeiten].id, row);
            saC++;
        }
        if ((saC + saU) % 10 === 0) write('  … ' + (saC + saU) + '/' + pack.schularbeiten.length);
    }
    write('  Schularbeiten +' + saC + ' / ~' + saU);

    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
        window.ms365ActionLog.append({
            tool: 'sharepoint',
            action: 'seed-schularbeiten-demo',
            target: url,
            summary:
                'Demo ' +
                pack.schoolYear +
                ': ' +
                pack.counts.schularbeiten +
                ' SA, ' +
                pack.counts.terminfenster +
                ' Fenster'
        });
    }

    write('Fertig. Stammdaten-Codes siehe Demo-Paket (Lehrer/Klassen/Fächer) – lokal in tenant.html pflegen oder JSON nutzen.');
    return {
        ok: true,
        counts: pack.counts || {
            schularbeiten: (pack.schularbeiten || []).length,
            terminfenster: (pack.terminfenster || []).length,
            fachMeta: (pack.fachMeta || []).length
        },
        created: { fachMeta: metaCreated, terminfenster: tfC, schularbeiten: saC },
        updated: { fachMeta: metaUpdated, terminfenster: tfU, schularbeiten: saU },
        stammdaten: pack.stammdaten
    };
}

/**
 * JSON-Datei (docs/demo-data/schularbeiten-2026-27.json) parsen und prüfen.
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
    if (!Array.isArray(data.schularbeiten) || !data.schularbeiten.length) {
        throw new Error('Feld „schularbeiten“ fehlt oder ist leer.');
    }
    if (!data.regelwerk || typeof data.regelwerk !== 'object') {
        throw new Error('Feld „regelwerk“ fehlt.');
    }
    if (!Array.isArray(data.terminfenster)) data.terminfenster = [];
    if (!Array.isArray(data.fachMeta)) data.fachMeta = [];
    if (!data.stammdaten) data.stammdaten = { subjects: [], classes: [], teachers: [] };
    if (!data.counts) {
        data.counts = {
            schularbeiten: data.schularbeiten.length,
            terminfenster: data.terminfenster.length,
            fachMeta: data.fachMeta.length,
            teachers: (data.stammdaten.teachers || []).length,
            classes: (data.stammdaten.classes || []).length,
            subjects: (data.stammdaten.subjects || []).length
        };
    }
    return data;
}

/**
 * Stammdaten aus Demo-Paket in tenant-settings schreiben.
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
 * Demo-Paket → Planer-State (auch ohne SharePoint nutzbar).
 * @param {object} pack
 */
export function packToLocalPlanerState(pack) {
    const p = parseDemoImportJson(pack);
    const items = (p.schularbeiten || []).map((f, i) => ({
        itemId: 'local-' + String(f.SchularbeitId || i),
        schularbeitId: String(f.SchularbeitId || ''),
        thema: String(f.Title || ''),
        fachCode: String(f.FachCode || ''),
        klasseCode: String(f.KlasseCode || ''),
        lehrerCode: String(f.LehrerCode || ''),
        lehrerEmail: String(f.LehrerEmail || '').toLowerCase(),
        datum: String(f.Datum || '').slice(0, 10),
        dauerMinuten: Number(f.DauerMinuten) || 100,
        semester: String(f.Semester || 'WS'),
        status: String(f.Status || 'beantragt').toLowerCase(),
        notiz: String(f.Notiz || ''),
        ablehnungsGrund: String(f.AblehnungsGrund || ''),
        beantragtVon: String(f.BeantragtVon || ''),
        fixiertVon: String(f.FixiertVon || ''),
        fixiertAm: f.FixiertAm ? String(f.FixiertAm) : '',
        schulterminKey: String(f.SchulterminKey || ''),
        _localOnly: true
    }));
    const windows = (p.terminfenster || []).map((f, i) => ({
        itemId: 'local-tf-' + String(f.TerminfensterId || i),
        terminfensterId: String(f.TerminfensterId || ''),
        titel: String(f.Title || ''),
        typ: f.Typ === 'erlaubt' ? 'erlaubt' : 'gesperrt',
        startdatum: String(f.Startdatum || '').slice(0, 10),
        enddatum: String(f.Enddatum || '').slice(0, 10),
        beschreibung: String(f.Beschreibung || '')
    }));
    const fachMeta = (p.fachMeta || []).map((f, i) => ({
        itemId: 'local-fm-' + String(f.FachCode || i),
        fachCode: String(f.FachCode || ''),
        name: String(f.Title || f.FachCode || ''),
        farbe: String(f.Farbe || ''),
        hatSchularbeiten: f.HatSchularbeiten !== false,
        proSemester: Number(f.ProSemester) || 2,
        standardDauer: Number(f.StandardDauer) || 100
    }));
    const rw = p.regelwerk || {};
    const rules = {
        itemId: 'local-rw',
        regelwerkId: String(rw.RegelwerkId || 'rw-demo'),
        name: String(rw.Title || 'Demo-Regelwerk'),
        maxProTag: Number(rw.MaxProTag) || 1,
        maxProWoche: Number(rw.MaxProWoche) || 2,
        ankuendigungsfristTage: Number(rw.AnkuendigungsfristTage) || 7,
        sperreVorNotenkonferenzTage: Number(rw.SperreVorNotenkonferenzTage) || 7,
        aktiv: rw.Aktiv !== false
    };
    return {
        items,
        windows,
        fachMeta,
        rules,
        stammdaten: p.stammdaten,
        siteDefault: p.siteDefault || '',
        schoolYear: p.schoolYear || '',
        counts: p.counts,
        pack: p
    };
}

export { getDemoSeedPackage };
