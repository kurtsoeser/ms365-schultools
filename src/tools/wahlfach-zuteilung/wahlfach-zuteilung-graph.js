/**
 * SharePoint-CRUD Wahlfach-Zuteilung (4 Listen = System of Record).
 * Local-first: Fehler (inkl. 403) wipe’t den lokalen Store nicht.
 * Transport: shared/graph-client.js (R22 / b11); Token/Site-Helfer weiter über spo-graph-shared.
 */
import { graphJson as sharedGraphJson } from '../../shared/graph-client.js';

import { loadSettings } from './wahlfach-zuteilung-settings.js';
import {
    normalizeAnmeldung,
    normalizeAngebot,
    normalizeRunde,
    normalizeState,
    normalizeZuteilung
} from './wahlfach-zuteilung-model.js';
import { isWzAccessDeniedError } from './wahlfach-zuteilung-logic.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

async function graphJson(method, pathOrUrl, token, body, version) {
    return sharedGraphJson(method, pathOrUrl, token, body, undefined, {
        version: version === 'beta' ? 'beta' : 'v1.0'
    });
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
    const data = await graphJson('GET', path, tok, undefined, 'v1.0');
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
        const data = await graphJson(
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

function fieldNum(fields, key) {
    const n = Number(fields && fields[key]);
    return Number.isFinite(n) ? n : 0;
}

function fieldBool(fields, key) {
    const v = fields && fields[key];
    return v === true || v === 'true' || v === 1 || v === '1' || v === 'Ja' || v === 'ja';
}

function itemsPath(ctx, listKey) {
    return (
        G().graphPathSite(ctx.siteId) +
        '/lists/' +
        encodeURIComponent(ctx.lists[listKey].id) +
        '/items'
    );
}

function itemFieldsPath(ctx, listKey, spoItemId) {
    return (
        itemsPath(ctx, listKey) + '/' + encodeURIComponent(spoItemId) + '/fields'
    );
}

async function createListItem(ctx, listKey, fields) {
    const tok = await token();
    const created = await graphJson('POST', itemsPath(ctx, listKey), tok, { fields }, 'v1.0');
    return created;
}

async function patchListItemFields(ctx, listKey, spoItemId, fields) {
    const tok = await token();
    await graphJson('PATCH', itemFieldsPath(ctx, listKey, spoItemId), tok, fields, 'v1.0');
}

async function deleteListItem(ctx, listKey, spoItemId) {
    const spoId = String(spoItemId || '').trim();
    if (!spoId) return;
    const tok = await token();
    const path = itemsPath(ctx, listKey) + '/' + encodeURIComponent(spoId);
    await graphJson('DELETE', path, tok, undefined, 'v1.0');
}

async function findItemByBusinessId(ctx, listKey, fieldName, businessId) {
    const want = String(businessId || '').trim();
    if (!want) return null;
    const tok = await token();
    const rows = await fetchAllItems(tok, ctx.siteId, ctx.lists[listKey].id);
    for (let i = 0; i < rows.length; i++) {
        const f = (rows[i] && rows[i].fields) || {};
        if (fieldStr(f, fieldName) === want) return rows[i];
    }
    return null;
}

/**
 * @param {string} webUrl
 * @param {{ listRunden?: string, listAngebote?: string, listAnmeldungen?: string, listZuteilungen?: string }} [lists]
 */
export async function resolveWzContext(webUrl, lists) {
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die SharePoint-Website-URL eintragen.');
    const cfg = lists || loadSettings();
    const tok = await token();
    const site = await G().resolveSiteFromWebUrl(tok, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    async function needList(title, label) {
        const list = await findListByTitle(tok, siteId, title);
        if (!list) {
            throw new Error(
                'Liste „' + title + '“ (' + label + ') fehlt. Bitte unter Wahlfach-Listen anlegen.'
            );
        }
        return { id: String(list.id), webUrl: list.webUrl || '', title: list.displayName || title };
    }

    const runden = await needList(cfg.listRunden || 'Wahlfach-Runden', 'Runden');
    const angebote = await needList(cfg.listAngebote || 'Wahlfach-Angebote', 'Angebote');
    const anmeldungen = await needList(cfg.listAnmeldungen || 'Wahlfach-Anmeldungen', 'Anmeldungen');
    const zuteilungen = await needList(cfg.listZuteilungen || 'Wahlfach-Zuteilungen', 'Zuteilungen');

    return {
        webUrl: url,
        siteId,
        siteName: site.displayName || '',
        lists: { runden, angebote, anmeldungen, zuteilungen }
    };
}

/** @param {object} item */
export function mapRundeFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeRunde({
        id: fieldStr(f, 'RundeId') || (item && item.id != null ? 'runde-spo-' + item.id : ''),
        schuljahr: fieldStr(f, 'Schuljahr'),
        titel: fieldStr(f, 'Title') || fieldStr(f, 'Titel'),
        anmeldeVon: fieldStr(f, 'AnmeldeVon'),
        anmeldeBis: fieldStr(f, 'AnmeldeBis'),
        status: fieldStr(f, 'Status') || 'entwurf',
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/** @param {object} runde */
export function mapRundeToFields(runde) {
    const r = normalizeRunde(runde);
    return {
        Title: r.titel || 'Wahlfach-Runde',
        RundeId: r.id,
        Schuljahr: r.schuljahr,
        AnmeldeVon: r.anmeldeVon || null,
        AnmeldeBis: r.anmeldeBis || null,
        Status: r.status || 'entwurf'
    };
}

/** @param {object} item */
export function mapAngebotFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeAngebot({
        id: fieldStr(f, 'AngebotId') || (item && item.id != null ? 'ang-spo-' + item.id : ''),
        bezeichnung: fieldStr(f, 'Title') || fieldStr(f, 'Bezeichnung'),
        kapazitaetMin: fieldNum(f, 'KapazitaetMin'),
        kapazitaetMax: fieldNum(f, 'KapazitaetMax'),
        zielJahrgaenge: fieldStr(f, 'ZielJahrgaenge'),
        aktiv: f.Aktiv == null ? true : fieldBool(f, 'Aktiv'),
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/**
 * @param {object} angebot
 * @param {{ runde?: object|null }} [meta]
 */
export function mapAngebotToFields(angebot, meta) {
    const a = normalizeAngebot(angebot);
    const runde = meta && meta.runde;
    return {
        Title: a.bezeichnung || 'Angebot',
        AngebotId: a.id,
        RundeId: (runde && runde.id) || '',
        KapazitaetMin: a.kapazitaetMin,
        KapazitaetMax: a.kapazitaetMax,
        ZielJahrgaenge: a.zielJahrgaenge,
        Aktiv: !!a.aktiv,
        Schuljahr: (runde && runde.schuljahr) || ''
    };
}

/** @param {object} item */
export function mapAnmeldungFromItem(item) {
    const f = (item && item.fields) || {};
    const createdBy =
        (item && item.createdBy && item.createdBy.user && item.createdBy.user.email) || '';
    return normalizeAnmeldung({
        id: fieldStr(f, 'AnmeldungId') || (item && item.id != null ? 'anm-spo-' + item.id : ''),
        schuelerName: fieldStr(f, 'SchuelerName') || fieldStr(f, 'Title'),
        klasse: fieldStr(f, 'Klasse'),
        schuelerEmail: fieldStr(f, 'SchuelerEmail') || String(createdBy || '').toLowerCase(),
        wunsch1Id: fieldStr(f, 'Wunsch1Id'),
        wunsch2Id: fieldStr(f, 'Wunsch2Id'),
        wunsch3Id: fieldStr(f, 'Wunsch3Id'),
        status: fieldStr(f, 'Status') || 'eingereicht',
        eingereichtAm: fieldStr(f, 'EingereichtAm') || (item && item.createdDateTime) || '',
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/**
 * @param {object} anmeldung
 * @param {{ runde?: object|null, angebote?: object[] }} [meta]
 */
export function mapAnmeldungToFields(anmeldung, meta) {
    const a = normalizeAnmeldung(anmeldung);
    const runde = meta && meta.runde;
    const angebote = (meta && meta.angebote) || [];
    const label = (id) => {
        const x = angebote.find((g) => g && g.id === id);
        return x ? x.bezeichnung : '';
    };
    return {
        Title: a.schuelerName || 'Anmeldung',
        AnmeldungId: a.id,
        RundeId: (runde && runde.id) || '',
        SchuelerName: a.schuelerName,
        Klasse: a.klasse,
        SchuelerEmail: a.schuelerEmail || '',
        Wunsch1Id: a.wunsch1Id,
        Wunsch2Id: a.wunsch2Id || '',
        Wunsch3Id: a.wunsch3Id || '',
        Wunsch1Text: label(a.wunsch1Id),
        Wunsch2Text: label(a.wunsch2Id),
        Wunsch3Text: label(a.wunsch3Id),
        Status: a.status || 'eingereicht',
        EingereichtAm: a.eingereichtAm || new Date().toISOString(),
        Schuljahr: (runde && runde.schuljahr) || ''
    };
}

/** @param {object} item */
export function mapZuteilungFromItem(item) {
    const f = (item && item.fields) || {};
    return normalizeZuteilung({
        anmeldungId: fieldStr(f, 'AnmeldungId'),
        angebotId: fieldStr(f, 'AngebotId') || null,
        begruendung: fieldStr(f, 'Begruendung'),
        zugeteiltVon: fieldStr(f, 'ZugeteiltVon'),
        zugeteiltAm: fieldStr(f, 'ZugeteiltAm'),
        spoItemId: item && item.id != null ? String(item.id) : ''
    });
}

/**
 * @param {object} zuteilung
 * @param {{ runde?: object|null, angebote?: object[], anmeldungen?: object[] }} [meta]
 */
export function mapZuteilungToFields(zuteilung, meta) {
    const z = normalizeZuteilung(zuteilung);
    const runde = meta && meta.runde;
    const angebote = (meta && meta.angebote) || [];
    const anmeldungen = (meta && meta.anmeldungen) || [];
    const ang = angebote.find((a) => a && a.id === z.angebotId);
    const anm = anmeldungen.find((a) => a && a.id === z.anmeldungId);
    const exportKey =
        ((runde && runde.schuljahr) || '') +
        '|' +
        (z.angebotId || '') +
        '|' +
        (z.anmeldungId || '');
    return {
        Title: (anm && anm.schuelerName) || z.anmeldungId || 'Zuteilung',
        ZuteilungId: z.anmeldungId,
        RundeId: (runde && runde.id) || '',
        AnmeldungId: z.anmeldungId,
        AngebotId: z.angebotId || '',
        AngebotText: ang ? ang.bezeichnung : '',
        SchuelerName: (anm && anm.schuelerName) || '',
        Klasse: (anm && anm.klasse) || '',
        Begruendung: z.begruendung || '',
        ZugeteiltVon: z.zugeteiltVon || '',
        ZugeteiltAm: z.zugeteiltAm || new Date().toISOString(),
        Schuljahr: (runde && runde.schuljahr) || '',
        AccessExportKey: exportKey
    };
}

/**
 * Vollständigen Zustand aus SharePoint laden.
 * @param {Awaited<ReturnType<typeof resolveWzContext>>} ctx
 */
export async function pullWzStateFromSharePoint(ctx) {
    const tok = await token();
    const siteId = ctx.siteId;
    const [rundeItems, angebotItems, anmeldungItems, zuteilungItems] = await Promise.all([
        fetchAllItems(tok, siteId, ctx.lists.runden.id),
        fetchAllItems(tok, siteId, ctx.lists.angebote.id),
        fetchAllItems(tok, siteId, ctx.lists.anmeldungen.id),
        fetchAllItems(tok, siteId, ctx.lists.zuteilungen.id)
    ]);

    const runden = rundeItems.map(mapRundeFromItem).filter((r) => r && r.id);
    const rank = (s) =>
        ({ offen: 4, zuteilung: 3, geschlossen: 2, veroeffentlicht: 1, entwurf: 0, archiviert: -1 })[
            String(s || '')
        ] ?? -2;
    runden.sort((a, b) => rank(b.status) - rank(a.status));
    const runde = runden[0] || null;
    const rundeId = runde && runde.id;

    let angebote = angebotItems.map(mapAngebotFromItem).filter((a) => a && a.id);
    let anmeldungen = anmeldungItems.map(mapAnmeldungFromItem).filter((a) => a && a.id);
    let zuteilungen = zuteilungItems
        .map(mapZuteilungFromItem)
        .filter((z) => z && z.anmeldungId);

    if (rundeId) {
        const keepRunde = (item, fieldKey) => {
            const f = (item && item.fields) || {};
            const rid = fieldStr(f, fieldKey);
            return !rid || rid === rundeId;
        };
        angebote = angebotItems
            .filter((it) => keepRunde(it, 'RundeId'))
            .map(mapAngebotFromItem)
            .filter((a) => a && a.id);
        anmeldungen = anmeldungItems
            .filter((it) => keepRunde(it, 'RundeId'))
            .map(mapAnmeldungFromItem)
            .filter((a) => a && a.id);
        zuteilungen = zuteilungItems
            .filter((it) => keepRunde(it, 'RundeId'))
            .map(mapZuteilungFromItem)
            .filter((z) => z && z.anmeldungId);
    }

    return normalizeState({ runde, angebote, anmeldungen, zuteilungen });
}

/**
 * @param {Awaited<ReturnType<typeof resolveWzContext>>} ctx
 * @param {object} runde
 */
export async function upsertWzRundeItem(ctx, runde) {
    const local = normalizeRunde(runde);
    const fields = mapRundeToFields(local);
    let spoId = String(local.spoItemId || '').trim();
    if (!spoId) {
        const existing = await findItemByBusinessId(ctx, 'runden', 'RundeId', local.id);
        if (existing && existing.id != null) spoId = String(existing.id);
    }
    if (spoId) {
        await patchListItemFields(ctx, 'runden', spoId, fields);
        return { ...local, spoItemId: spoId };
    }
    const created = await createListItem(ctx, 'runden', fields);
    return { ...local, spoItemId: created && created.id != null ? String(created.id) : '' };
}

/**
 * @param {Awaited<ReturnType<typeof resolveWzContext>>} ctx
 * @param {object} angebot
 * @param {{ runde?: object|null }} [meta]
 */
export async function upsertWzAngebotItem(ctx, angebot, meta) {
    const local = normalizeAngebot(angebot);
    const fields = mapAngebotToFields(local, meta);
    let spoId = String(local.spoItemId || '').trim();
    if (!spoId) {
        const existing = await findItemByBusinessId(ctx, 'angebote', 'AngebotId', local.id);
        if (existing && existing.id != null) spoId = String(existing.id);
    }
    if (spoId) {
        await patchListItemFields(ctx, 'angebote', spoId, fields);
        return { ...local, spoItemId: spoId };
    }
    const created = await createListItem(ctx, 'angebote', fields);
    return { ...local, spoItemId: created && created.id != null ? String(created.id) : '' };
}

/**
 * @param {Awaited<ReturnType<typeof resolveWzContext>>} ctx
 * @param {string} spoItemId
 */
export async function deleteWzAngebotItem(ctx, spoItemId) {
    await deleteListItem(ctx, 'angebote', spoItemId);
}

/**
 * Anmeldung nach SharePoint schreiben (Create oder Update).
 * @param {Awaited<ReturnType<typeof resolveWzContext>>} ctx
 * @param {object} anmeldung
 * @param {{ runde?: object|null, angebote?: object[] }} [meta]
 */
export async function pushWzAnmeldung(ctx, anmeldung, meta) {
    const local = normalizeAnmeldung(anmeldung);
    const fields = mapAnmeldungToFields(local, meta);
    let spoId = String(local.spoItemId || '').trim();
    if (!spoId) {
        const existing = await findItemByBusinessId(ctx, 'anmeldungen', 'AnmeldungId', local.id);
        if (existing && existing.id != null) spoId = String(existing.id);
    }
    if (spoId) {
        await patchListItemFields(ctx, 'anmeldungen', spoId, fields);
        return { ...local, spoItemId: spoId };
    }
    const created = await createListItem(ctx, 'anmeldungen', fields);
    return { ...local, spoItemId: created && created.id != null ? String(created.id) : '' };
}

/**
 * @param {Awaited<ReturnType<typeof resolveWzContext>>} ctx
 * @param {object} zuteilung
 * @param {{ runde?: object|null, angebote?: object[], anmeldungen?: object[] }} [meta]
 */
export async function upsertWzZuteilungItem(ctx, zuteilung, meta) {
    const local = normalizeZuteilung(zuteilung);
    if (!local.anmeldungId || !local.angebotId) {
        throw new Error('Zuteilung braucht Anmeldung und Angebot.');
    }
    const fields = mapZuteilungToFields(local, meta);
    let spoId = String(local.spoItemId || '').trim();
    if (!spoId) {
        const existing = await findItemByBusinessId(ctx, 'zuteilungen', 'AnmeldungId', local.anmeldungId);
        if (existing && existing.id != null) spoId = String(existing.id);
    }
    if (spoId) {
        await patchListItemFields(ctx, 'zuteilungen', spoId, fields);
        return { ...local, spoItemId: spoId };
    }
    const created = await createListItem(ctx, 'zuteilungen', fields);
    return { ...local, spoItemId: created && created.id != null ? String(created.id) : '' };
}

/**
 * @param {Awaited<ReturnType<typeof resolveWzContext>>} ctx
 * @param {string} spoItemId
 */
export async function deleteWzZuteilungItem(ctx, spoItemId) {
    await deleteListItem(ctx, 'zuteilungen', spoItemId);
}

/**
 * Soft-Sync: bei Erfolg State ersetzen; bei 403/Fehler softError, lokaler Stand bleibt.
 * @param {string} siteUrl
 * @param {{ localDemoOnly?: boolean, hasLocalData?: boolean }} [opts]
 * @returns {Promise<{ ok: boolean, soft: boolean, state: object|null, softError: string, siteId: string, error: string }>}
 */
export async function softPullWzFromSharePoint(siteUrl, opts) {
    const url = String(siteUrl || '').trim();
    const o = opts || {};
    if (!url) {
        return {
            ok: false,
            soft: true,
            state: null,
            softError: 'Keine SharePoint-Site – lokal weiterarbeiten oder unter System eine URL setzen.',
            siteId: '',
            error: ''
        };
    }
    try {
        const ctx = await resolveWzContext(url);
        const state = await pullWzStateFromSharePoint(ctx);
        return {
            ok: true,
            soft: false,
            state,
            softError: '',
            siteId: ctx.siteId,
            error: ''
        };
    } catch (e) {
        const denied = isWzAccessDeniedError(e);
        if (o.localDemoOnly || o.hasLocalData) {
            return {
                ok: false,
                soft: true,
                state: null,
                softError: denied
                    ? 'SharePoint-Zugriff eingeschränkt – lokaler Stand bleibt erhalten.'
                    : 'SharePoint nicht erreichbar – lokaler Stand bleibt erhalten.',
                siteId: '',
                error: ''
            };
        }
        if (denied) {
            return {
                ok: false,
                soft: true,
                state: null,
                softError:
                    'Kein SharePoint-Zugriff – Site unter System prüfen oder Demo lokal laden (?demo=1).',
                siteId: '',
                error: ''
            };
        }
        return {
            ok: false,
            soft: false,
            state: null,
            softError: '',
            siteId: '',
            error: (e && e.message) || String(e)
        };
    }
}
