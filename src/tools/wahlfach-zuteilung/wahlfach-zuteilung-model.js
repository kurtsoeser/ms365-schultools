/**
 * Domänenmodell Wahlfach-Zuteilung – Normalisierung & Konstanten.
 */

export const STORAGE_KEY = 'ms365-wahlfach-zuteilung-v1';

/** @type {readonly string[]} */
export const RUNDE_STATUS = Object.freeze([
    'entwurf',
    'offen',
    'geschlossen',
    'zuteilung',
    'veroeffentlicht',
    'archiviert'
]);

/** @type {readonly string[]} */
export const ANMELDUNG_STATUS = Object.freeze([
    'eingereicht',
    'zugeordnet',
    'ersatz',
    'offen',
    'storniert'
]);

/** @type {Record<string, string>} */
export const RUNDE_STATUS_LABEL = Object.freeze({
    entwurf: 'Entwurf',
    offen: 'Offen',
    geschlossen: 'Geschlossen',
    zuteilung: 'Zuteilung',
    veroeffentlicht: 'Veröffentlicht',
    archiviert: 'Archiviert'
});

/** @type {Record<string, string>} */
export const ANMELDUNG_STATUS_LABEL = Object.freeze({
    eingereicht: 'Eingereicht',
    zugeordnet: 'Zugeordnet',
    ersatz: 'Ersatz',
    offen: 'Offen',
    storniert: 'Storniert'
});

/**
 * @param {string} [prefix]
 * @returns {string}
 */
export function newId(prefix = 'wz') {
    const p = String(prefix || 'wz').replace(/[^a-z0-9-]/gi, '').slice(0, 12) || 'wz';
    const rand =
        typeof crypto !== 'undefined' && crypto.randomUUID
            ? crypto.randomUUID().replace(/-/g, '').slice(0, 12)
            : String(Date.now().toString(36) + Math.random().toString(36).slice(2, 10));
    return p + '-' + rand;
}

/**
 * @param {unknown} v
 * @returns {string} YYYY-MM-DD oder ''
 */
export function toIsoDateOnly(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date && !Number.isNaN(v.getTime())) {
        const y = v.getFullYear();
        const m = String(v.getMonth() + 1).padStart(2, '0');
        const d = String(v.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + d;
    }
    const s = String(v).trim();
    const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : '';
}

/**
 * @param {unknown} v
 * @returns {string}
 */
export function normalizeRundeStatus(v) {
    const s = String(v || '').trim().toLowerCase();
    return RUNDE_STATUS.includes(s) ? s : 'entwurf';
}

/**
 * @param {unknown} v
 * @returns {string}
 */
export function normalizeAnmeldungStatus(v) {
    const s = String(v || '').trim().toLowerCase();
    return ANMELDUNG_STATUS.includes(s) ? s : 'eingereicht';
}

/**
 * @param {unknown} raw
 */
export function normalizeRunde(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const spoItemId = String(o.spoItemId || '').trim();
    return {
        id: String(o.id || '').trim() || newId('runde'),
        schuljahr: String(o.schuljahr || '').trim(),
        titel: String(o.titel || '').trim(),
        anmeldeVon: toIsoDateOnly(o.anmeldeVon),
        anmeldeBis: toIsoDateOnly(o.anmeldeBis),
        status: normalizeRundeStatus(o.status),
        ...(spoItemId ? { spoItemId } : {})
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeAngebot(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const max = Number(o.kapazitaetMax);
    const min = Number(o.kapazitaetMin);
    const spoItemId = String(o.spoItemId || '').trim();
    return {
        id: String(o.id || '').trim() || newId('ang'),
        bezeichnung: String(o.bezeichnung || '').trim(),
        kapazitaetMax: Number.isFinite(max) && max >= 0 ? Math.floor(max) : 0,
        kapazitaetMin: Number.isFinite(min) && min >= 0 ? Math.floor(min) : 0,
        zielJahrgaenge: String(o.zielJahrgaenge || '').trim(),
        aktiv: o.aktiv !== false && o.aktiv !== 0 && o.aktiv !== '0',
        ...(spoItemId ? { spoItemId } : {})
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeAnmeldung(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const w1 = String(o.wunsch1Id || '').trim();
    const w2 = String(o.wunsch2Id || '').trim();
    const w3 = String(o.wunsch3Id || '').trim();
    const email = String(o.schuelerEmail || o.beantragtVon || '')
        .trim()
        .toLowerCase();
    const spoItemId = String(o.spoItemId || '').trim();
    return {
        id: String(o.id || '').trim() || newId('anm'),
        schuelerName: String(o.schuelerName || '').trim(),
        klasse: String(o.klasse || '').trim(),
        ...(email ? { schuelerEmail: email } : {}),
        wunsch1Id: w1,
        wunsch2Id: w2,
        ...(w3 ? { wunsch3Id: w3 } : {}),
        eingereichtAm: String(o.eingereichtAm || '').trim() || new Date().toISOString(),
        status: normalizeAnmeldungStatus(o.status),
        ...(o.seedTag ? { seedTag: String(o.seedTag) } : {}),
        ...(spoItemId ? { spoItemId } : {})
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeZuteilung(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const angebotId = o.angebotId == null || o.angebotId === '' ? null : String(o.angebotId).trim();
    const spoItemId = String(o.spoItemId || '').trim();
    return {
        anmeldungId: String(o.anmeldungId || '').trim(),
        angebotId: angebotId || null,
        begruendung: String(o.begruendung || '').trim(),
        zugeteiltVon: String(o.zugeteiltVon || '').trim(),
        zugeteiltAm: String(o.zugeteiltAm || '').trim(),
        ...(spoItemId ? { spoItemId } : {})
    };
}

/**
 * @returns {{ runde: ReturnType<typeof normalizeRunde>|null, angebote: object[], anmeldungen: object[], zuteilungen: object[] }}
 */
export function emptyState() {
    return {
        runde: null,
        angebote: [],
        anmeldungen: [],
        zuteilungen: []
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeState(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        runde: o.runde ? normalizeRunde(o.runde) : null,
        angebote: Array.isArray(o.angebote) ? o.angebote.map(normalizeAngebot) : [],
        anmeldungen: Array.isArray(o.anmeldungen) ? o.anmeldungen.map(normalizeAnmeldung) : [],
        zuteilungen: Array.isArray(o.zuteilungen)
            ? o.zuteilungen.map(normalizeZuteilung).filter((z) => z.anmeldungId)
            : []
    };
}
