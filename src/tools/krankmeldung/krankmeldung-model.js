/**
 * Domänenmodell Krankmeldung (local-first; SoR = SharePoint).
 */

export const STORAGE_KEY = 'ms365-krankmeldung-v1';

/** Seed-Tag für Demo-Zeilen (nicht nach SPO pushen). */
export const DEMO_SEED_TAG = 'krankmeldung-demo-v1';

export const STATUS = {
    NEU: 'neu',
    ZUR_KENNTNIS: 'zur_kenntnis',
    ARCHIVIERT: 'archiviert'
};

export const STATUS_LABEL = {
    neu: 'Neu',
    zur_kenntnis: 'Zur Kenntnis',
    archiviert: 'Archiviert'
};

export const KATEGORIE = {
    KRANKHEIT: 'krankheit',
    ARZT: 'arzt',
    SONSTIGES: 'sonstiges'
};

export const KATEGORIE_LABEL = {
    krankheit: 'Krankheit',
    arzt: 'Arzttermin',
    sonstiges: 'Sonstiges'
};

/** Max. Anhanggröße (lokal IndexedDB). */
export const UPLOAD_MAX_BYTES = 6 * 1024 * 1024;

/** Max. Meldezeitraum in Tagen (Sonderfall: Tippfehler / Missbrauch). */
export const MAX_RANGE_DAYS = 30;

/** Von-Datum darf höchstens so viele Tage in der Vergangenheit liegen. */
export const MAX_PAST_DAYS = 14;

/** Von-Datum darf höchstens so viele Tage in der Zukunft liegen. */
export const MAX_FUTURE_DAYS = 7;

/**
 * @param {unknown} value
 * @returns {string}
 */
export function toIsoDateOnly(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return value.toISOString().slice(0, 10);
    }
    const s = String(value || '').trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    return '';
}

/**
 * @returns {string}
 */
export function newId(prefix) {
    const p = String(prefix || 'km').trim() || 'km';
    try {
        if (typeof crypto !== 'undefined' && crypto.randomUUID) {
            return p + '-' + crypto.randomUUID().replace(/-/g, '').slice(0, 12);
        }
    } catch {
        /* ignore */
    }
    return p + '-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

/**
 * @param {unknown} raw
 */
export function normalizeMeldung(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const von = toIsoDateOnly(o.von || o.Von);
    let bis = toIsoDateOnly(o.bis || o.Bis) || von;
    if (von && bis && bis < von) bis = von;
    return {
        id: String(o.id || o.MeldungId || '').trim() || newId('km'),
        schuelerName: String(o.schuelerName || o.SchuelerName || '').trim(),
        klasseCode: String(o.klasseCode || o.KlasseCode || '').trim(),
        von,
        bis,
        kategorie: String(o.kategorie || o.Kategorie || KATEGORIE.KRANKHEIT)
            .trim()
            .toLowerCase(),
        gemeldetAm: String(o.gemeldetAm || o.GemeldetAm || '').trim() || new Date().toISOString(),
        gemeldetVonEmail: String(o.gemeldetVonEmail || o.GemeldetVonEmail || '')
            .trim()
            .toLowerCase(),
        gemeldetVonName: String(o.gemeldetVonName || o.GemeldetVonName || '').trim(),
        status: String(o.status || o.Status || STATUS.NEU)
            .trim()
            .toLowerCase(),
        kenntnisLehrerEmail: String(o.kenntnisLehrerEmail || o.KenntnisLehrerEmail || '')
            .trim()
            .toLowerCase(),
        notiz: String(o.notiz || o.Notiz || '').trim(),
        anhangName: String(o.anhangName || o.AnhangName || '').trim(),
        anhangMime: String(o.anhangMime || o.AnhangMime || '').trim(),
        anhangSize: Number(o.anhangSize || 0) || 0,
        anhangKey: String(o.anhangKey || '').trim(),
        anhangUrl: String(o.anhangUrl || o.AnhangUrl || '').trim(),
        quelle: String(o.quelle || o.Quelle || 'ui')
            .trim()
            .toLowerCase(),
        seedTag: String(o.seedTag || o.SeedTag || '').trim(),
        spoItemId: String(o.spoItemId || o.itemId || '').trim()
    };
}

/**
 * @param {unknown} item
 */
export function isDemoSeedItem(item) {
    if (!item || typeof item !== 'object') return false;
    if (String(item.quelle || '').toLowerCase() === 'demo') return true;
    return String(item.seedTag || '') === DEMO_SEED_TAG;
}

/**
 * @returns {{ meldungen: object[], localDemoOnly: boolean, updatedAt: string }}
 */
export function emptyState() {
    return {
        meldungen: [],
        localDemoOnly: false,
        updatedAt: new Date().toISOString()
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeState(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        meldungen: Array.isArray(o.meldungen) ? o.meldungen.map(normalizeMeldung) : [],
        localDemoOnly: !!o.localDemoOnly,
        updatedAt: String(o.updatedAt || '').trim() || new Date().toISOString()
    };
}
