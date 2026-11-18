/**
 * Domänenmodell Technik-Meldungen (IT-Inbox, Familie betrieb).
 */

export const STORAGE_KEY = 'ms365-technik-meldungen-v1';
export const DEMO_SEED_TAG = 'technik-demo-v1';

export const STATUS = {
    OFFEN: 'offen',
    IN_ARBEIT: 'in_arbeit',
    ERLEDIGT: 'erledigt',
    ABGELEHNT: 'abgelehnt'
};

export const STATUS_LABEL = {
    offen: 'Offen',
    in_arbeit: 'In Arbeit',
    erledigt: 'Erledigt',
    abgelehnt: 'Abgelehnt'
};

export const PRIORITAET = {
    NIEDRIG: 'niedrig',
    MITTEL: 'mittel',
    HOCH: 'hoch',
    DRINGEND: 'dringend'
};

export const PRIORITAET_LABEL = {
    niedrig: 'Niedrig',
    mittel: 'Mittel',
    hoch: 'Hoch',
    dringend: 'Dringend'
};

export const KATEGORIE = {
    PC: 'pc',
    BEAMER: 'beamer',
    AUDIO: 'audio',
    NETZWERK: 'netzwerk',
    SONSTIGES: 'sonstiges'
};

export const KATEGORIE_LABEL = {
    pc: 'PC',
    beamer: 'Beamer',
    audio: 'Audio',
    netzwerk: 'Netzwerk',
    sonstiges: 'Sonstiges'
};

export const DEMO_ME_EMAIL = 'mueller@schule.demo';
export const DEMO_ME_NAME = 'Müller';
export const DEMO_SCHUELER_EMAIL = 'schueler.demo@schule.demo';
export const DEMO_SCHUELER_NAME = 'Demo Schüler';

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
 * @param {string} [prefix]
 * @returns {string}
 */
export function newId(prefix) {
    const p = String(prefix || 'tm').trim() || 'tm';
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
    const kategorie = String(o.kategorie || o.Kategorie || KATEGORIE.SONSTIGES)
        .trim()
        .toLowerCase();
    const prioritaet = String(o.prioritaet || o.Prioritaet || PRIORITAET.MITTEL)
        .trim()
        .toLowerCase();
    const status = String(o.status || o.Status || STATUS.OFFEN)
        .trim()
        .toLowerCase();
    return {
        id: String(o.id || o.MeldungId || '').trim() || newId('tm'),
        title: String(o.title || o.Title || o.titel || '').trim(),
        beschreibung: String(o.beschreibung || o.Beschreibung || '').trim(),
        raumCode: String(o.raumCode || o.RaumCode || '').trim(),
        raumName: String(o.raumName || o.RaumName || '').trim(),
        kategorie: Object.values(KATEGORIE).includes(kategorie) ? kategorie : KATEGORIE.SONSTIGES,
        prioritaet: Object.values(PRIORITAET).includes(prioritaet) ? prioritaet : PRIORITAET.MITTEL,
        status: Object.values(STATUS).includes(status) ? status : STATUS.OFFEN,
        melderName: String(o.melderName || o.MelderName || '').trim(),
        melderEmail: String(o.melderEmail || o.MelderEmail || '')
            .trim()
            .toLowerCase(),
        gemeldetAm: String(o.gemeldetAm || o.GemeldetAm || '').trim() || new Date().toISOString(),
        erledigtAm: toIsoDateOnly(o.erledigtAm || o.ErledigtAm),
        kommentar: String(o.kommentar || o.Kommentar || '').trim(),
        bearbeitetVonName: String(o.bearbeitetVonName || o.BearbeitetVonName || '').trim(),
        bearbeitetVonEmail: String(o.bearbeitetVonEmail || o.BearbeitetVonEmail || '')
            .trim()
            .toLowerCase(),
        bearbeitetAm: String(o.bearbeitetAm || o.BearbeitetAm || '').trim(),
        seedTag: String(o.seedTag || o.SeedTag || '').trim(),
        spoItemId: String(o.spoItemId || o.itemId || '').trim()
    };
}

/**
 * @param {unknown} item
 */
export function isDemoSeedItem(item) {
    return String((item && item.seedTag) || '') === DEMO_SEED_TAG;
}

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
