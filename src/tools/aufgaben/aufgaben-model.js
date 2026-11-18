/**
 * Domänenmodell Aufgaben (interne Zuweisungen, Familie betrieb).
 */

export const STORAGE_KEY = 'ms365-aufgaben-v1';
export const DEMO_SEED_TAG = 'aufgaben-demo-v1';

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

export const ROLLE = {
    DIREKTION: 'direktion',
    IT: 'it',
    KV: 'kv',
    LEHRER: 'lehrer',
    VERWALTUNG: 'verwaltung'
};

export const ROLLE_LABEL = {
    direktion: 'Direktion',
    it: 'IT',
    kv: 'Klassenvorstand',
    lehrer: 'Lehrkraft',
    verwaltung: 'Verwaltung'
};

/** Demo-„Ich“ für Filter Mir / Von mir */
export const DEMO_ME_EMAIL = 'mueller@schule.demo';
export const DEMO_ME_NAME = 'Müller';

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
    const p = String(prefix || 'af').trim() || 'af';
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
export function normalizeAufgabe(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const rolle = String(o.rolle || o.Rolle || '')
        .trim()
        .toLowerCase();
    const prioritaet = String(o.prioritaet || o.Prioritaet || PRIORITAET.MITTEL)
        .trim()
        .toLowerCase();
    const status = String(o.status || o.Status || STATUS.OFFEN)
        .trim()
        .toLowerCase();
    return {
        id: String(o.id || o.AufgabeId || '').trim() || newId('af'),
        title: String(o.title || o.Title || o.titel || '').trim(),
        beschreibung: String(o.beschreibung || o.Beschreibung || '').trim(),
        zuweisungName: String(o.zuweisungName || o.ZuweisungName || '').trim(),
        zuweisungEmail: String(o.zuweisungEmail || o.ZuweisungEmail || '')
            .trim()
            .toLowerCase(),
        rolle: Object.values(ROLLE).includes(rolle) ? rolle : '',
        prioritaet: Object.values(PRIORITAET).includes(prioritaet) ? prioritaet : PRIORITAET.MITTEL,
        frist: toIsoDateOnly(o.frist || o.Frist),
        status: Object.values(STATUS).includes(status) ? status : STATUS.OFFEN,
        erstelltVon: String(o.erstelltVon || o.ErstelltVon || '').trim(),
        erstelltVonEmail: String(o.erstelltVonEmail || o.ErstelltVonEmail || '')
            .trim()
            .toLowerCase(),
        erstelltAm: String(o.erstelltAm || o.ErstelltAm || '').trim() || new Date().toISOString(),
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
        aufgaben: [],
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
        aufgaben: Array.isArray(o.aufgaben) ? o.aufgaben.map(normalizeAufgabe) : [],
        localDemoOnly: !!o.localDemoOnly,
        updatedAt: String(o.updatedAt || '').trim() || new Date().toISOString()
    };
}
