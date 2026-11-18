/**
 * Domänenmodell Ausgabenverwaltung (flexibles Inventar + Vorgänge).
 */

export const STORAGE_KEY = 'ms365-ausgaben-v1';
export const DEMO_SEED_TAG = 'ausgaben-demo-v1';

export const VORGANG_STATUS = {
    AUSGEGEBEN: 'ausgegeben',
    RETOUR: 'retour',
    VERLOREN: 'verloren',
    ERSETZT: 'ersetzt'
};

export const VORGANG_STATUS_LABEL = {
    ausgegeben: 'Ausgegeben',
    retour: 'Retour',
    verloren: 'Verloren',
    ersetzt: 'Ersetzt'
};

export const EMPFAENGER_TYP = {
    LEHRER: 'lehrer',
    SCHUELER: 'schueler',
    EXTERN: 'extern'
};

export const EMPFAENGER_TYP_LABEL = {
    lehrer: 'Lehrkraft',
    schueler: 'Schüler:in',
    extern: 'Extern'
};

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
    const p = String(prefix || 'ag').trim() || 'ag';
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
export function normalizeTyp(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const id = String(o.id || o.typ || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '_')
        .replace(/[^a-z0-9_\-äöüß]/gi, '');
    return {
        id: id || newId('typ'),
        label: String(o.label || o.Label || id || 'Sonstiges').trim(),
        icon: String(o.icon || 'bi-box').trim() || 'bi-box',
        hinweis: String(o.hinweis || '').trim()
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeGegenstand(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const typ = String(o.typ || o.Typ || 'sonstiges')
        .trim()
        .toLowerCase();
    return {
        id: String(o.id || o.GegenstandId || '').trim() || newId('gg'),
        typ,
        typLabel: String(o.typLabel || o.TypLabel || '').trim(),
        kennung: String(o.kennung || o.Kennung || '').trim(),
        bezeichnung: String(o.bezeichnung || o.Bezeichnung || '').trim(),
        ortZone: String(o.ortZone || o.OrtZone || '').trim(),
        aktiv: o.aktiv !== false && o.Aktiv !== false && o.Aktiv !== 'false',
        notiz: String(o.notiz || o.Notiz || '').trim(),
        quelle: String(o.quelle || o.Quelle || 'ui')
            .trim()
            .toLowerCase(),
        seedTag: String(o.seedTag || o.SeedTag || '').trim(),
        spoItemId: String(o.spoItemId || o.itemId || '').trim(),
        createdAt: String(o.createdAt || '').trim() || new Date().toISOString()
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeVorgang(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        id: String(o.id || o.VorgangId || '').trim() || newId('vg'),
        gegenstandId: String(o.gegenstandId || o.GegenstandId || '').trim(),
        kennung: String(o.kennung || o.Kennung || '').trim(),
        empfaengerTyp: String(o.empfaengerTyp || o.EmpfaengerTyp || EMPFAENGER_TYP.LEHRER)
            .trim()
            .toLowerCase(),
        empfaengerName: String(o.empfaengerName || o.EmpfaengerName || '').trim(),
        empfaengerEmail: String(o.empfaengerEmail || o.EmpfaengerEmail || '')
            .trim()
            .toLowerCase(),
        klasseCode: String(o.klasseCode || o.KlasseCode || '').trim(),
        ausgegebenAm: toIsoDateOnly(o.ausgegebenAm || o.AusgegebenAm) || toIsoDateOnly(new Date()),
        faelligAm: toIsoDateOnly(o.faelligAm || o.FaelligAm),
        zurueckAm: toIsoDateOnly(o.zurueckAm || o.ZurueckAm),
        status: String(o.status || o.Status || VORGANG_STATUS.AUSGEGEBEN)
            .trim()
            .toLowerCase(),
        kautionHinweis: String(o.kautionHinweis || o.KautionHinweis || '').trim(),
        unterschriftEmpfang: !!(
            o.unterschriftEmpfang === true ||
            o.UnterschriftEmpfang === true ||
            o.unterschriftEmpfang === 'true' ||
            o.UnterschriftEmpfang === 'true'
        ),
        notiz: String(o.notiz || o.Notiz || '').trim(),
        ausgegebenVon: String(o.ausgegebenVon || o.AusgegebenVon || '').trim(),
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
        typen: [],
        gegenstaende: [],
        vorgaenge: [],
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
        typen: Array.isArray(o.typen) ? o.typen.map(normalizeTyp) : [],
        gegenstaende: Array.isArray(o.gegenstaende) ? o.gegenstaende.map(normalizeGegenstand) : [],
        vorgaenge: Array.isArray(o.vorgaenge) ? o.vorgaenge.map(normalizeVorgang) : [],
        localDemoOnly: !!o.localDemoOnly,
        updatedAt: String(o.updatedAt || '').trim() || new Date().toISOString()
    };
}
