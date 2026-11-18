/**
 * Domänenmodell Kollegiumsumfragen (local-first, SharePoint SoR).
 */

export const STORAGE_KEY = 'ms365-kollegium-umfragen-v1';
export const DEMO_SEED_TAG = 'kollegium-demo-v1';

export const STATUS = {
    ENTWURF: 'entwurf',
    AKTIV: 'aktiv',
    GESCHLOSSEN: 'geschlossen',
    ARCHIVIERT: 'archiviert'
};

export const STATUS_LABEL = {
    entwurf: 'Entwurf',
    aktiv: 'Aktiv',
    geschlossen: 'Geschlossen',
    archiviert: 'Archiviert'
};

export const ZIELGRUPPE = {
    ALLE: 'alle_lehrkraefte',
    KV: 'klassenvorstand',
    ABTEILUNG: 'abteilung',
    MANUELL: 'manuell'
};

export const ZIELGRUPPE_LABEL = {
    alle_lehrkraefte: 'Gesamtes Kollegium',
    klassenvorstand: 'Nur Klassenvorstände / Klassenlehrer',
    abteilung: 'Abteilung',
    manuell: 'Manuelle Liste (Mails)'
};

export const FRAGE_TYP = {
    SINGLE: 'single',
    MULTI: 'multi',
    TEXT: 'text',
    SKALA: 'skala',
    DATEI: 'datei'
};

export const FRAGE_TYP_LABEL = {
    single: 'Einzelauswahl',
    multi: 'Mehrfachauswahl',
    text: 'Freitext',
    skala: 'Skala',
    datei: 'Datei-Upload'
};

/** Max. Anhanggröße (IndexedDB). */
export const UPLOAD_MAX_BYTES = 6 * 1024 * 1024;

export const ANTWORT_STATUS = {
    OFFEN: 'offen',
    ABGEGEBEN: 'abgegeben'
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
    const p = String(prefix || 'ku').trim() || 'ku';
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
export function normalizeFrage(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const typ = String(o.typ || FRAGE_TYP.SINGLE)
        .trim()
        .toLowerCase();
    let optionen = [];
    if (Array.isArray(o.optionen)) {
        optionen = o.optionen.map((x) => String(x || '').trim()).filter(Boolean);
    } else if (typeof o.optionen === 'string') {
        optionen = o.optionen
            .split(/\r?\n|;/)
            .map((x) => x.trim())
            .filter(Boolean);
    }
    const skalaMin = Number(o.skalaMin != null ? o.skalaMin : 1) || 1;
    const skalaMax = Number(o.skalaMax != null ? o.skalaMax : 5) || 5;
    return {
        id: String(o.id || '').trim() || newId('q'),
        typ: Object.values(FRAGE_TYP).includes(typ) ? typ : FRAGE_TYP.SINGLE,
        text: String(o.text || '').trim(),
        pflicht: o.pflicht !== false,
        optionen,
        skalaMin,
        skalaMax: skalaMax < skalaMin ? skalaMin : skalaMax,
        skalaLabelMin: String(o.skalaLabelMin || '').trim(),
        skalaLabelMax: String(o.skalaLabelMax || '').trim()
    };
}

/**
 * @param {unknown} raw
 * @returns {ReturnType<typeof normalizeFrage>[]}
 */
export function normalizeFragen(raw) {
    let list = raw;
    if (typeof raw === 'string') {
        try {
            list = JSON.parse(raw);
        } catch {
            list = [];
        }
    }
    if (!Array.isArray(list)) return [];
    return list.map(normalizeFrage).filter((f) => f.text);
}

/**
 * @param {unknown} raw
 */
export function normalizeUmfrage(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const fragen = normalizeFragen(o.fragen != null ? o.fragen : o.FragenJson);
    return {
        id: String(o.id || o.UmfrageId || '').trim() || newId('ku'),
        titel: String(o.titel || o.Title || '').trim(),
        beschreibung: String(o.beschreibung || o.Beschreibung || '').trim(),
        vorlageId: String(o.vorlageId || o.VorlageId || '').trim(),
        komplexitaet: String(o.komplexitaet || o.Komplexitaet || (fragen.length > 3 ? 'komplex' : 'einfach'))
            .trim()
            .toLowerCase(),
        zielgruppe: String(o.zielgruppe || o.Zielgruppe || ZIELGRUPPE.ALLE)
            .trim()
            .toLowerCase(),
        abteilung: String(o.abteilung || o.Abteilung || '').trim(),
        klasseCodes: String(o.klasseCodes || o.KlasseCodes || '')
            .trim()
            .replace(/,/g, ';'),
        empfaengerMails: String(o.empfaengerMails || o.EmpfaengerMails || '')
            .trim()
            .replace(/,/g, ';')
            .toLowerCase(),
        fragen,
        faelligAm: toIsoDateOnly(o.faelligAm || o.FaelligAm),
        anonym: !!(o.anonym === true || o.Anonym === true || o.anonym === 'true'),
        anhangErlaubt: !!(
            o.anhangErlaubt === true ||
            o.AnhangErlaubt === true ||
            o.anhangErlaubt === 'true'
        ),
        status: String(o.status || o.Status || STATUS.ENTWURF)
            .trim()
            .toLowerCase(),
        erstellerEmail: String(o.erstellerEmail || o.ErstellerEmail || '')
            .trim()
            .toLowerCase(),
        erstellerName: String(o.erstellerName || o.ErstellerName || '').trim(),
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
export function normalizeAntwort(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    let werte = o.werte != null ? o.werte : o.AntwortenJson;
    if (typeof werte === 'string') {
        try {
            werte = JSON.parse(werte);
        } catch {
            werte = {};
        }
    }
    if (!werte || typeof werte !== 'object' || Array.isArray(werte)) werte = {};
    return {
        id: String(o.id || o.AntwortId || '').trim() || newId('ka'),
        umfrageId: String(o.umfrageId || o.UmfrageId || '').trim(),
        lehrerEmail: String(o.lehrerEmail || o.LehrerEmail || '')
            .trim()
            .toLowerCase(),
        lehrerName: String(o.lehrerName || o.LehrerName || '').trim(),
        werte: /** @type {Record<string, unknown>} */ (werte),
        anhangName: String(o.anhangName || o.AnhangName || '').trim(),
        anhangMime: String(o.anhangMime || o.AnhangMime || '').trim(),
        anhangSize: Number(o.anhangSize || 0) || 0,
        anhangKey: String(o.anhangKey || '').trim(),
        anhangUrl: String(o.anhangUrl || o.AnhangUrl || '').trim(),
        status: String(o.status || o.Status || ANTWORT_STATUS.OFFEN)
            .trim()
            .toLowerCase(),
        abgegebenAm: String(o.abgegebenAm || o.AbgegebenAm || '').trim(),
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
        umfragen: [],
        antworten: [],
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
        umfragen: Array.isArray(o.umfragen) ? o.umfragen.map(normalizeUmfrage) : [],
        antworten: Array.isArray(o.antworten) ? o.antworten.map(normalizeAntwort) : [],
        localDemoOnly: !!o.localDemoOnly,
        updatedAt: String(o.updatedAt || '').trim() || new Date().toISOString()
    };
}
