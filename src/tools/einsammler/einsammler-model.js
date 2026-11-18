/**
 * Domänenmodell Einsammler (local-first).
 */

export const STORAGE_KEY = 'ms365-einsammler-v1';

export const EINSAMMLER_STATUS = {
    ENTWURF: 'entwurf',
    AKTIV: 'aktiv',
    GESCHLOSSEN: 'geschlossen',
    ARCHIVIERT: 'archiviert'
};

export const EINSAMMLER_STATUS_LABEL = {
    entwurf: 'Entwurf',
    aktiv: 'Aktiv',
    geschlossen: 'Geschlossen',
    archiviert: 'Archiviert'
};

export const MODUS = {
    EINWILLIGUNG: 'einwilligung',
    LESEBESTAETIGUNG: 'lesebestaetigung',
    ABGABE: 'abgabe'
};

export const MODUS_LABEL = {
    einwilligung: 'Einwilligung',
    lesebestaetigung: 'Lesebestätigung',
    abgabe: 'Datei-/Bildabgabe'
};

export const TYP_LABEL = {
    exkursion: 'Exkursion',
    foto: 'Foto / Jahrbuch',
    datenschutz: 'Datenschutz',
    wandertag: 'Wandertag',
    elternbrief: 'Elternbrief',
    sonstiges: 'Sonstiges'
};

export const ANTWORT_STATUS = {
    OFFEN: 'offen',
    ZUGESTIMMT: 'zugestimmt',
    ABGELEHNT: 'abgelehnt',
    GELESEN: 'gelesen',
    ABGEGEBEN: 'abgegeben',
    ABGELAUFEN: 'abgelaufen'
};

export const ANTWORT_STATUS_LABEL = {
    offen: 'Offen',
    zugestimmt: 'Zugestimmt',
    abgelehnt: 'Abgelehnt',
    gelesen: 'Gelesen',
    abgegeben: 'Abgegeben',
    abgelaufen: 'Abgelaufen'
};

/** Max. Dateigröße für Eltern-Upload (Bytes) – lokal/API. */
export const UPLOAD_MAX_BYTES = 4 * 1024 * 1024;

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
    const p = String(prefix || 'ein').trim() || 'ein';
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
export function normalizeEinsammler(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        id: String(o.id || o.EinsammlerId || '').trim() || newId('ein'),
        titel: String(o.titel || o.Title || '').trim(),
        modus: String(o.modus || o.Modus || MODUS.EINWILLIGUNG)
            .trim()
            .toLowerCase(),
        typ: String(o.typ || o.Typ || 'sonstiges')
            .trim()
            .toLowerCase(),
        textInhalt: String(o.textInhalt || o.TextInhalt || o.text || '').trim(),
        klasseCodes: String(o.klasseCodes || o.KlasseCodes || '')
            .trim()
            .replace(/,/g, ';'),
        faelligAm: toIsoDateOnly(o.faelligAm || o.FaelligAm),
        status: String(o.status || o.Status || EINSAMMLER_STATUS.ENTWURF)
            .trim()
            .toLowerCase(),
        uploadErlaubt:
            o.uploadErlaubt === true ||
            o.UploadErlaubt === true ||
            o.uploadErlaubt === 'true' ||
            String(o.modus || o.Modus || '')
                .trim()
                .toLowerCase() === MODUS.ABGABE,
        uploadPflicht:
            o.uploadPflicht === true ||
            o.UploadPflicht === true ||
            o.uploadPflicht === 'true' ||
            String(o.modus || o.Modus || '')
                .trim()
                .toLowerCase() === MODUS.ABGABE,
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
    return {
        id: String(o.id || o.AntwortId || '').trim() || newId('ant'),
        einsammlerId: String(o.einsammlerId || o.EinsammlerId || '').trim(),
        schuelerName: String(o.schuelerName || o.SchuelerName || '').trim(),
        klasseCode: String(o.klasseCode || o.KlasseCode || '').trim(),
        elternEmail: String(o.elternEmail || o.ElternEmail || '')
            .trim()
            .toLowerCase(),
        status: String(o.status || o.Status || ANTWORT_STATUS.OFFEN)
            .trim()
            .toLowerCase(),
        antwortAm: String(o.antwortAm || o.AntwortAm || '').trim(),
        kommentar: String(o.kommentar || o.Kommentar || '').trim(),
        anhangName: String(o.anhangName || o.AnhangName || '').trim(),
        anhangMime: String(o.anhangMime || o.AnhangMime || '').trim(),
        anhangSize: Number(o.anhangSize || o.AnhangSize || 0) || 0,
        anhangUrl: String(o.anhangUrl || o.AnhangUrl || '').trim(),
        anhangKey: String(o.anhangKey || o.AnhangKey || '').trim(),
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
    try {
        // lazy: seed-tag constant lives in demo.js – compare string only here
        return String(item.seedTag || '') === 'einsammler-demo-v1';
    } catch {
        return false;
    }
}

/**
 * @returns {{ einsammler: object[], antworten: object[], localDemoOnly: boolean, updatedAt: string }}
 */
export function emptyState() {
    return {
        einsammler: [],
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
    const einsammler = Array.isArray(o.einsammler) ? o.einsammler.map(normalizeEinsammler) : [];
    const antworten = Array.isArray(o.antworten) ? o.antworten.map(normalizeAntwort) : [];
    return {
        einsammler,
        antworten,
        localDemoOnly: !!o.localDemoOnly,
        updatedAt: String(o.updatedAt || '').trim() || new Date().toISOString()
    };
}
