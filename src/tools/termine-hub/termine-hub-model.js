/**
 * Domänenmodell Kalender und Termine (lokale Termine + Sync-Status).
 */

export const STORAGE_KEY = 'ms365-termine-hub-v1';

export const SICHTBARKEIT = {
    OEFFENTLICH: 'oeffentlich',
    INTERN: 'intern'
};

export const SICHTBARKEIT_LABEL = {
    oeffentlich: 'Öffentlich',
    intern: 'Intern'
};

/** Fachlicher Termin-Typ – getrennt von Sichtbarkeit (wer darf sehen). */
export const TERMIN_TYP = {
    KONFERENZ: 'konferenz',
    FORTBILDUNG: 'fortbildung',
    VERANSTALTUNG: 'veranstaltung',
    PRUEFUNG: 'pruefung',
    ORGANISATION: 'organisation',
    ARGE: 'arge',
    BERATUNG: 'beratung',
    SONSTIGES: 'sonstiges'
};

export const TERMIN_TYP_LABEL = {
    konferenz: 'Konferenz',
    fortbildung: 'Fortbildung',
    veranstaltung: 'Veranstaltung',
    pruefung: 'Prüfung',
    organisation: 'Organisation',
    arge: 'ARGE',
    beratung: 'Beratung',
    sonstiges: 'Sonstiges'
};

/** Legacy-Wert „intern“ war Sichtbarkeit, nicht Typ. */
const TERMIN_TYP_ALIASES = {
    intern: TERMIN_TYP.ORGANISATION
};

export const SYNC_STATUS = {
    LOKAL: 'lokal',
    PENDING: 'pending',
    SYNCED: 'synced'
};

export const SYNC_STATUS_LABEL = {
    lokal: 'Nur lokal',
    pending: 'Für Outlook vorgemerkt',
    synced: 'Mit Outlook synchron'
};

/**
 * Lokales Kalenderdatum YYYY-MM-DD (ohne UTC-Verschiebung).
 * @param {unknown} value
 * @returns {string}
 */
export function toLocalDateOnly(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        const y = value.getFullYear();
        const m = String(value.getMonth() + 1).padStart(2, '0');
        const d = String(value.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + d;
    }
    const s = String(value || '').trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const d = new Date(s);
    if (Number.isNaN(d.getTime())) return '';
    return toLocalDateOnly(d);
}

/**
 * @param {unknown} value
 * @returns {string}
 */
export function toIsoDateTime(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return value.toISOString();
    }
    const s = String(value || '').trim();
    if (!s) return '';
    if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
        const d = new Date(s);
        return Number.isNaN(d.getTime()) ? s : d.toISOString();
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
        // Mittag lokal → stabile ganztägige Speicherung
        const d = new Date(s + 'T12:00:00');
        return Number.isNaN(d.getTime()) ? s + 'T12:00:00.000Z' : d.toISOString();
    }
    const d = new Date(s);
    return Number.isNaN(d.getTime()) ? '' : d.toISOString();
}

/**
 * @param {unknown} value
 * @returns {string} YYYY-MM-DD
 */
export function toIsoDateOnly(value) {
    return toLocalDateOnly(value);
}

/**
 * @param {string} [prefix]
 * @returns {string}
 */
export function newId(prefix) {
    const p = String(prefix || 'th').trim() || 'th';
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
export function normalizeTermin(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    let sichtbarkeit = String(o.sichtbarkeit || o.Sichtbarkeit || SICHTBARKEIT.INTERN)
        .trim()
        .toLowerCase();
    if (sichtbarkeit === 'öffentlich' || sichtbarkeit === 'oeffentlich') {
        sichtbarkeit = SICHTBARKEIT.OEFFENTLICH;
    }
    let terminTyp = String(o.terminTyp || o.TerminTyp || TERMIN_TYP.ORGANISATION)
        .trim()
        .toLowerCase();
    if (TERMIN_TYP_ALIASES[terminTyp]) terminTyp = TERMIN_TYP_ALIASES[terminTyp];
    const syncStatus = String(o.syncStatus || o.SyncStatus || SYNC_STATUS.LOKAL)
        .trim()
        .toLowerCase();
    const allDay = !!(o.allDay === true || o.AllDay === true || o.allDay === 'true');
    let start = toIsoDateTime(o.start || o.Beginn || o.Start);
    let end = toIsoDateTime(o.end || o.Ende || o.End);
    if (allDay && start) {
        const day = toLocalDateOnly(o.start || o.Beginn || o.Start || start);
        start = day ? new Date(day + 'T12:00:00').toISOString() : start;
    }
    if (allDay && end) {
        const day = toLocalDateOnly(o.end || o.Ende || o.End || end);
        end = day ? new Date(day + 'T12:00:00').toISOString() : end;
    }
    if (!end && start) end = start;

    return {
        id: String(o.id || o.TerminId || '').trim() || newId('th'),
        title: String(o.title || o.Title || o.titel || '').trim(),
        start,
        end,
        allDay,
        ort: String(o.ort || o.Ort || '').trim(),
        beschreibung: String(o.beschreibung || o.Beschreibung || o.Info || '').trim(),
        sichtbarkeit: Object.values(SICHTBARKEIT).includes(sichtbarkeit)
            ? sichtbarkeit
            : SICHTBARKEIT.INTERN,
        terminTyp: Object.values(TERMIN_TYP).includes(terminTyp) ? terminTyp : TERMIN_TYP.SONSTIGES,
        zielgruppe: String(o.zielgruppe || o.Zielgruppe || '').trim(),
        graphEventId: String(o.graphEventId || o.GraphEventId || o.OutlookEventID || '').trim(),
        syncStatus: Object.values(SYNC_STATUS).includes(syncStatus) ? syncStatus : SYNC_STATUS.LOKAL,
        createdAt: String(o.createdAt || '').trim() || new Date().toISOString(),
        updatedAt: String(o.updatedAt || '').trim() || new Date().toISOString()
    };
}

export function emptyState() {
    return {
        termine: [],
        updatedAt: new Date().toISOString()
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeState(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        termine: Array.isArray(o.termine) ? o.termine.map(normalizeTermin) : [],
        updatedAt: String(o.updatedAt || '').trim() || new Date().toISOString()
    };
}
