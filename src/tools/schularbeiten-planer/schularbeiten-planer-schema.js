/**
 * SharePoint-Listen-Schema für den Schularbeiten-Planer (MVP).
 * Stammdaten (Klassen, Lehrer, Fächer) kommen aus tenant-settings – keine Lookup-Listen.
 */

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

/** Präfix „SAP“ = Schularbeiten-Planer (eindeutig auf der SharePoint-Site). */
export const LIST_PREFIX = 'SAP';

/** Standardwerte SAP-FachMeta (UI + neue Einträge). */
export const DEFAULT_FACH_META_STANDARD_DAUER = 50;
export const DEFAULT_FACH_META_PRO_SEMESTER = 2;

export const FACH_META_COLOR_PALETTE = [
    '#6366f1',
    '#0ea5e9',
    '#10b981',
    '#f59e0b',
    '#ef4444',
    '#8b5cf6',
    '#e11d48',
    '#64748b',
    '#14b8a6',
    '#f97316',
    '#78716c'
];

export const LIST_TITLES = {
    regelwerk: 'SAP-Regelwerk',
    terminfenster: 'SAP-Terminfenster',
    schularbeiten: 'SAP-Schularbeiten',
    fachMeta: 'SAP-FachMeta'
};

/** Alte generische Namen (Migration beim Setup). */
export const LEGACY_LIST_TITLES = {
    regelwerk: 'Regelwerk',
    terminfenster: 'Terminfenster',
    schularbeiten: 'Schularbeiten',
    fachMeta: 'SA-FachMeta'
};

export const LIST_KEYS = ['regelwerk', 'terminfenster', 'schularbeiten', 'fachMeta'];

/** @type {Record<string, string>} Kurzbeschreibung für neue Listen */
export const LIST_DESCRIPTIONS = {
    regelwerk: 'Schularbeiten-Planer: Regeln (Max/Tag, Fristen) – pro Schuljahr möglich.',
    terminfenster: 'Schularbeiten-Planer: erlaubte/gesperrte Zeiträume je Schuljahr.',
    schularbeiten: 'Schularbeiten-Planer: Anträge und Termine (Codes aus Stammdaten).',
    fachMeta: 'Schularbeiten-Planer: Farben, Kontingente, Standarddauer je Fach.'
};

/**
 * @param {string} listKey
 * @returns {string[]}
 */
export function titlesForListKey(listKey) {
    const canonical = LIST_TITLES[listKey];
    const legacy = LEGACY_LIST_TITLES[listKey];
    const out = [];
    if (canonical) out.push(canonical);
    if (legacy && legacy !== canonical) out.push(legacy);
    return out;
}

/** @type {GraphColumnDef} */
export const SCHULJAHR_COLUMN = {
    name: 'Schuljahr',
    displayName: 'Schuljahr',
    text: { allowMultipleLines: false, maxLength: 12 }
};

/** @type {GraphColumnDef[]} */
export const REGELWERK_COLUMNS = [
    {
        name: 'RegelwerkId',
        displayName: 'Regelwerk-ID',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'MaxProTag',
        displayName: 'Max. pro Tag',
        number: {}
    },
    {
        name: 'MaxProWoche',
        displayName: 'Max. pro Woche',
        number: {}
    },
    {
        name: 'AnkuendigungsfristTage',
        displayName: 'Ankündigungsfrist (Tage)',
        number: {}
    },
    {
        name: 'SperreVorNotenkonferenzTage',
        displayName: 'Sperre vor Notenkonferenz (Tage)',
        number: {}
    },
    {
        name: 'Aktiv',
        displayName: 'Aktiv',
        boolean: {}
    },
    SCHULJAHR_COLUMN
];

/** @type {GraphColumnDef[]} */
export const TERMINFENSTER_COLUMNS = [
    {
        name: 'TerminfensterId',
        displayName: 'Terminfenster-ID',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'Typ',
        displayName: 'Typ',
        choice: {
            allowTextEntry: false,
            choices: ['gesperrt', 'erlaubt']
        }
    },
    {
        name: 'Startdatum',
        displayName: 'Startdatum',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Enddatum',
        displayName: 'Enddatum',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Beschreibung',
        displayName: 'Beschreibung',
        text: { allowMultipleLines: true, maxLength: 4000 }
    },
    SCHULJAHR_COLUMN
];

/** @type {GraphColumnDef[]} */
export const SCHULARBEITEN_COLUMNS = [
    {
        name: 'SchularbeitId',
        displayName: 'Schularbeit-ID',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'Titel',
        displayName: 'Titel',
        text: { allowMultipleLines: false, maxLength: 250 }
    },
    {
        name: 'Thema',
        displayName: 'Thema',
        text: { allowMultipleLines: true, maxLength: 4000 }
    },
    {
        name: 'FachCode',
        displayName: 'Fach-Code',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'KlasseCode',
        displayName: 'Klasse-Code',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'LehrerCode',
        displayName: 'Lehrer-Kürzel',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'LehrerEmail',
        displayName: 'Lehrer-E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'Datum',
        displayName: 'Datum',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'BeginnUhrzeit',
        displayName: 'Beginn (Uhrzeit)',
        text: { allowMultipleLines: false, maxLength: 5 }
    },
    {
        name: 'DauerMinuten',
        displayName: 'Dauer (Min.)',
        number: {}
    },
    {
        name: 'Semester',
        displayName: 'Semester',
        choice: {
            allowTextEntry: false,
            choices: ['WS', 'SS']
        }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: {
            allowTextEntry: false,
            choices: ['beantragt', 'fixiert', 'abgelehnt']
        }
    },
    {
        name: 'Notiz',
        displayName: 'Notiz',
        text: { allowMultipleLines: true, maxLength: 4000 }
    },
    {
        name: 'AblehnungsGrund',
        displayName: 'Ablehnungs-Grund',
        text: { allowMultipleLines: true, maxLength: 4000 }
    },
    {
        name: 'BeantragtVon',
        displayName: 'Beantragt von',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'FixiertVon',
        displayName: 'Fixiert von',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'FixiertAm',
        displayName: 'Fixiert am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'SchulterminKey',
        displayName: 'Schultermin-Key',
        text: { allowMultipleLines: false, maxLength: 80 }
    },
    {
        name: 'TeamsCalendarEventId',
        displayName: 'Teams-Kalender-Event-ID',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    SCHULJAHR_COLUMN
];

/** @type {GraphColumnDef[]} */
export const FACHMETA_COLUMNS = [
    {
        name: 'FachCode',
        displayName: 'Fach-Code',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'Farbe',
        displayName: 'Farbe',
        text: { allowMultipleLines: false, maxLength: 20 }
    },
    {
        name: 'HatSchularbeiten',
        displayName: 'Hat Schularbeiten',
        boolean: {}
    },
    {
        name: 'ProSemester',
        displayName: 'Anzahl pro Semester',
        number: {}
    },
    {
        name: 'StandardDauer',
        displayName: 'Standard-Dauer (Min.)',
        number: {}
    },
    SCHULJAHR_COLUMN
];

export const REQUIRED_COLUMNS = {
    [LIST_TITLES.regelwerk]: REGELWERK_COLUMNS.map((c) => c.name),
    [LIST_TITLES.terminfenster]: TERMINFENSTER_COLUMNS.map((c) => c.name),
    [LIST_TITLES.schularbeiten]: SCHULARBEITEN_COLUMNS.map((c) => c.name),
    [LIST_TITLES.fachMeta]: FACHMETA_COLUMNS.map((c) => c.name)
};

/** @type {Record<string, string[]>} */
export const REQUIRED_COLUMNS_BY_KEY = {
    regelwerk: REGELWERK_COLUMNS.map((c) => c.name),
    terminfenster: TERMINFENSTER_COLUMNS.map((c) => c.name),
    schularbeiten: SCHULARBEITEN_COLUMNS.map((c) => c.name),
    fachMeta: FACHMETA_COLUMNS.map((c) => c.name)
};

/** Standard-Regelwerk (Seed). */
export const DEFAULT_REGELWERK_FIELDS = {
    Title: 'Standard HAK Regelwerk',
    RegelwerkId: 'rw-1',
    MaxProTag: 1,
    MaxProWoche: 2,
    AnkuendigungsfristTage: 7,
    SperreVorNotenkonferenzTage: 7,
    Aktiv: true
};

/**
 * Graph-Spaltendefinitionen inkl. displayName für Create.
 * @param {GraphColumnDef} def
 */
export function toGraphColumnBody(def) {
    const body = { ...def };
    return body;
}

/**
 * Kurze technische ID, z. B. sa-a1b2c3d4
 * @param {string} prefix
 */
export function newEntityId(prefix) {
    const p = String(prefix || 'id').replace(/[^a-z0-9-]/gi, '').slice(0, 12) || 'id';
    const rand =
        typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
            ? crypto.randomUUID().replace(/-/g, '').slice(0, 8)
            : String(Date.now().toString(36) + Math.random().toString(36).slice(2, 8)).slice(0, 8);
    return p + '-' + rand;
}
