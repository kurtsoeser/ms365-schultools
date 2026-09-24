/**
 * SharePoint-Listen-Schema für den Schularbeiten-Planer (MVP).
 * Stammdaten (Klassen, Lehrer, Fächer) kommen aus tenant-settings – keine Lookup-Listen.
 */

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export const LIST_TITLES = {
    regelwerk: 'Regelwerk',
    terminfenster: 'Terminfenster',
    schularbeiten: 'Schularbeiten',
    fachMeta: 'SA-FachMeta'
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
    }
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
    }
];

/** @type {GraphColumnDef[]} */
export const SCHULARBEITEN_COLUMNS = [
    {
        name: 'SchularbeitId',
        displayName: 'Schularbeit-ID',
        text: { allowMultipleLines: false, maxLength: 40 }
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
    }
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
    }
];

export const REQUIRED_COLUMNS = {
    Regelwerk: REGELWERK_COLUMNS.map((c) => c.name),
    Terminfenster: TERMINFENSTER_COLUMNS.map((c) => c.name),
    Schularbeiten: SCHULARBEITEN_COLUMNS.map((c) => c.name),
    'SA-FachMeta': FACHMETA_COLUMNS.map((c) => c.name)
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
