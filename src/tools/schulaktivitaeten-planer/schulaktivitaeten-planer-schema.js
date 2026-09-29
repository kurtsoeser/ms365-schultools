/**
 * SharePoint-Listen-Schema: Schulaktivitäten / Exkursionen.
 * Stammdaten (Klassen, Lehrer) aus tenant-settings – keine Lookup-Listen.
 */

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export const LIST_TITLES = {
    aktivitaeten: 'Schulaktivitaeten',
    regelwerk: 'Aktivitaet-Regelwerk'
};

export const TYP_CHOICES = ['Exkursion', 'Schulaktivitaet', 'Veranstaltung', 'Sonstiges'];
export const STATUS_CHOICES = ['beantragt', 'genehmigt', 'abgelehnt'];

/** @type {GraphColumnDef[]} */
export const AKTIVITAETEN_COLUMNS = [
    {
        name: 'AktivitaetId',
        displayName: 'Aktivitäts-ID',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'Typ',
        displayName: 'Typ',
        choice: { allowTextEntry: false, choices: [...TYP_CHOICES] }
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
        name: 'Begleitung',
        displayName: 'Begleitung',
        text: { allowMultipleLines: true, maxLength: 4000 }
    },
    {
        name: 'Ort',
        displayName: 'Ort / Ziel',
        text: { allowMultipleLines: false, maxLength: 255 }
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
        name: 'StartZeit',
        displayName: 'Startzeit',
        text: { allowMultipleLines: false, maxLength: 10 }
    },
    {
        name: 'EndZeit',
        displayName: 'Endzeit',
        text: { allowMultipleLines: false, maxLength: 10 }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: [...STATUS_CHOICES] }
    },
    {
        name: 'Notiz',
        displayName: 'Notiz',
        text: { allowMultipleLines: true, maxLength: 4000 }
    },
    {
        name: 'AblehnungsGrund',
        displayName: 'Ablehnungsgrund',
        text: { allowMultipleLines: true, maxLength: 4000 }
    },
    {
        name: 'BeantragtVon',
        displayName: 'Beantragt von',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'GenehmigtVon',
        displayName: 'Genehmigt von',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'GenehmigtAm',
        displayName: 'Genehmigt am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'Verkehrsmittel',
        displayName: 'Verkehrsmittel',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'KostenHinweis',
        displayName: 'Kostenhinweis',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'SchulterminKey',
        displayName: 'Schultermin-Key',
        text: { allowMultipleLines: false, maxLength: 80 }
    }
];

/** @type {GraphColumnDef[]} */
export const REGELWERK_COLUMNS = [
    {
        name: 'RegelwerkId',
        displayName: 'Regelwerk-ID',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'MinVorlaufTage',
        displayName: 'Min. Vorlauf (Tage)',
        number: {}
    },
    {
        name: 'MaxGleichzeitigProKlasse',
        displayName: 'Max. gleichzeitig pro Klasse',
        number: {}
    },
    {
        name: 'Aktiv',
        displayName: 'Aktiv',
        boolean: {}
    }
];

export const REQUIRED_COLUMNS = {
    Schulaktivitaeten: AKTIVITAETEN_COLUMNS.map((c) => c.name),
    'Aktivitaet-Regelwerk': REGELWERK_COLUMNS.map((c) => c.name)
};

export const DEFAULT_REGELWERK_FIELDS = {
    Title: 'Standard Schulaktivitäten',
    RegelwerkId: 'akt-rw-1',
    MinVorlaufTage: 7,
    MaxGleichzeitigProKlasse: 1,
    Aktiv: true
};

/**
 * @param {GraphColumnDef} def
 */
export function toGraphColumnBody(def) {
    const body = {
        name: def.name,
        displayName: def.displayName || def.name
    };
    if (def.text) body.text = def.text;
    if (def.number) body.number = def.number;
    if (def.boolean) body.boolean = def.boolean;
    if (def.dateTime) body.dateTime = def.dateTime;
    if (def.choice) body.choice = def.choice;
    return body;
}

export function newEntityId(prefix) {
    const p = String(prefix || 'akt').replace(/[^a-z0-9-]/gi, '') || 'akt';
    const rand =
        typeof crypto !== 'undefined' && crypto.randomUUID
            ? crypto.randomUUID().replace(/-/g, '').slice(0, 10)
            : String(Date.now()).slice(-10);
    return p + '-' + rand;
}
