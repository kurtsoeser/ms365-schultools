/**
 * SharePoint-Listen-Schema: Schulaktivitäten / Exkursionen.
 * Stammdaten (Klassen, Lehrer): Zielbild SharePoint-Listen der Schule
 * (siehe school-stammdaten-spo-lists.js); Browser-Cache/tenant-settings als Arbeitskopie.
 *
 * Personenbeziehungen (kein zweites Listenobjekt):
 * - BeantragtVon      → Antragsteller (E-Mail)
 * - LehrerCode/Email  → verantwortliche Lehrkraft
 * - BegleitungCodes   → Begleitpersonen (Kürzel, ;-getrennt)
 * - Begleitung        → Anzeigenamen der Begleitung (lesbar in SharePoint)
 * KV-Zuordnung Person↔Klasse liegt in den Stammdaten (headEmail), nicht hier.
 */

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export const LIST_TITLES = {
    aktivitaeten: 'Schulaktivitaeten',
    regelwerk: 'Aktivitaet-Regelwerk'
};

export const TYP_CHOICES = [
    'Exkursion',
    'Schulaktivitaet',
    'Veranstaltung',
    'Klassenabwesenheit',
    'Sonstiges'
];

/**
 * Fachliche Statuscodes (SharePoint Choice).
 * Anzeige-Labels: siehe statusLabel() – ruhige Lehrkraftsprache.
 * entwurf: lokal / noch nicht eingereicht
 * beantragt: eingereicht, wartet auf Direktion
 * kv_stellungnahme: in Prüfung beim KV
 * nachforderung: Ergänzung erforderlich
 * abgesagt: zurückgezogen nach Genehmigung / Planung
 */
export const STATUS_CHOICES = [
    'entwurf',
    'beantragt',
    'kv_stellungnahme',
    'nachforderung',
    'genehmigt',
    'abgelehnt',
    'abgesagt'
];

/** Genehmigungspfad je Schule (Regelwerk). */
export const GENEHMIGUNG_PFAD_CHOICES = ['direktion', 'kv_dann_direktion'];

/** Vorlauf-Enforcement: sperre | warnung | ausnahme (Begründung). */
export const VORLAUF_MODUS_CHOICES = ['sperre', 'warnung', 'ausnahme'];

/** Wie Teilnehmer:innen ausgewählt werden. */
export const TEILNAHME_MODUS_CHOICES = ['klasse', 'klassen', 'gruppe', 'schueler'];

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
        name: 'KlasseCodes',
        displayName: 'Klassen-Codes',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'TeilnahmeModus',
        displayName: 'Teilnahme-Modus',
        choice: { allowTextEntry: false, choices: [...TEILNAHME_MODUS_CHOICES] }
    },
    {
        name: 'GruppenText',
        displayName: 'Unterrichtsgruppe',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'TeilnehmendeText',
        displayName: 'Teilnehmende',
        text: { allowMultipleLines: true, maxLength: 4000 }
    },
    {
        name: 'NichtTeilnehmende',
        displayName: 'Nicht teilnehmend',
        text: { allowMultipleLines: true, maxLength: 4000 }
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
        name: 'BegleitungCodes',
        displayName: 'Begleitung-Codes',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'ExterneVortragende',
        displayName: 'Externe Vortragende',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'Ort',
        displayName: 'Ort / Ziel',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'Ablauf',
        displayName: 'Ablaufbeschreibung',
        text: { allowMultipleLines: true, maxLength: 4000 }
    },
    {
        name: 'Unterrichtsbezug',
        displayName: 'Unterrichtsbezug',
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
        name: 'StartUe',
        displayName: 'Start-UE',
        number: {}
    },
    {
        name: 'EndUe',
        displayName: 'End-UE',
        number: {}
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
        name: 'RueckfrageText',
        displayName: 'Rückfrage / Nachforderung',
        text: { allowMultipleLines: true, maxLength: 4000 }
    },
    {
        name: 'KvStellungnahme',
        displayName: 'KV-Stellungnahme',
        text: { allowMultipleLines: true, maxLength: 4000 }
    },
    {
        name: 'KvStellungnahmeVon',
        displayName: 'KV-Stellungnahme von',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'KvStellungnahmeAm',
        displayName: 'KV-Stellungnahme am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'AenderungsLog',
        displayName: 'Änderungsprotokoll',
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
        name: 'UntisErledigt',
        displayName: 'In Untis verarbeitet',
        boolean: {}
    },
    {
        name: 'UntisErledigtVon',
        displayName: 'Untis erledigt von',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'UntisErledigtAm',
        displayName: 'Untis erledigt am',
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
        name: 'GenehmigungPfad',
        displayName: 'Genehmigungspfad',
        choice: { allowTextEntry: false, choices: [...GENEHMIGUNG_PFAD_CHOICES] }
    },
    {
        name: 'VorlaufModus',
        displayName: 'Vorlauf-Modus',
        choice: { allowTextEntry: false, choices: [...VORLAUF_MODUS_CHOICES] }
    },
    {
        name: 'Aktiv',
        displayName: 'Aktiv',
        boolean: {}
    },
    {
        name: 'KostenLimitsJson',
        displayName: 'Kostenlimits (JSON)',
        text: { allowMultipleLines: true, maxLength: 8000 }
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
    GenehmigungPfad: 'direktion',
    VorlaufModus: 'sperre',
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
