/**
 * SharePoint-Listen: Ausgabenverwaltung (Inventar + Vorgänge).
 * Typen sind bewusst Text (nicht starre Choice) – Schulen erweitern selbst.
 */
import { toGraphColumnBody } from '../../shared/spo-graph-column.js';

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export { toGraphColumnBody };

export const LIST_TITLES = {
    gegenstaende: 'AusgabenGegenstaende',
    vorgaenge: 'AusgabenVorgaenge'
};

export const EMPFAENGER_TYP_CHOICES = ['lehrer', 'schueler', 'extern'];

export const VORGANG_STATUS_CHOICES = ['ausgegeben', 'retour', 'verloren', 'ersetzt'];

/** @type {GraphColumnDef[]} */
export const GEGENSTAENDE_COLUMNS = [
    {
        name: 'GegenstandId',
        displayName: 'Gegenstand-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Typ',
        displayName: 'Typ',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'TypLabel',
        displayName: 'Typ-Bezeichnung',
        text: { allowMultipleLines: false, maxLength: 80 }
    },
    {
        name: 'Kennung',
        displayName: 'Kennung',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'Bezeichnung',
        displayName: 'Bezeichnung',
        text: { allowMultipleLines: false, maxLength: 160 }
    },
    {
        name: 'OrtZone',
        displayName: 'Ort / Zone',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'Aktiv',
        displayName: 'Aktiv',
        boolean: {}
    },
    {
        name: 'Notiz',
        displayName: 'Notiz',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'Quelle',
        displayName: 'Quelle',
        choice: { allowTextEntry: false, choices: ['ui', 'vorlage', 'import'] }
    }
];

/** @type {GraphColumnDef[]} */
export const VORGAENGE_COLUMNS = [
    {
        name: 'VorgangId',
        displayName: 'Vorgang-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'GegenstandId',
        displayName: 'Gegenstand-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Kennung',
        displayName: 'Kennung',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'EmpfaengerTyp',
        displayName: 'Empfänger-Typ',
        choice: { allowTextEntry: false, choices: [...EMPFAENGER_TYP_CHOICES] }
    },
    {
        name: 'EmpfaengerName',
        displayName: 'Empfänger-Name',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'EmpfaengerEmail',
        displayName: 'Empfänger-E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'KlasseCode',
        displayName: 'Klasse-Code',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'AusgegebenAm',
        displayName: 'Ausgegeben am',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'FaelligAm',
        displayName: 'Fällig am',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'ZurueckAm',
        displayName: 'Zurück am',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: [...VORGANG_STATUS_CHOICES] }
    },
    {
        name: 'KautionHinweis',
        displayName: 'Kaution (Hinweis)',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'UnterschriftEmpfang',
        displayName: 'Unterschrift Empfang',
        boolean: {}
    },
    {
        name: 'Notiz',
        displayName: 'Notiz',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'AusgegebenVon',
        displayName: 'Ausgegeben von',
        text: { allowMultipleLines: false, maxLength: 120 }
    }
];

export const REQUIRED_COLUMNS = {
    [LIST_TITLES.gegenstaende]: GEGENSTAENDE_COLUMNS.map((c) => c.name),
    [LIST_TITLES.vorgaenge]: VORGAENGE_COLUMNS.map((c) => c.name)
};

