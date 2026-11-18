/**
 * SharePoint-Listen: Kollegiumsumfragen (Köpfe + Antworten).
 * SoR bewusst SharePoint – nicht Microsoft Forms (Forms bleibt optionaler Export später).
 */
import { toGraphColumnBody } from '../../shared/spo-graph-column.js';

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export { toGraphColumnBody };

export const LIST_TITLES = {
    umfragen: 'KollegiumUmfragen',
    antworten: 'KollegiumUmfrageAntworten'
};

export const STATUS_CHOICES = ['entwurf', 'aktiv', 'geschlossen', 'archiviert'];

export const ZIELGRUPPE_CHOICES = ['alle_lehrkraefte', 'klassenvorstand', 'abteilung', 'manuell'];

export const KOMPLEXITAET_CHOICES = ['einfach', 'komplex'];

/** @type {GraphColumnDef[]} */
export const UMFRAGEN_COLUMNS = [
    {
        name: 'UmfrageId',
        displayName: 'Umfrage-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Beschreibung',
        displayName: 'Beschreibung',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'VorlageId',
        displayName: 'Vorlage-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Komplexitaet',
        displayName: 'Komplexität',
        choice: { allowTextEntry: false, choices: [...KOMPLEXITAET_CHOICES] }
    },
    {
        name: 'Zielgruppe',
        displayName: 'Zielgruppe',
        choice: { allowTextEntry: false, choices: [...ZIELGRUPPE_CHOICES] }
    },
    {
        name: 'Abteilung',
        displayName: 'Abteilung',
        text: { allowMultipleLines: false, maxLength: 80 }
    },
    {
        name: 'KlasseCodes',
        displayName: 'Klassen-Codes',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'EmpfaengerMails',
        displayName: 'Empfänger-Mails',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'FragenJson',
        displayName: 'Fragen (JSON)',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'FaelligAm',
        displayName: 'Fällig am',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Anonym',
        displayName: 'Anonym auswerten',
        boolean: {}
    },
    {
        name: 'AnhangErlaubt',
        displayName: 'Anhang erlaubt',
        boolean: {}
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: [...STATUS_CHOICES] }
    },
    {
        name: 'ErstellerEmail',
        displayName: 'Ersteller-E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'ErstellerName',
        displayName: 'Ersteller',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'Quelle',
        displayName: 'Quelle',
        choice: { allowTextEntry: false, choices: ['ui', 'vorlage', 'import'] }
    },
    {
        name: 'SeedTag',
        displayName: 'Seed-Tag',
        text: { allowMultipleLines: false, maxLength: 60 }
    }
];

/** @type {GraphColumnDef[]} */
export const ANTWORTEN_COLUMNS = [
    {
        name: 'AntwortId',
        displayName: 'Antwort-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'UmfrageId',
        displayName: 'Umfrage-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'LehrerEmail',
        displayName: 'Lehrkraft-E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'LehrerName',
        displayName: 'Lehrkraft',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'AntwortenJson',
        displayName: 'Antworten (JSON)',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: ['offen', 'abgegeben'] }
    },
    {
        name: 'AbgegebenAm',
        displayName: 'Abgegeben am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'AnhangName',
        displayName: 'Anhang-Name',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'AnhangMime',
        displayName: 'Anhang-MIME',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'AnhangUrl',
        displayName: 'Anhang-URL',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'SeedTag',
        displayName: 'Seed-Tag',
        text: { allowMultipleLines: false, maxLength: 60 }
    }
];

export const REQUIRED_COLUMNS = {
    [LIST_TITLES.umfragen]: UMFRAGEN_COLUMNS.map((c) => c.name),
    [LIST_TITLES.antworten]: ANTWORTEN_COLUMNS.map((c) => c.name)
};

