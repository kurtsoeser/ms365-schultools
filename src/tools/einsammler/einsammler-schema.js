/**
 * SharePoint-Listen-Schema: Einsammler (Eltern & Ablauf).
 * Kopf + Antworten – System of Record auf der Intranet-Site.
 */
import { toGraphColumnBody } from '../../shared/spo-graph-column.js';

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export { toGraphColumnBody };

export const LIST_TITLES = {
    koepfe: 'Einsammler',
    antworten: 'EinsammlerAntworten'
};

export const EINSAMMLER_STATUS_CHOICES = ['entwurf', 'aktiv', 'geschlossen', 'archiviert'];

export const EINSAMMLER_MODUS_CHOICES = ['einwilligung', 'lesebestaetigung', 'abgabe'];

export const EINSAMMLER_TYP_CHOICES = [
    'exkursion',
    'foto',
    'datenschutz',
    'wandertag',
    'elternbrief',
    'sonstiges'
];

export const ANTWORT_STATUS_CHOICES = [
    'offen',
    'zugestimmt',
    'abgelehnt',
    'gelesen',
    'abgegeben',
    'abgelaufen'
];

/** @type {GraphColumnDef[]} */
export const KOEPFE_COLUMNS = [
    {
        name: 'EinsammlerId',
        displayName: 'Einsammler-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Modus',
        displayName: 'Modus',
        choice: { allowTextEntry: false, choices: [...EINSAMMLER_MODUS_CHOICES] }
    },
    {
        name: 'Typ',
        displayName: 'Typ',
        choice: { allowTextEntry: false, choices: [...EINSAMMLER_TYP_CHOICES] }
    },
    {
        name: 'TextInhalt',
        displayName: 'Text / Inhalt',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'KlasseCodes',
        displayName: 'Klassen-Codes',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'FaelligAm',
        displayName: 'Fällig am',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: [...EINSAMMLER_STATUS_CHOICES] }
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
        name: 'UploadErlaubt',
        displayName: 'Upload erlaubt',
        boolean: {}
    },
    {
        name: 'UploadPflicht',
        displayName: 'Upload Pflicht',
        boolean: {}
    },
    {
        name: 'Quelle',
        displayName: 'Quelle',
        choice: { allowTextEntry: false, choices: ['ui', 'public', 'import'] }
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
        name: 'EinsammlerId',
        displayName: 'Einsammler-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'SchuelerName',
        displayName: 'Schüler:in',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'KlasseCode',
        displayName: 'Klasse-Code',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'ElternEmail',
        displayName: 'Eltern-E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: [...ANTWORT_STATUS_CHOICES] }
    },
    {
        name: 'AntwortAm',
        displayName: 'Antwort am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'Kommentar',
        displayName: 'Kommentar',
        text: { allowMultipleLines: true, textType: 'plain' }
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
        displayName: 'Anhang-URL (SharePoint)',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'Quelle',
        displayName: 'Quelle',
        choice: { allowTextEntry: false, choices: ['ui', 'public', 'import'] }
    }
];

export const REQUIRED_COLUMNS = {
    [LIST_TITLES.koepfe]: KOEPFE_COLUMNS.map((c) => c.name),
    [LIST_TITLES.antworten]: ANTWORTEN_COLUMNS.map((c) => c.name)
};

