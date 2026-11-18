/**
 * SharePoint-Liste SchulTermineHub für Kalender und Termine (oder Erweiterung von Schultermine).
 * Zusatzfelder: Sichtbarkeit, TerminTyp, Ort, Zielgruppe, GraphEventId.
 */
import { toGraphColumnBody } from '../../shared/spo-graph-column.js';

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export { toGraphColumnBody };

export const LIST_TITLES = {
    termine: 'SchulTermineHub'
};

/** Bestehende Schultermine-Liste – Felder zum Nachziehen. */
export const LEGACY_LIST_TITLE = 'Schultermine';

export const SICHTBARKEIT_CHOICES = ['oeffentlich', 'intern'];
export const TERMIN_TYP_CHOICES = [
    'konferenz',
    'fortbildung',
    'veranstaltung',
    'pruefung',
    'organisation',
    'arge',
    'beratung',
    'sonstiges'
];
export const SYNC_STATUS_CHOICES = ['lokal', 'pending', 'synced'];

/** @type {GraphColumnDef[]} */
export const TERMINE_COLUMNS = [
    {
        name: 'TerminId',
        displayName: 'Termin-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Title',
        displayName: 'Titel',
        text: { allowMultipleLines: false, maxLength: 200 }
    },
    {
        name: 'Beginn',
        displayName: 'Beginn',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'Ende',
        displayName: 'Ende',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'AllDay',
        displayName: 'Ganztägig',
        boolean: {}
    },
    {
        name: 'Ort',
        displayName: 'Ort',
        text: { allowMultipleLines: false, maxLength: 160 }
    },
    {
        name: 'Beschreibung',
        displayName: 'Beschreibung',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'Sichtbarkeit',
        displayName: 'Sichtbarkeit',
        choice: { allowTextEntry: false, choices: [...SICHTBARKEIT_CHOICES] }
    },
    {
        name: 'TerminTyp',
        displayName: 'Termin-Typ',
        choice: { allowTextEntry: false, choices: [...TERMIN_TYP_CHOICES] }
    },
    {
        name: 'Zielgruppe',
        displayName: 'Zielgruppe',
        text: { allowMultipleLines: false, maxLength: 160 }
    },
    {
        name: 'GraphEventId',
        displayName: 'Graph Event-ID',
        text: { allowMultipleLines: false, maxLength: 200 }
    },
    {
        name: 'SyncStatus',
        displayName: 'Sync-Status',
        choice: { allowTextEntry: false, choices: [...SYNC_STATUS_CHOICES] }
    }
];

/** Spalten, die an bestehende Liste „Schultermine“ angehängt werden können. */
export const SCHULTERMINE_EXTENSION_COLUMNS = TERMINE_COLUMNS.filter((c) =>
    ['Ort', 'Sichtbarkeit', 'TerminTyp', 'Zielgruppe', 'GraphEventId'].includes(c.name)
);

export const REQUIRED_COLUMNS = {
    [LIST_TITLES.termine]: TERMINE_COLUMNS.map((c) => c.name),
    [LEGACY_LIST_TITLE]: [
        'Title',
        'Beginn',
        'Ende',
        'AllDay',
        'Ort',
        'Sichtbarkeit',
        'TerminTyp',
        'Zielgruppe',
        'GraphEventId',
        'SyncStatus'
    ]
};
