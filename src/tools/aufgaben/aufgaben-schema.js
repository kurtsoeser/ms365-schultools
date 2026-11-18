/**
 * SharePoint-Liste: Aufgaben (interne Zuweisungen).
 */
import { toGraphColumnBody } from '../../shared/spo-graph-column.js';

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export { toGraphColumnBody };

export const LIST_TITLES = {
    aufgaben: 'Aufgaben'
};

export const STATUS_CHOICES = ['offen', 'in_arbeit', 'erledigt', 'abgelehnt'];
export const PRIORITAET_CHOICES = ['niedrig', 'mittel', 'hoch', 'dringend'];
export const ROLLE_CHOICES = ['direktion', 'it', 'kv', 'lehrer', 'verwaltung'];

/** @type {GraphColumnDef[]} */
export const AUFGABEN_COLUMNS = [
    {
        name: 'AufgabeId',
        displayName: 'Aufgabe-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Title',
        displayName: 'Titel',
        text: { allowMultipleLines: false, maxLength: 200 }
    },
    {
        name: 'Beschreibung',
        displayName: 'Beschreibung',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'ZuweisungName',
        displayName: 'Zuweisung Name',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'ZuweisungEmail',
        displayName: 'Zuweisung E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'Rolle',
        displayName: 'Rolle',
        choice: { allowTextEntry: true, choices: [...ROLLE_CHOICES] }
    },
    {
        name: 'Prioritaet',
        displayName: 'Priorität',
        choice: { allowTextEntry: false, choices: [...PRIORITAET_CHOICES] }
    },
    {
        name: 'Frist',
        displayName: 'Frist',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: [...STATUS_CHOICES] }
    },
    {
        name: 'ErstelltVon',
        displayName: 'Erstellt von',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'ErstelltVonEmail',
        displayName: 'Erstellt von E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'ErstelltAm',
        displayName: 'Erstellt am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'Kommentar',
        displayName: 'Kommentar',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'BearbeitetVonName',
        displayName: 'Bearbeitet von (Name)',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'BearbeitetVonEmail',
        displayName: 'Bearbeitet von (E-Mail)',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'BearbeitetAm',
        displayName: 'Bearbeitet am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'SeedTag',
        displayName: 'Seed-Tag',
        text: { allowMultipleLines: false, maxLength: 60 }
    }
];

export const REQUIRED_COLUMNS = {
    [LIST_TITLES.aufgaben]: AUFGABEN_COLUMNS.map((c) => c.name)
};
