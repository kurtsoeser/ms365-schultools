/**
 * SharePoint-Liste: TechnikMeldungen (IT-Inbox).
 */
import { toGraphColumnBody } from '../../shared/spo-graph-column.js';

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export { toGraphColumnBody };

export const LIST_TITLES = {
    meldungen: 'TechnikMeldungen'
};

export const STATUS_CHOICES = ['offen', 'in_arbeit', 'erledigt', 'abgelehnt'];
export const PRIORITAET_CHOICES = ['niedrig', 'mittel', 'hoch', 'dringend'];
export const KATEGORIE_CHOICES = ['pc', 'beamer', 'audio', 'netzwerk', 'sonstiges'];

/** @type {GraphColumnDef[]} */
export const MELDUNGEN_COLUMNS = [
    {
        name: 'MeldungId',
        displayName: 'Meldung-ID',
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
        name: 'RaumCode',
        displayName: 'Raum-Code',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'RaumName',
        displayName: 'Raum-Name',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'Kategorie',
        displayName: 'Kategorie',
        choice: { allowTextEntry: false, choices: [...KATEGORIE_CHOICES] }
    },
    {
        name: 'Prioritaet',
        displayName: 'Priorität',
        choice: { allowTextEntry: false, choices: [...PRIORITAET_CHOICES] }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: [...STATUS_CHOICES] }
    },
    {
        name: 'MelderName',
        displayName: 'Melder-Name',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'MelderEmail',
        displayName: 'Melder-E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'GemeldetAm',
        displayName: 'Gemeldet am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'ErledigtAm',
        displayName: 'Erledigt am',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
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
        text: { allowMultipleLines: false, maxLength: 40 }
    }
];

export const REQUIRED_COLUMNS = {
    [LIST_TITLES.meldungen]: MELDUNGEN_COLUMNS.map((c) => c.name)
};
