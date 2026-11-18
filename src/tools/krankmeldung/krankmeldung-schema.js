/**
 * SharePoint-Liste: Krankmeldungen (Eltern & Ablauf).
 */
import { toGraphColumnBody } from '../../shared/spo-graph-column.js';

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export { toGraphColumnBody };

export const LIST_TITLES = {
    meldungen: 'Krankmeldungen'
};

export const STATUS_CHOICES = ['neu', 'zur_kenntnis', 'archiviert'];

export const KATEGORIE_CHOICES = ['krankheit', 'arzt', 'sonstiges'];

/** @type {GraphColumnDef[]} */
export const MELDUNGEN_COLUMNS = [
    {
        name: 'MeldungId',
        displayName: 'Meldungs-ID',
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
        name: 'Von',
        displayName: 'Von',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Bis',
        displayName: 'Bis',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Kategorie',
        displayName: 'Kategorie',
        choice: { allowTextEntry: false, choices: [...KATEGORIE_CHOICES] }
    },
    {
        name: 'GemeldetAm',
        displayName: 'Gemeldet am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'GemeldetVonEmail',
        displayName: 'Gemeldet von (E-Mail)',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'GemeldetVonName',
        displayName: 'Gemeldet von (Name)',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: [...STATUS_CHOICES] }
    },
    {
        name: 'KenntnisLehrerEmail',
        displayName: 'Kenntnis Lehrkraft',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'Notiz',
        displayName: 'Notiz',
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
        displayName: 'Anhang-URL',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'Quelle',
        displayName: 'Quelle',
        choice: { allowTextEntry: false, choices: ['public', 'ui', 'import'] }
    }
];

export const REQUIRED_COLUMNS = {
    [LIST_TITLES.meldungen]: MELDUNGEN_COLUMNS.map((c) => c.name)
};

