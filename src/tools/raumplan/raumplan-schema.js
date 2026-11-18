/**
 * SharePoint-Liste: Raeume (Raumplan Stammdaten).
 */
import { toGraphColumnBody } from '../../shared/spo-graph-column.js';

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export { toGraphColumnBody };

export const LIST_TITLES = {
    raeume: 'Raeume'
};

export const RAUM_TYP_CHOICES = ['klasse', 'fachraum', 'buero', 'sporthalle', 'sonstiges'];

/** @type {GraphColumnDef[]} */
export const RAEUME_COLUMNS = [
    {
        name: 'RaumId',
        displayName: 'Raum-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Title',
        displayName: 'Titel',
        text: { allowMultipleLines: false, maxLength: 200 }
    },
    {
        name: 'Code',
        displayName: 'Code',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'Name',
        displayName: 'Name',
        text: { allowMultipleLines: false, maxLength: 160 }
    },
    {
        name: 'Gebaeude',
        displayName: 'Gebäude',
        text: { allowMultipleLines: false, maxLength: 80 }
    },
    {
        name: 'Stockwerk',
        displayName: 'Stockwerk',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'Kapazitaet',
        displayName: 'Kapazität',
        number: {}
    },
    {
        name: 'Typ',
        displayName: 'Typ',
        choice: { allowTextEntry: false, choices: [...RAUM_TYP_CHOICES] }
    },
    {
        name: 'Notiz',
        displayName: 'Notiz',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'SeedTag',
        displayName: 'Seed-Tag',
        text: { allowMultipleLines: false, maxLength: 40 }
    }
];

export const REQUIRED_COLUMNS = {
    [LIST_TITLES.raeume]: RAEUME_COLUMNS.map((c) => c.name)
};
