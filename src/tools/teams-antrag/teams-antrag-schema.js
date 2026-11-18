/**
 * SharePoint-Liste: Team-Anträge.
 */
import { toGraphColumnBody } from '../../shared/spo-graph-column.js';

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export { toGraphColumnBody };

export const LIST_TITLES = {
    antraege: 'TeamAntraege'
};

export const STATUS_CHOICES = [
    'entwurf',
    'beantragt',
    'in_pruefung',
    'genehmigt',
    'angelegt',
    'abgelehnt',
    'zurueckgestellt',
    'anlage_fehlgeschlagen'
];

export const ART_CHOICES = ['projekt', 'ag', 'fachgruppe', 'klasse', 'diplom_vwa', 'sonstiges'];

/** @type {GraphColumnDef[]} */
export const ANTRAEGE_COLUMNS = [
    {
        name: 'AntragId',
        displayName: 'Antrag-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Zweck',
        displayName: 'Zweck / Begründung',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'Art',
        displayName: 'Art',
        choice: { allowTextEntry: false, choices: [...ART_CHOICES] }
    },
    {
        name: 'Namensvorschlag',
        displayName: 'Namensvorschlag',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'BesitzerMails',
        displayName: 'Besitzer-Mails',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'MitgliederMails',
        displayName: 'Mitglieder-Mails',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'Schuljahr',
        displayName: 'Schuljahr',
        text: { allowMultipleLines: false, maxLength: 20 }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: [...STATUS_CHOICES] }
    },
    {
        name: 'AntragstellerEmail',
        displayName: 'Antragsteller-E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'AntragstellerName',
        displayName: 'Antragsteller',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'FinalDisplayName',
        displayName: 'Finaler Anzeigename',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'FinalMailNickname',
        displayName: 'Finaler Alias',
        text: { allowMultipleLines: false, maxLength: 64 }
    },
    {
        name: 'TeamId',
        displayName: 'Team-ID',
        text: { allowMultipleLines: false, maxLength: 80 }
    },
    {
        name: 'TeamWebUrl',
        displayName: 'Team-URL',
        text: { allowMultipleLines: false, maxLength: 500 }
    },
    {
        name: 'Ablehnungsgrund',
        displayName: 'Ablehnungsgrund',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'NotizIt',
        displayName: 'IT-Notiz',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'Quelle',
        displayName: 'Quelle',
        choice: { allowTextEntry: false, choices: ['ui', 'demo', 'import'] }
    }
];
