/**
 * SharePoint-Listen-Schema: Freistellungen Lehrkräfte (eine Genehmigung: Direktion).
 * Power Automate: Trigger auf neue Liste, Approvals nur Direktion, Status zurück in die Liste.
 */

export const LIST_TITLE_DEFAULT = 'Lehrer-Freistellungen';

export const STATUS_CHOICES = ['Ausstehend', 'Genehmigt', 'Abgelehnt'];

export const KATEGORIE_CHOICES = [
    'Fortbildung',
    'Arzttermin',
    'Persönliche Angelegenheit',
    'Dienstfreistellung',
    'Sonstiges'
];

/** @returns {string} */
export function newAntragId() {
    return (
        'lfr-' +
        Date.now().toString(36) +
        '-' +
        Math.random().toString(36).slice(2, 10)
    );
}

/** @type {import('../freistellung-planer/freistellung-planer-schema.js').GraphColumnDef[]} */
export const LFR_COLUMNS = [
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
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: [...STATUS_CHOICES] }
    },
    {
        name: 'Kategorie',
        displayName: 'Kategorie',
        choice: { allowTextEntry: true, choices: [...KATEGORIE_CHOICES] }
    },
    {
        name: 'Beschreibung',
        displayName: 'Beschreibung',
        text: { allowMultipleLines: true, maxLength: 8000 }
    },
    {
        name: 'LehrerName',
        displayName: 'Lehrkraft',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'LehrerEmail',
        displayName: 'E-Mail Lehrkraft',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'AntragId',
        displayName: 'Antrag-ID',
        text: { allowMultipleLines: false, maxLength: 64 }
    },
    {
        name: 'GenehmigtVonDirektion',
        displayName: 'Genehmigt von (Direktion)',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'GenehmigtAmDirektion',
        displayName: 'Genehmigt am (Direktion)',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'AbgelehntVon',
        displayName: 'Abgelehnt von',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'AbgelehntAm',
        displayName: 'Abgelehnt am',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'BemerkungDirektion',
        displayName: 'Bemerkung Direktion',
        text: { allowMultipleLines: true, maxLength: 2000 }
    }
];
