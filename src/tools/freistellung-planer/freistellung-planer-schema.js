/**
 * SharePoint-Listen-Schema: Freistellungen (kompatibel mit freistellung-setup + PA-Flow).
 */

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export const LIST_TITLE_DEFAULT = 'Freistellungen';

/** Status-Werte wie in freistellung-setup / Power-Automate-Flow */
export const STATUS_CHOICES = ['Ausstehend', 'Genehmigt', 'Abgelehnt'];

export const KATEGORIE_CHOICES = [
    'Ärztlicher Termin',
    'Familiäre Angelegenheit',
    'Bewerbung / Schnuppertag',
    'Sonstiges'
];

/** Ab diesem inklusiven Tagesumfang gilt der Antrag als mehrtägig (KV + Direktion). */
export const MULTI_DAY_THRESHOLD = 2;

/** @type {GraphColumnDef[]} */
export const FREISTELLUNG_COLUMNS = [
    {
        name: 'Beginn',
        displayName: 'Beginn',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Ende',
        displayName: 'Ende',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: {
            allowTextEntry: false,
            choices: [...STATUS_CHOICES]
        }
    },
    {
        name: 'Klasse',
        displayName: 'Klasse',
        choice: {
            allowTextEntry: true,
            choices: ['1AHW', '2AHW', '3AHW', '4AHW', '5AHW']
        }
    },
    {
        name: 'Klassenvorstand',
        displayName: 'Klassenvorstand',
        personOrGroup: {
            allowMultipleSelection: false,
            chooseFromType: 'peopleOnly'
        }
    },
    {
        name: 'Kategorie',
        displayName: 'Kategorie',
        choice: {
            allowTextEntry: true,
            choices: [...KATEGORIE_CHOICES]
        }
    },
    {
        name: 'Beschreibung',
        displayName: 'Beschreibung',
        text: { allowMultipleLines: true, maxLength: 8000 }
    },
    {
        name: 'Bemerkungen',
        displayName: 'Bemerkungen',
        text: { allowMultipleLines: true, maxLength: 8000 }
    }
];

export const REQUIRED_COLUMN_NAMES = FREISTELLUNG_COLUMNS.map((c) => c.name);

/**
 * @param {GraphColumnDef} def
 */
export function toGraphColumnBody(def) {
    const body = {
        name: def.name,
        displayName: def.displayName || def.name
    };
    if (def.text) body.text = def.text;
    if (def.number) body.number = def.number;
    if (def.boolean) body.boolean = def.boolean;
    if (def.dateTime) body.dateTime = def.dateTime;
    if (def.choice) body.choice = def.choice;
    if (def.personOrGroup) body.personOrGroup = def.personOrGroup;
    return body;
}
