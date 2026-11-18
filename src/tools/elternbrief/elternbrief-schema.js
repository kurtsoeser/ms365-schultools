/**
 * SharePoint-Listen: Elternbriefe (Köpfe + Lesebestätigungen).
 */
import { toGraphColumnBody } from '../../shared/spo-graph-column.js';

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export { toGraphColumnBody };

export const LIST_TITLES = {
    briefe: 'Elternbriefe',
    lesebestaetigungen: 'ElternbriefLesebestaetigungen'
};

export const STATUS_CHOICES = ['entwurf', 'bereit', 'gesendet', 'geschlossen', 'archiviert'];

export const ABSENDER_MODUS_CHOICES = ['eigenes_konto', 'zentrales_postfach'];

/** @type {GraphColumnDef[]} */
export const BRIEFE_COLUMNS = [
    {
        name: 'BriefId',
        displayName: 'Brief-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Betreff',
        displayName: 'Betreff',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'InhaltHtml',
        displayName: 'Inhalt (HTML/Text)',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'VorlageId',
        displayName: 'Vorlage-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'KlasseCodes',
        displayName: 'Klassen-Codes',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'AbsenderModus',
        displayName: 'Absender-Modus',
        choice: { allowTextEntry: false, choices: [...ABSENDER_MODUS_CHOICES] }
    },
    {
        name: 'AbsenderEmail',
        displayName: 'Absender-E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'AbsenderName',
        displayName: 'Absender-Name',
        text: { allowMultipleLines: false, maxLength: 120 }
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
        name: 'FaelligAm',
        displayName: 'Lesebestätigung bis',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Lesebestaetigung',
        displayName: 'Lesebestätigung aktiv',
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
        name: 'GesendetAm',
        displayName: 'Gesendet am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'SeedTag',
        displayName: 'Seed-Tag',
        text: { allowMultipleLines: false, maxLength: 80 }
    }
];

/** @type {GraphColumnDef[]} */
export const LESE_COLUMNS = [
    {
        name: 'LeseId',
        displayName: 'Lese-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'BriefId',
        displayName: 'Brief-ID',
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
        choice: { allowTextEntry: false, choices: ['offen', 'gelesen', 'abgelaufen'] }
    },
    {
        name: 'GelesenAm',
        displayName: 'Gelesen am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'SeedTag',
        displayName: 'Seed-Tag',
        text: { allowMultipleLines: false, maxLength: 80 }
    }
];

export const REQUIRED_COLUMNS = {
    [LIST_TITLES.briefe]: BRIEFE_COLUMNS.map((c) => c.name),
    [LIST_TITLES.lesebestaetigungen]: LESE_COLUMNS.map((c) => c.name)
};

