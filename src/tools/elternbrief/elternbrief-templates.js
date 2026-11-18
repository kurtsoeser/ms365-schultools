/**
 * Vorlagen für Elternbriefe (Schule kann Texte anpassen).
 */

import { newId } from './elternbrief-model.js';

/** @type {Array<{ id: string, titel: string, betreff: string, inhalt: string, hinweis: string }>} */
export const VORLAGEN = [
    {
        id: 'schulbeginn',
        titel: 'Schulbeginn / Willkommen',
        betreff: 'Willkommen zum neuen Schuljahr',
        hinweis: 'Startinfo an alle Klassen',
        inhalt:
            'Liebe Erziehungsberechtigte,\n\nwir begrüßen Sie und Ihr Kind herzlich zum neuen Schuljahr.\n\nWichtige Termine und Hinweise folgen in Kürze. Bei Fragen wenden Sie sich bitte an die Klassenvorständin / den Klassenvorstand.\n\nMit freundlichen Grüßen\nDie Schulleitung'
    },
    {
        id: 'schularbeitenplan',
        titel: 'Schularbeitenplan',
        betreff: 'Schularbeitenplan – bitte zur Kenntnis',
        hinweis: 'Mit Lesebestätigung',
        inhalt:
            'Liebe Erziehungsberechtigte,\n\nim Anhang bzw. unten finden Sie den Schularbeitenplan für das aktuelle Quartal.\n\nBitte nehmen Sie die Termine zur Kenntnis und bestätigen Sie die Lesebestätigung über den persönlichen Link.\n\nMit freundlichen Grüßen'
    },
    {
        id: 'ausflug',
        titel: 'Ausflug / Exkursion (Info)',
        betreff: 'Information zum Klassenausflug',
        hinweis: 'Oft kombiniert mit Einsammler-Einwilligung',
        inhalt:
            'Liebe Erziehungsberechtigte,\n\nwir planen einen Klassenausflug. Details (Datum, Treffpunkt, Kosten, Ausrüstung) siehe unten bzw. Anhang.\n\nEine allfällige Einwilligung erhalten Sie ggf. zusätzlich über den Einsammler.\n\nMit freundlichen Grüßen'
    },
    {
        id: 'elternabend',
        titel: 'Elternabend / Termin',
        betreff: 'Einladung zum Elternabend',
        hinweis: 'Terminankündigung',
        inhalt:
            'Liebe Erziehungsberechtigte,\n\nwir laden Sie herzlich zum Elternabend ein.\n\nDatum / Uhrzeit / Ort: bitte ergänzen\n\nThemen: …\n\nMit freundlichen Grüßen'
    },
    {
        id: 'gesundheit',
        titel: 'Gesundheit / Absenzen-Hinweis',
        betreff: 'Hinweis zu Krankenmeldungen',
        hinweis: 'Organisation Sekretariat',
        inhalt:
            'Liebe Erziehungsberechtigte,\n\nbitte melden Sie Erkrankungen Ihres Kindes rechtzeitig der Schule (Sekretariat / Krankmelde-Formular).\n\nVielen Dank für Ihre Unterstützung.\n\nMit freundlichen Grüßen'
    },
    {
        id: 'frei',
        titel: 'Freier Elternbrief',
        betreff: '',
        hinweis: 'Leere Vorlage',
        inhalt: 'Liebe Erziehungsberechtigte,\n\n\n\nMit freundlichen Grüßen'
    }
];

/**
 * @param {string} id
 */
export function getVorlage(id) {
    return VORLAGEN.find((v) => v.id === String(id || '').trim()) || null;
}

/**
 * @param {string} vorlageId
 */
export function draftFromVorlage(vorlageId) {
    const v = getVorlage(vorlageId) || getVorlage('frei');
    if (!v) return null;
    return {
        id: newId('eb'),
        betreff: v.betreff,
        inhalt: v.inhalt,
        vorlageId: v.id,
        klasseCodes: '',
        lesebestaetigung: true,
        status: 'entwurf'
    };
}
