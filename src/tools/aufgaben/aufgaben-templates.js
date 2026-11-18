/**
 * Vorlagen für Aufgaben (P3-MVP).
 */

import { PRIORITAET, ROLLE } from './aufgaben-model.js';

/**
 * @typedef {{
 *   id: string,
 *   titel: string,
 *   beschreibung: string,
 *   rolle?: string,
 *   prioritaet?: string,
 *   hinweis?: string
 * }} AufgabeVorlage
 */

/** @type {AufgabeVorlage[]} */
export const AUFGABEN_VORLAGEN = [
    {
        id: 'unterschrift',
        titel: 'Bitte Unterschrift',
        beschreibung: 'Formular / Antrag zur Unterschrift vorlegen und zurückgeben.',
        rolle: ROLLE.DIREKTION,
        prioritaet: PRIORITAET.MITTEL,
        hinweis: 'Oft Direktion oder Verwaltung'
    },
    {
        id: 'unterlagen',
        titel: 'Unterlagen nachreichen',
        beschreibung: 'Fehlende Unterlagen bis zur Frist an das Sekretariat senden.',
        rolle: ROLLE.LEHRER,
        prioritaet: PRIORITAET.HOCH,
        hinweis: 'Frist setzen'
    },
    {
        id: 'raum-pruefen',
        titel: 'Raum prüfen',
        beschreibung: 'Raum nach Veranstaltung auf Vollständigkeit und Ordnung prüfen.',
        rolle: ROLLE.KV,
        prioritaet: PRIORITAET.MITTEL,
        hinweis: 'Raum/Ort in der Beschreibung nennen'
    },
    {
        id: 'elternkontakt',
        titel: 'Elternkontakt',
        beschreibung: 'Erziehungsberechtigte wegen Fehlstunden / Anliegen kontaktieren und Rückmeldung notieren.',
        rolle: ROLLE.KV,
        prioritaet: PRIORITAET.DRINGEND,
        hinweis: 'Schüler/Klasse im Text nennen'
    }
];

/**
 * @param {string} id
 * @returns {AufgabeVorlage|null}
 */
export function getAufgabeVorlage(id) {
    const key = String(id || '').trim();
    return AUFGABEN_VORLAGEN.find((v) => v.id === key) || null;
}

/**
 * Draft-Felder für das Anlege-Formular (ohne Zuweisung/ID).
 * @param {string} id
 */
export function draftFromAufgabeVorlage(id) {
    const v = getAufgabeVorlage(id);
    if (!v) return null;
    return {
        title: v.titel,
        beschreibung: v.beschreibung,
        rolle: v.rolle || '',
        prioritaet: v.prioritaet || PRIORITAET.MITTEL,
        vorlageId: v.id
    };
}
