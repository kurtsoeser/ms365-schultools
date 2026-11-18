/**
 * Vorlagen für Kollegiumsumfragen (einfach → komplex).
 */

import { FRAGE_TYP, ZIELGRUPPE, newId } from './kollegium-umfragen-model.js';

/**
 * @typedef {{
 *   id: string,
 *   titel: string,
 *   beschreibung: string,
 *   komplexitaet: 'einfach'|'komplex',
 *   zielgruppe: string,
 *   fragen: Array<object>
 * }} Vorlage
 */

/** @type {Vorlage[]} */
export const VORLAGEN = [
    {
        id: 'ja-nein-beschluss',
        titel: 'Ja/Nein-Beschluss',
        beschreibung: 'Schnelle Abstimmung im Kollegium zu einem konkreten Antrag.',
        komplexitaet: 'einfach',
        zielgruppe: ZIELGRUPPE.ALLE,
        fragen: [
            {
                id: 'q1',
                typ: FRAGE_TYP.SINGLE,
                text: 'Stimmen Sie dem Antrag zu?',
                pflicht: true,
                optionen: ['Ja', 'Nein', 'Enthaltung']
            },
            {
                id: 'q2',
                typ: FRAGE_TYP.TEXT,
                text: 'Kurzbegründung (optional)',
                pflicht: false,
                optionen: []
            }
        ]
    },
    {
        id: 'klassenwunsch',
        titel: 'Klassenwunsch nächstes Jahr',
        beschreibung: 'Lehrkräfte geben gewünschte Klassen bzw. Stufen an.',
        komplexitaet: 'einfach',
        zielgruppe: ZIELGRUPPE.ALLE,
        fragen: [
            {
                id: 'q1',
                typ: FRAGE_TYP.TEXT,
                text: 'Welche Klasse / Stufe wünschen Sie sich für das nächste Jahr?',
                pflicht: true,
                optionen: []
            },
            {
                id: 'q2',
                typ: FRAGE_TYP.SINGLE,
                text: 'Priorität',
                pflicht: true,
                optionen: ['1. Wunsch', '2. Wunsch', 'Flexibel']
            },
            {
                id: 'q3',
                typ: FRAGE_TYP.TEXT,
                text: 'Hinweise (Team, Fächer, besondere Gründe)',
                pflicht: false,
                optionen: []
            }
        ]
    },
    {
        id: 'schilf-themen',
        titel: 'SCHILF-Themen',
        beschreibung: 'Sammelt Wunschthemen für schulinterne Fortbildung.',
        komplexitaet: 'einfach',
        zielgruppe: ZIELGRUPPE.ALLE,
        fragen: [
            {
                id: 'q1',
                typ: FRAGE_TYP.MULTI,
                text: 'Welche Themen interessieren Sie?',
                pflicht: true,
                optionen: [
                    'Digitale Medien / KI',
                    'Klassenführung',
                    'Inklusion',
                    'Fachdidaktik',
                    'Gesundheit / Belastung',
                    'Sonstiges'
                ]
            },
            {
                id: 'q2',
                typ: FRAGE_TYP.TEXT,
                text: 'Eigenes Thema / Wunschformat',
                pflicht: false,
                optionen: []
            },
            {
                id: 'q3',
                typ: FRAGE_TYP.SKALA,
                text: 'Wie dringend ist Fortbildung für Sie aktuell?',
                pflicht: true,
                optionen: [],
                skalaMin: 1,
                skalaMax: 5,
                skalaLabelMin: 'gering',
                skalaLabelMax: 'sehr hoch'
            }
        ]
    },
    {
        id: 'stundenplan-rand',
        titel: 'Stundenplan-Randbedingungen',
        beschreibung: 'Rahmenbedingungen für die Stundenplanung (nicht anonym empfohlen).',
        komplexitaet: 'komplex',
        zielgruppe: ZIELGRUPPE.ALLE,
        fragen: [
            {
                id: 'q1',
                typ: FRAGE_TYP.MULTI,
                text: 'Bevorzugte freie Halbtage',
                pflicht: false,
                optionen: ['Mo vormittag', 'Mo nachmittag', 'Di', 'Mi', 'Do', 'Fr']
            },
            {
                id: 'q2',
                typ: FRAGE_TYP.SINGLE,
                text: 'Früher Unterricht (1./2. Stunde) möglich?',
                pflicht: true,
                optionen: ['Ja', 'Eingeschränkt', 'Nein']
            },
            {
                id: 'q3',
                typ: FRAGE_TYP.SINGLE,
                text: 'Später Unterricht (nach 15 Uhr) möglich?',
                pflicht: true,
                optionen: ['Ja', 'Eingeschränkt', 'Nein']
            },
            {
                id: 'q4',
                typ: FRAGE_TYP.TEXT,
                text: 'Weitere Randbedingungen',
                pflicht: false,
                optionen: []
            }
        ]
    },
    {
        id: 'kv-rueckmeldung',
        titel: 'KV-Kurzfeedback (nur Klassenvorstände)',
        beschreibung: 'Gezielte Abfrage an Klassenlehrer:innen – Zielgruppe Klassenvorstand.',
        komplexitaet: 'einfach',
        zielgruppe: ZIELGRUPPE.KV,
        fragen: [
            {
                id: 'q1',
                typ: FRAGE_TYP.TEXT,
                text: 'Für welche Klasse antworten Sie?',
                pflicht: true,
                optionen: []
            },
            {
                id: 'q2',
                typ: FRAGE_TYP.SKALA,
                text: 'Wie läuft die Klassengemeinschaft derzeit?',
                pflicht: true,
                optionen: [],
                skalaMin: 1,
                skalaMax: 5,
                skalaLabelMin: 'schwierig',
                skalaLabelMax: 'sehr gut'
            },
            {
                id: 'q3',
                typ: FRAGE_TYP.TEXT,
                text: 'Was braucht die Klasse / die KV-Rolle gerade?',
                pflicht: false,
                optionen: []
            }
        ]
    },
    {
        id: 'verfuegbarkeit',
        titel: 'Verfügbarkeit (Termin)',
        beschreibung: 'Einfache Termin-/Verfügbarkeitsabfrage.',
        komplexitaet: 'einfach',
        zielgruppe: ZIELGRUPPE.ALLE,
        fragen: [
            {
                id: 'q1',
                typ: FRAGE_TYP.MULTI,
                text: 'An welchen Terminen sind Sie verfügbar?',
                pflicht: true,
                optionen: ['Termin A', 'Termin B', 'Termin C', 'Keiner davon']
            },
            {
                id: 'q2',
                typ: FRAGE_TYP.TEXT,
                text: 'Alternative / Hinweis',
                pflicht: false,
                optionen: []
            }
        ]
    },
    {
        id: 'beleg',
        titel: 'Rückmeldung mit Beleg',
        beschreibung: 'Kurze Abfrage plus Datei-Upload (z. B. Bestätigung, Konzept, Foto).',
        komplexitaet: 'einfach',
        zielgruppe: ZIELGRUPPE.ALLE,
        anhangErlaubt: true,
        fragen: [
            {
                id: 'q1',
                typ: FRAGE_TYP.SINGLE,
                text: 'Status',
                pflicht: true,
                optionen: ['Erledigt', 'In Arbeit', 'Nicht möglich']
            },
            {
                id: 'q2',
                typ: FRAGE_TYP.DATEI,
                text: 'Bitte Beleg / Dokument hochladen',
                pflicht: true,
                optionen: []
            },
            {
                id: 'q3',
                typ: FRAGE_TYP.TEXT,
                text: 'Kurzkommentar',
                pflicht: false,
                optionen: []
            }
        ]
    }
];

/**
 * @param {string} id
 */
export function getVorlage(id) {
    return VORLAGEN.find((v) => v.id === String(id || '').trim()) || null;
}

/**
 * Erzeugt Umfrage-Draft aus Vorlage (frische Frage-IDs).
 * @param {string} vorlageId
 */
export function draftFromVorlage(vorlageId) {
    const v = getVorlage(vorlageId);
    if (!v) return null;
    return {
        id: newId('ku'),
        titel: v.titel,
        beschreibung: v.beschreibung,
        vorlageId: v.id,
        komplexitaet: v.komplexitaet,
        zielgruppe: v.zielgruppe,
        abteilung: '',
        klasseCodes: '',
        empfaengerMails: '',
        fragen: v.fragen.map((f) => ({
            ...f,
            id: newId('q')
        })),
        faelligAm: '',
        anonym: false,
        anhangErlaubt: !!v.anhangErlaubt,
        status: 'entwurf',
        quelle: 'vorlage'
    };
}
