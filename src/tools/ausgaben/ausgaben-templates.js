/**
 * Vorlagen für Ausgaben-Typen (Schulen können erweitern / umbenennen).
 */

import { newId } from './ausgaben-model.js';

/**
 * @typedef {{ id: string, label: string, icon: string, hinweis: string, beispiele?: Array<{kennung:string,bezeichnung:string,ortZone:string}> }} TypVorlage
 */

/** @type {TypVorlage[]} */
export const TYP_VORLAGEN = [
    {
        id: 'schluessel',
        label: 'Schlüssel',
        icon: 'bi-key',
        hinweis: 'Klassenzimmer, Generalschlüssel, Serverraum, Beamer-Koffer …',
        beispiele: [
            { kennung: 'S-12b', bezeichnung: 'Klassenzimmer 12b', ortZone: 'Trakt A' },
            { kennung: 'S-GEN', bezeichnung: 'Generalschlüssel Gebäude A', ortZone: 'Pforte' }
        ]
    },
    {
        id: 'raumkarte',
        label: 'Raumkarte',
        icon: 'bi-credit-card',
        hinweis: 'Elektronische Zutrittskarten für Räume / Labore',
        beispiele: [
            { kennung: 'RK-IT-01', bezeichnung: 'IT-Raum Karte 01', ortZone: 'IT-Trakt' }
        ]
    },
    {
        id: 'schrankenkarte',
        label: 'Schrankenkarte',
        icon: 'bi-door-closed',
        hinweis: 'Zufahrt Schranke / Parkplatzstor',
        beispiele: [
            { kennung: 'SK-07', bezeichnung: 'Schrankenkarte 07', ortZone: 'Tor Nord' }
        ]
    },
    {
        id: 'parkplatz',
        label: 'Parkplatz / Zufahrt',
        icon: 'bi-p-circle',
        hinweis: 'Parkberechtigung, Kennzeichen-Hinweis (ohne Kennzeichenerkennung)',
        beispiele: [
            { kennung: 'P-A12', bezeichnung: 'Parkplatz A12', ortZone: 'Hof West' }
        ]
    },
    {
        id: 'laptop',
        label: 'Laptop / Tablet',
        icon: 'bi-laptop',
        hinweis: 'Leihgeräte – gleiche Logik, anderer Typ',
        beispiele: [
            { kennung: 'NB-104', bezeichnung: 'Notebook 104', ortZone: 'Medienwagen' }
        ]
    },
    {
        id: 'sonstiges',
        label: 'Sonstiges',
        icon: 'bi-box',
        hinweis: 'Freier Typ – Schule benennt selbst',
        beispiele: []
    }
];

/**
 * @param {string} id
 */
export function getTypVorlage(id) {
    return TYP_VORLAGEN.find((t) => t.id === String(id || '').trim()) || null;
}

/**
 * Typ + optionale Beispiel-Gegenstände aus Vorlage.
 * @param {string} vorlageId
 * @param {{ withExamples?: boolean }} [opts]
 */
export function materializeTypVorlage(vorlageId, opts = {}) {
    const v = getTypVorlage(vorlageId);
    if (!v) return null;
    const typ = {
        id: v.id,
        label: v.label,
        icon: v.icon,
        hinweis: v.hinweis
    };
    const gegenstaende = opts.withExamples
        ? (v.beispiele || []).map((b) => ({
              id: newId('gg'),
              typ: v.id,
              typLabel: v.label,
              kennung: b.kennung,
              bezeichnung: b.bezeichnung,
              ortZone: b.ortZone,
              aktiv: true,
              notiz: '',
              quelle: 'vorlage'
          }))
        : [];
    return { typ, gegenstaende };
}
