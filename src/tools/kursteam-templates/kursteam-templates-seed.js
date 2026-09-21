/**
 * Mitgelieferte MAM-Vorlagen (HAKB): Lehrplan-Modul = Halbjahr.
 * Modul 3 = Stufe 10 WS, 4 = 10 SS, 5 = 11 WS, 6 = 11 SS, 7 = 12 WS, 8 = 12 SS.
 */
import { normalizeTemplateList, hakLehrplanModulToMeta } from './kursteam-templates-logic.js';

function mam(modul, description, channels) {
    const meta = hakLehrplanModulToMeta(modul);
    const schulstufe = meta.schulstufe;
    const semester = meta.semester;
    return {
        id: 'tpl-mam-hakb-stufe-' + schulstufe + '-' + semester.toLowerCase(),
        name: 'HAKB MAM – Modul ' + modul + ' (Schulstufe ' + schulstufe + ' ' + semester + ')',
        schoolForm: 'HAKB',
        subjectCode: 'MAM',
        schulstufe,
        semester,
        description,
        channels
    };
}

/** @type {import('./kursteam-templates-logic.js').ChannelTemplate[]} */
const RAW = [
    mam(3, 'Modul 3 · Schulstufe 10 WS – Grundlagen bis quadratische Gleichungen', [
        '01 - Grundlagen der Mathematik',
        '02 - die 4 Grundrechnungsarten',
        '03 - Rechnen mit Prozenten',
        '04 - Maßeinheiten',
        '05 - Potenzen',
        '06 - Gleitkommadarstellung',
        '07 - Terme',
        '08 - Formeln',
        '09 - Lineare Gleichungen',
        '10 - Quadratische Gleichungen'
    ]),
    mam(4, 'Modul 4 · Schulstufe 10 SS – Funktionen, Matrizen, Trigonometrie', [
        '01-Funktionen - Grundlagen',
        '02-Lineare Funktionen',
        '03-Lineare Gleichungssysteme',
        '04-Anwendungsaufgaben linearer Gleichungen',
        '05-Quadratische Funktionen',
        '06-Quadratische Gleichungen',
        '07-Matrizen',
        '08-Trigonometrie'
    ]),
    mam(5, 'Modul 5 · Schulstufe 11 WS – Finanzmathematik und Wachstum', [
        '00-🗓️-Organisatorisches',
        '01-💵-Zins- und Zinseszinsrechnung',
        '02-💸-Rentenrechnung',
        '03-🪙-Tilgungspläne',
        '04-📈-Exponential- und Logarithmusfunktion',
        '05-↗️-lineare Wachstum- und Abnahmeprozesse',
        '06-🦠-exponentielle Wachstum- und Abnahmeprozesse',
        '07-⛔-beschränkte Wachstum- und Abnahmeprozesse',
        '08-🔀-Logistisches Wachstum'
    ]),
    mam(6, 'Modul 6 · Schulstufe 11 SS – Differenzial- und Integralrechnung', [
        '01-Grundlagen Differenzialrechnung',
        '02-Differenzialrechnung - Funktionsbetrachtungen',
        '03-Kosten- und Preistheorie',
        '04-Optimierungsprozesse',
        '05-Integralrechnung'
    ]),
    mam(7, 'Modul 7 · Schulstufe 12 WS – Investition, Statistik, Wahrscheinlichkeit', [
        '00-🗓️-Organisatorisches',
        '01-📈-Dynamische Investitionsrechnung',
        '02-📉-Kurs- und Rentabilitätsrechnung',
        '03-📊-Grundlagen der Statistik',
        '04-📊-Beschreibende Statistik',
        '05-🎲-Grundlagen Wahrscheinlichkeitsrechnung'
    ]),
    mam(8, 'Modul 8 · Schulstufe 12 SS – Verteilungen, sRDP, Wiederholung', [
        '01- 📊 Wahrscheinlichkeitsverteilungen',
        '02- 🎲 Binomialverteilung',
        '03- ＝ Normalverteilung',
        '04- 〰️ Approximation der BV durch NV',
        'sRDP - Vorbereitung',
        'WH-Wiederholung'
    ])
];

export function getSeedTemplates() {
    return normalizeTemplateList(RAW);
}
