/**
 * Schulform-Profile für sRDP/sRP-Anmeldung (HAK, HTL, HLW, BAFEB, AHS).
 * Quellen: BMB Teilprüfungen-Übersicht, hak.cc, HTL-Varianten, HUM/HLW, BAfEP, AHS-SRDP.
 */

import { VARIANT_LABELS } from './srdp-anmeldung-schema-constants.js';

export { VARIANT_LABELS };

/** @typedef {{ key: string, label: string, schriftlich: string[], muendlich: string[] }} SrdpVariant */
/** @typedef {{ displayname: string, fields: string[] }} FormSection */
/**
 * @typedef {{
 *   id: string,
 *   label: string,
 *   examShort: string,
 *   listTitleCode: string,
 *   arbeitLabel: string,
 *   variants: SrdpVariant[],
 *   defaultWahlfaecher: string[],
 *   defaultSeminare: string[],
 *   lfsHints: string[],
 *   lfsFieldLabel: string,
 *   formSections: FormSection[],
 *   extraColumnNames: string[],
 *   notes?: string
 * }} SrdpProfile
 */

const WAHL_GENERIC = [
    'Religion / Ethik',
    'Geschichte und Sozialkunde / Politische Bildung',
    'Geografie und Wirtschaftskunde',
    'Biologie und Umweltkunde',
    'Chemie',
    'Physik',
    'Musik',
    'Bildnerische Erziehung',
    'Bewegung und Sport',
    'Informatik',
    'Psychologie und Philosophie',
    'Seminar',
    'Freigegenstand'
];

const WAHL_HAK = [
    'Religion / Ethik',
    'Kultur',
    'Geschichte und Internationale Wirtschafts- und Kulturräume',
    'Geografie und Internationale Wirtschafts- und Kulturräume',
    'Naturwissenschaften',
    'Recht',
    'Volkswirtschaft',
    'Berufsbezogene Kommunikation in der lebenden Fremdsprache',
    'Mehrsprachigkeit',
    'Wirtschaftsinformatik',
    'Seminar',
    'Freigegenstand'
];

const WAHL_HLW = [
    'Fachkolloquium-Schwerpunkt (schulautonom)',
    'Berufsbezogene Kommunikation in der Fremdsprache',
    'Mehrsprachigkeit',
    'Kultur und gesellschaftliche Reflexion',
    'Globalwirtschaft, Wirtschaftsgeografie und Volkswirtschaft',
    'Naturwissenschaften',
    'Ernährung und Lebensmitteltechnologie',
    'Ökomanagement',
    'Seminar',
    'Freigegenstand'
];

const WAHL_BAFEB = [
    'Biologie und Ökologie',
    'Angewandte Naturwissenschaften (Physik)',
    'Angewandte Naturwissenschaften (Chemie)',
    'Geografie und Wirtschaftskunde',
    'Geschichte und Politische Bildung',
    'Musik',
    'Rhythmik-Musik-Bewegung',
    'Bildnerische Erziehung',
    'Bewegung und Sport',
    'Seminar',
    'Freigegenstand / berufsspezifisches Prüfungsgebiet'
];

const LFS_HINTS = ['EN', 'ENWS', 'FR', 'FRWS', 'IT', 'ITWS', 'SP', 'SPWS'];

/** @type {SrdpProfile} */
export const PROFILE_HAK = {
    id: 'hak',
    label: 'HAK',
    examShort: 'sRDP',
    listTitleCode: 'HAK',
    arbeitLabel: 'Diplomarbeit',
    variants: [
        {
            key: '1',
            label: VARIANT_LABELS[1],
            schriftlich: ['D (5h)', 'BFK (6h)', 'LFS (5h)'],
            muendlich: ['BKO', 'AM', 'Wahlfach']
        },
        {
            key: '2',
            label: VARIANT_LABELS[2],
            schriftlich: ['D (5h)', 'BFK (6h)', 'AM (4,5h)'],
            muendlich: ['BKO', 'LFS', 'Wahlfach']
        },
        {
            key: '3',
            label: VARIANT_LABELS[3],
            schriftlich: ['D (5h)', 'BFK (6h)', 'LFS (5h)', 'AM (4,5h)'],
            muendlich: ['BKO', 'Wahlfach']
        }
    ],
    defaultWahlfaecher: WAHL_HAK,
    defaultSeminare: [],
    lfsHints: LFS_HINTS,
    lfsFieldLabel: 'LFS',
    extraColumnNames: ['LFS', 'LehrerLFS', 'LehrerBKO'],
    formSections: [
        { displayname: 'Person & Klasse', fields: ['Klasse', 'Nachname', 'Vorname'] },
        {
            displayname: 'Diplomarbeit',
            fields: ['Titel Diplomarbeit', 'BetreuungslehrerIn Diplomarbeit']
        },
        {
            displayname: 'Prüfungsvariante',
            fields: ['Variante', 'LFS', 'LehrerIn LFS', 'LehrerIn BKO mündlich']
        },
        {
            displayname: 'Wahlfach',
            fields: ['Wahlfach mündlich', 'Seminar', 'LehrerIn Wahlfach']
        },
        { displayname: 'Bestätigung', fields: ['Bestätigung Anmeldung'] }
    ],
    notes: 'Quelle: hak.cc / BMB BHS – Variante 1–3.'
};

/** @type {SrdpProfile} */
export const PROFILE_HTL = {
    id: 'htl',
    label: 'HTL',
    examShort: 'sRDP',
    listTitleCode: 'HTL',
    arbeitLabel: 'Diplomarbeit',
    variants: [
        {
            key: '1',
            label: VARIANT_LABELS[1],
            schriftlich: ['D', 'AM', 'Fachtheorie (SWP)'],
            muendlich: ['Englisch', 'Schwerpunkt', 'Wahlfach']
        },
        {
            key: '2',
            label: VARIANT_LABELS[2],
            schriftlich: ['Englisch', 'AM', 'Fachtheorie (SWP)'],
            muendlich: ['D', 'Schwerpunkt', 'Wahlfach']
        },
        {
            key: '3',
            label: VARIANT_LABELS[3],
            schriftlich: ['D', 'Englisch', 'AM', 'Fachtheorie (SWP)'],
            muendlich: ['Schwerpunkt', 'Wahlfach']
        }
    ],
    defaultWahlfaecher: WAHL_GENERIC,
    defaultSeminare: [],
    lfsHints: ['EN', 'Englisch'],
    lfsFieldLabel: 'Englisch / LFS',
    extraColumnNames: ['LFS', 'LehrerLFS', 'SchwerpunktFach', 'LehrerSchwerpunkt'],
    formSections: [
        { displayname: 'Person & Klasse', fields: ['Klasse', 'Nachname', 'Vorname'] },
        {
            displayname: 'Diplomarbeit',
            fields: ['Titel Diplomarbeit', 'BetreuungslehrerIn Diplomarbeit']
        },
        {
            displayname: 'Prüfungsvariante',
            fields: [
                'Variante',
                'Englisch / LFS',
                'LehrerIn LFS',
                'Schwerpunktfach',
                'LehrerIn Schwerpunkt'
            ]
        },
        {
            displayname: 'Wahlfach',
            fields: ['Wahlfach mündlich', 'Seminar', 'LehrerIn Wahlfach']
        },
        { displayname: 'Bestätigung', fields: ['Bestätigung Anmeldung'] }
    ],
    notes: 'Quelle: BMB Teilprüfungen + typische HTL-Varianten (D/E/AM/SWP).'
};

/** @type {SrdpProfile} */
export const PROFILE_HLW = {
    id: 'hlw',
    label: 'HLW',
    examShort: 'sRDP',
    listTitleCode: 'HLW',
    arbeitLabel: 'Diplomarbeit',
    variants: [
        {
            key: '1',
            label: VARIANT_LABELS[1],
            schriftlich: ['D', '2 aus LFS / AM / BW-RW'],
            muendlich: ['Fachkolloquium', 'Wahlfach', 'nicht gewähltes Klausurfach']
        },
        {
            key: '2',
            label: VARIANT_LABELS[2],
            schriftlich: ['D', 'LFS', 'AM', 'BW-RW'],
            muendlich: ['Fachkolloquium', 'Wahlfach']
        }
    ],
    defaultWahlfaecher: WAHL_HLW,
    defaultSeminare: [],
    lfsHints: LFS_HINTS,
    lfsFieldLabel: 'LFS',
    extraColumnNames: [
        'LFS',
        'LehrerLFS',
        'KlausurKombi',
        'Fachkolloquium',
        'LehrerFachkolloquium'
    ],
    formSections: [
        { displayname: 'Person & Klasse', fields: ['Klasse', 'Nachname', 'Vorname'] },
        {
            displayname: 'Diplomarbeit',
            fields: ['Titel Diplomarbeit', 'BetreuungslehrerIn Diplomarbeit']
        },
        {
            displayname: 'Prüfungsvariante',
            fields: [
                'Variante',
                'Klausur-Kombination',
                'LFS',
                'LehrerIn LFS',
                'Fachkolloquium',
                'LehrerIn Fachkolloquium'
            ]
        },
        {
            displayname: 'Wahlfach',
            fields: ['Wahlfach mündlich', 'Seminar', 'LehrerIn Wahlfach']
        },
        { displayname: 'Bestätigung', fields: ['Bestätigung Anmeldung'] }
    ],
    notes: 'HUM/HLW: D Pflicht; 2 oder 3 aus LFS/AM/BW-RW; mündlich Fachkolloquium + Wahlfach.'
};

/** @type {SrdpProfile} */
export const PROFILE_BAFEB = {
    id: 'bafeb',
    label: 'BAFEB',
    examShort: 'sRDP',
    listTitleCode: 'BAFEB',
    arbeitLabel: 'Diplomarbeit',
    variants: [
        {
            key: '1',
            label: VARIANT_LABELS[1],
            schriftlich: ['D', 'AM', 'Englisch', 'Fachtheorie (Did/Päd)'],
            muendlich: ['Fachtheorie mündlich', 'Wahlfach / berufsspezifisch']
        },
        {
            key: '2',
            label: VARIANT_LABELS[2],
            schriftlich: ['D', 'AM oder Englisch', 'Fachtheorie (Did/Päd)'],
            muendlich: ['Fachtheorie mündlich', 'AM oder Englisch', 'Wahlfach']
        }
    ],
    defaultWahlfaecher: WAHL_BAFEB,
    defaultSeminare: [],
    lfsHints: ['EN', 'Englisch'],
    lfsFieldLabel: 'Englisch / LFS',
    extraColumnNames: [
        'LFS',
        'LehrerLFS',
        'Fachtheorie',
        'LehrerFachtheorie',
        'MuendlichExtra'
    ],
    formSections: [
        { displayname: 'Person & Klasse', fields: ['Klasse', 'Nachname', 'Vorname'] },
        {
            displayname: 'Diplomarbeit',
            fields: ['Titel Diplomarbeit', 'BetreuungslehrerIn Diplomarbeit']
        },
        {
            displayname: 'Prüfungsvariante',
            fields: [
                'Variante',
                'Englisch / LFS',
                'LehrerIn LFS',
                'Fachtheorie',
                'LehrerIn Fachtheorie',
                'Mündlich Extra (AM/EN)'
            ]
        },
        {
            displayname: 'Wahlfach',
            fields: ['Wahlfach mündlich', 'Seminar', 'LehrerIn Wahlfach']
        },
        { displayname: 'Bestätigung', fields: ['Bestätigung Anmeldung'] }
    ],
    notes: 'BAfEP/BASOP: Fachtheorie Didaktik/Pädagogik; 3s+3m oder 4s+2m.'
};

/** @type {SrdpProfile} */
export const PROFILE_AHS = {
    id: 'ahs',
    label: 'AHS',
    examShort: 'sRP',
    listTitleCode: 'AHS',
    arbeitLabel: 'Abschließende Arbeit (ABA)',
    variants: [
        {
            key: '1',
            label: VARIANT_LABELS[1],
            schriftlich: ['D', 'M', 'LFS'],
            muendlich: ['mündlich 1', 'mündlich 2', 'mündlich 3']
        },
        {
            key: '2',
            label: VARIANT_LABELS[2],
            schriftlich: ['D', 'M', 'LFS', '4. Klausur'],
            muendlich: ['mündlich 1', 'mündlich 2']
        },
        {
            key: '3',
            label: VARIANT_LABELS[3],
            schriftlich: ['D', 'M', 'LFS', 'ggf. 4. Klausur'],
            muendlich: ['mündlich 1', 'mündlich 2', 'mündlich 3', '(ohne ABA: zusätzliche Prüfung)']
        }
    ],
    defaultWahlfaecher: WAHL_GENERIC,
    defaultSeminare: [],
    lfsHints: ['EN', 'FR', 'IT', 'SP', 'Latein', 'Griechisch'],
    lfsFieldLabel: 'Lebende Fremdsprache',
    extraColumnNames: [
        'HatABA',
        'TitelABA',
        'BetreuungslehrerABA',
        'LFS',
        'LehrerLFS',
        'VierteKlausur',
        'Muendlich1',
        'Muendlich2',
        'Muendlich3',
        'LehrerMuendlich1',
        'LehrerMuendlich2',
        'LehrerMuendlich3'
    ],
    formSections: [
        { displayname: 'Person & Klasse', fields: ['Klasse', 'Nachname', 'Vorname'] },
        {
            displayname: 'Abschließende Arbeit',
            fields: [
                'Mit ABA',
                'Titel ABA / VWA',
                'BetreuungslehrerIn ABA'
            ]
        },
        {
            displayname: 'Prüfungsvariante',
            fields: [
                'Variante',
                'Lebende Fremdsprache',
                'LehrerIn LFS',
                '4. Klausur',
                'Mündlich 1',
                'LehrerIn mündlich 1',
                'Mündlich 2',
                'LehrerIn mündlich 2',
                'Mündlich 3',
                'LehrerIn mündlich 3'
            ]
        },
        {
            displayname: 'Wahlfach / Seminar',
            fields: ['Wahlfach mündlich', 'Seminar', 'LehrerIn Wahlfach']
        },
        { displayname: 'Bestätigung', fields: ['Bestätigung Anmeldung'] }
    ],
    notes: 'AHS-SRDP: ABA bis 2028/29 optional; 3+3 oder 4+2 Klausuren/mündlich.'
};

/** @type {Record<string, SrdpProfile>} */
export const PROFILES = {
    hak: PROFILE_HAK,
    htl: PROFILE_HTL,
    hlw: PROFILE_HLW,
    bafeb: PROFILE_BAFEB,
    ahs: PROFILE_AHS
};

export const PROFILE_IDS = Object.keys(PROFILES);

/**
 * @param {string} [id]
 * @returns {SrdpProfile}
 */
export function getProfile(id) {
    const key = String(id || 'hak')
        .trim()
        .toLowerCase();
    return PROFILES[key] || PROFILE_HAK;
}

/**
 * @param {string|number} terminJahr
 * @param {string} [profileId]
 */
export function listTitleForYear(terminJahr, profileId) {
    const y = String(terminJahr || '').trim();
    if (!/^\d{4}$/.test(y)) throw new Error('Terminjahr muss vierstellig sein (z. B. 2026).');
    const p = getProfile(profileId);
    const prefix = p.examShort === 'sRP' ? 'sRP-Anmeldungen' : 'sRDP-Anmeldungen';
    return prefix + ' ' + p.listTitleCode + ' ' + y;
}

/**
 * Auch Legacy-Titel ohne Schulform-Code (nur HAK alt).
 * @param {string|number} terminJahr
 * @param {string} [profileId]
 * @returns {string[]}
 */
export function listTitleCandidates(terminJahr, profileId) {
    const primary = listTitleForYear(terminJahr, profileId);
    const p = getProfile(profileId);
    const out = [primary];
    if (p.id === 'hak') {
        const y = String(terminJahr || '').trim();
        out.push('sRDP-Anmeldungen ' + y);
    }
    return out;
}
