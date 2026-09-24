/**
 * SharePoint-Schema für sRDP/sRP-Anmeldungen (schulformabhängig).
 */
import { VARIANT_LABELS, STORAGE_KEY } from './srdp-anmeldung-schema-constants.js';
import {
    getProfile,
    listTitleForYear,
    listTitleCandidates,
    PROFILE_HAK,
    PROFILE_HTL,
    PROFILE_HLW,
    PROFILE_BAFEB,
    PROFILE_AHS,
    PROFILES,
    PROFILE_IDS
} from './srdp-anmeldung-profiles.js';

export {
    VARIANT_LABELS,
    STORAGE_KEY,
    getProfile,
    listTitleForYear,
    listTitleCandidates,
    PROFILE_HAK,
    PROFILE_HTL,
    PROFILE_HLW,
    PROFILE_BAFEB,
    PROFILE_AHS,
    PROFILES,
    PROFILE_IDS
};

/** @typedef {{ name: string, displayName: string, required?: boolean, [k: string]: unknown }} GraphColumnDef */
/** @typedef {{ id: string, title: string, groupBy: string, orderBy: string, fields: string[] }} SrdpViewDef */

/** @deprecated use PROFILE_HAK.variants */
export const HAK_VARIANTS = PROFILE_HAK.variants;
/** @deprecated */
export const HAK_DEFAULT_WAHLFAECHER = PROFILE_HAK.defaultWahlfaecher;
/** @deprecated */
export const HAK_DEFAULT_SEMINARE = PROFILE_HAK.defaultSeminare;
/** @deprecated */
export const HAK_LFS_HINTS = PROFILE_HAK.lfsHints;

/**
 * @param {{
 *   profileId?: string,
 *   klassen?: string[],
 *   lehrer?: string[],
 *   lfs?: string[],
 *   wahlfaecher?: string[],
 *   seminare?: string[],
 *   terminJahr?: string|number,
 *   schwerpunkte?: string[],
 *   fachkolloquien?: string[],
 *   fachtheorien?: string[],
 *   vierteKlausuren?: string[],
 *   muendlichFaecher?: string[]
 * }} [opts]
 * @returns {GraphColumnDef[]}
 */
export function buildAnmeldungColumns(opts) {
    const o = opts || {};
    const profile = getProfile(o.profileId);
    const klassen = uniqStrings(o.klassen);
    const lehrer = uniqStrings(o.lehrer);
    const lfs = uniqStrings(o.lfs);
    const wahlfaecher = uniqStrings(o.wahlfaecher);
    const seminare = uniqStrings(o.seminare);
    const schwerpunkte = uniqStrings(o.schwerpunkte);
    const fachkolloquien = uniqStrings(o.fachkolloquien);
    const fachtheorien = uniqStrings(o.fachtheorien);
    const vierteKlausuren = uniqStrings(o.vierteKlausuren);
    const muendlichFaecher = uniqStrings(o.muendlichFaecher);
    const lehrerChoices = lehrer.length ? lehrer : ['—'];
    const variantLabels = profile.variants.map((v) => v.label);

    /** @type {GraphColumnDef[]} */
    const cols = [
        {
            name: 'Klasse',
            displayName: 'Klasse',
            choice: { allowTextEntry: false, choices: klassen.length ? klassen : ['—'] }
        },
        {
            name: 'Nachname',
            displayName: 'Nachname',
            text: { allowMultipleLines: false, maxLength: 100 }
        },
        {
            name: 'Vorname',
            displayName: 'Vorname',
            text: { allowMultipleLines: false, maxLength: 100 }
        }
    ];

    if (profile.id === 'ahs') {
        cols.push(
            { name: 'HatABA', displayName: 'Mit ABA', boolean: {} },
            {
                name: 'TitelABA',
                displayName: 'Titel ABA / VWA',
                text: { allowMultipleLines: true, maxLength: 4000 }
            },
            {
                name: 'BetreuungslehrerABA',
                displayName: 'BetreuungslehrerIn ABA',
                choice: { allowTextEntry: false, choices: lehrerChoices }
            }
        );
    } else {
        cols.push(
            {
                name: 'TitelDiplomarbeit',
                displayName: 'Titel Diplomarbeit',
                text: { allowMultipleLines: true, maxLength: 4000 }
            },
            {
                name: 'BetreuungslehrerDA',
                displayName: 'BetreuungslehrerIn Diplomarbeit',
                choice: { allowTextEntry: false, choices: lehrerChoices }
            }
        );
    }

    cols.push({
        name: 'Variante',
        displayName: 'Variante',
        choice: { allowTextEntry: false, choices: variantLabels.length ? variantLabels : [VARIANT_LABELS[1]] }
    });

    if (profile.extraColumnNames.indexOf('KlausurKombi') !== -1) {
        cols.push({
            name: 'KlausurKombi',
            displayName: 'Klausur-Kombination',
            choice: {
                allowTextEntry: true,
                choices: [
                    'LFS + AM (BW-RW mündlich)',
                    'LFS + BW-RW (AM mündlich)',
                    'AM + BW-RW (LFS mündlich)',
                    'alle vier (Variante 2)',
                    '— bitte wählen oder eingeben —'
                ]
            }
        });
    }

    if (profile.extraColumnNames.indexOf('LFS') !== -1) {
        cols.push(
            {
                name: 'LFS',
                displayName: profile.lfsFieldLabel || 'LFS',
                choice: {
                    allowTextEntry: false,
                    choices: lfs.length ? lfs : (profile.lfsHints || []).slice()
                }
            },
            {
                name: 'LehrerLFS',
                displayName: 'LehrerIn LFS',
                choice: { allowTextEntry: false, choices: lehrerChoices }
            }
        );
    }

    if (profile.extraColumnNames.indexOf('LehrerBKO') !== -1) {
        cols.push({
            name: 'LehrerBKO',
            displayName: 'LehrerIn BKO mündlich',
            choice: { allowTextEntry: false, choices: lehrerChoices }
        });
    }

    if (profile.extraColumnNames.indexOf('SchwerpunktFach') !== -1) {
        cols.push(
            {
                name: 'SchwerpunktFach',
                displayName: 'Schwerpunktfach',
                choice: {
                    allowTextEntry: true,
                    choices: schwerpunkte.length
                        ? schwerpunkte
                        : ['Technischer Schwerpunkt', '— bitte wählen oder eingeben —']
                }
            },
            {
                name: 'LehrerSchwerpunkt',
                displayName: 'LehrerIn Schwerpunkt',
                choice: { allowTextEntry: false, choices: lehrerChoices }
            }
        );
    }

    if (profile.extraColumnNames.indexOf('Fachkolloquium') !== -1) {
        cols.push(
            {
                name: 'Fachkolloquium',
                displayName: 'Fachkolloquium',
                choice: {
                    allowTextEntry: true,
                    choices: fachkolloquien.length
                        ? fachkolloquien
                        : (profile.defaultWahlfaecher || []).slice(0, 6).concat(['— bitte wählen oder eingeben —'])
                }
            },
            {
                name: 'LehrerFachkolloquium',
                displayName: 'LehrerIn Fachkolloquium',
                choice: { allowTextEntry: false, choices: lehrerChoices }
            }
        );
    }

    if (profile.extraColumnNames.indexOf('Fachtheorie') !== -1) {
        cols.push(
            {
                name: 'Fachtheorie',
                displayName: 'Fachtheorie',
                choice: {
                    allowTextEntry: true,
                    choices: fachtheorien.length
                        ? fachtheorien
                        : ['Didaktik', 'Pädagogik', 'Didaktik / Pädagogik', '— bitte wählen oder eingeben —']
                }
            },
            {
                name: 'LehrerFachtheorie',
                displayName: 'LehrerIn Fachtheorie',
                choice: { allowTextEntry: false, choices: lehrerChoices }
            }
        );
    }

    if (profile.extraColumnNames.indexOf('MuendlichExtra') !== -1) {
        cols.push({
            name: 'MuendlichExtra',
            displayName: 'Mündlich Extra (AM/EN)',
            choice: {
                allowTextEntry: true,
                choices: ['Angewandte Mathematik', 'Englisch', '— / nicht nötig —']
            }
        });
    }

    if (profile.extraColumnNames.indexOf('VierteKlausur') !== -1) {
        cols.push({
            name: 'VierteKlausur',
            displayName: '4. Klausur',
            choice: {
                allowTextEntry: true,
                choices: vierteKlausuren.length
                    ? vierteKlausuren
                    : [
                          'Weitere lebende Fremdsprache',
                          'Latein',
                          'Darstellende Geometrie',
                          '— / nicht gewählt —'
                      ]
            }
        });
    }

    if (profile.extraColumnNames.indexOf('Muendlich1') !== -1) {
        const mf = muendlichFaecher.length ? muendlichFaecher : (profile.defaultWahlfaecher || []).slice();
        cols.push(
            {
                name: 'Muendlich1',
                displayName: 'Mündlich 1',
                choice: { allowTextEntry: true, choices: mf.concat(['—']) }
            },
            {
                name: 'LehrerMuendlich1',
                displayName: 'LehrerIn mündlich 1',
                choice: { allowTextEntry: false, choices: lehrerChoices }
            },
            {
                name: 'Muendlich2',
                displayName: 'Mündlich 2',
                choice: { allowTextEntry: true, choices: mf.concat(['—']) }
            },
            {
                name: 'LehrerMuendlich2',
                displayName: 'LehrerIn mündlich 2',
                choice: { allowTextEntry: false, choices: lehrerChoices }
            },
            {
                name: 'Muendlich3',
                displayName: 'Mündlich 3',
                choice: { allowTextEntry: true, choices: mf.concat(['— / nicht nötig —']) }
            },
            {
                name: 'LehrerMuendlich3',
                displayName: 'LehrerIn mündlich 3',
                choice: { allowTextEntry: false, choices: lehrerChoices }
            }
        );
    }

    cols.push(
        {
            name: 'WahlfachMuendlich',
            displayName: 'Wahlfach mündlich',
            choice: {
                allowTextEntry: false,
                choices: wahlfaecher.length ? wahlfaecher : (profile.defaultWahlfaecher || []).slice()
            }
        },
        {
            name: 'Seminar',
            displayName: 'Seminar',
            choice: {
                allowTextEntry: true,
                choices: seminare.length ? seminare : ['— bitte wählen oder eingeben —']
            }
        },
        {
            name: 'LehrerWahlfach',
            displayName: 'LehrerIn Wahlfach',
            choice: { allowTextEntry: false, choices: lehrerChoices }
        },
        {
            name: 'Bestaetigung',
            displayName: 'Bestätigung Anmeldung',
            boolean: {}
        },
        {
            name: 'Schulform',
            displayName: 'Schulform',
            choice: { allowTextEntry: false, choices: [profile.label] }
        },
        {
            name: 'TerminJahr',
            displayName: 'Terminjahr',
            text: { allowMultipleLines: false, maxLength: 4 }
        },
        {
            name: 'PruefplanKurz',
            displayName: 'Prüfplan (Info)',
            text: { allowMultipleLines: true, maxLength: 500 }
        }
    );

    return cols;
}

/**
 * @param {string} [profileId]
 */
export function requiredColumnNamesForProfile(profileId) {
    return buildAnmeldungColumns({ profileId }).map((c) => c.name);
}

/** @deprecated – HAK-Kern; für Health besser requiredColumnNamesForProfile nutzen */
export const REQUIRED_COLUMN_NAMES = requiredColumnNamesForProfile('hak');

/**
 * @param {string} [profileId]
 * @returns {SrdpViewDef[]}
 */
export function defaultViewsForProfile(profileId) {
    const profile = getProfile(profileId);
    const arbeitTitle =
        profile.id === 'ahs' ? 'TitelABA' : 'TitelDiplomarbeit';
    const arbeitLehrer =
        profile.id === 'ahs' ? 'BetreuungslehrerABA' : 'BetreuungslehrerDA';
    const extras = profile.extraColumnNames.filter((n) =>
        ['LFS', 'LehrerLFS', 'LehrerBKO', 'SchwerpunktFach', 'LehrerSchwerpunkt', 'Fachkolloquium', 'Fachtheorie'].includes(
            n
        )
    );
    const fields = [
        'Klasse',
        'Nachname',
        'Vorname',
        'Variante',
        arbeitTitle,
        arbeitLehrer,
        ...extras,
        'WahlfachMuendlich',
        'LehrerWahlfach',
        'Seminar'
    ];
    return [
        {
            id: 'nach-klasse',
            title: 'Nach Klasse',
            groupBy: 'Klasse',
            orderBy: 'Nachname',
            fields
        },
        {
            id: 'nach-variante',
            title: 'Nach Variante',
            groupBy: 'Variante',
            orderBy: 'Nachname',
            fields
        }
    ];
}

export const DEFAULT_VIEWS = defaultViewsForProfile('hak');

/**
 * @param {GraphColumnDef} def
 */
export function toGraphColumnBody(def) {
    return { ...def };
}

/**
 * @param {unknown} list
 * @returns {string[]}
 */
function uniqStrings(list) {
    const seen = new Set();
    const out = [];
    (Array.isArray(list) ? list : []).forEach((raw) => {
        const s = String(raw == null ? '' : raw).trim();
        if (!s) return;
        const key = s.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(s);
    });
    return out;
}
