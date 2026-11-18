/**
 * Feldkatalog + Presets (Kern / Erweitert / Minimal).
 * Steuert das Eltern-Formular; Export mappt auf alle Einschreiber-Spalten.
 */

/** @typedef {'aus' | 'sichtbar' | 'pflicht'} FieldMode */
/** @typedef {{ id: string, label: string, section: string, type?: string, optionsKey?: string, einschreiber?: string }} FieldDef */

/** @type {FieldDef[]} */
export const FIELD_CATALOG = [
    { id: 'schueler.familienname', label: 'Familienname', section: 'schueler', einschreiber: 'Familienname' },
    { id: 'schueler.vorname', label: 'Vorname', section: 'schueler', einschreiber: 'erster Vorname' },
    { id: 'schueler.weitereVornamen', label: 'Weitere Vornamen', section: 'schueler', einschreiber: 'weitere Vornamen' },
    {
        id: 'schueler.geschlecht',
        label: 'Geschlecht',
        section: 'schueler',
        type: 'select',
        optionsKey: 'geschlecht',
        einschreiber: 'Geschlecht'
    },
    { id: 'schueler.geburtsdatum', label: 'Geburtsdatum', section: 'schueler', type: 'date', einschreiber: 'Geburtsdatum' },
    { id: 'schueler.geburtsort', label: 'Geburtsort', section: 'schueler', einschreiber: 'Geburtsort' },
    { id: 'schueler.geburtsstaat', label: 'Geburtsstaat', section: 'schueler', einschreiber: 'Geburtsstaat' },
    {
        id: 'schueler.staatsbuergerschaft',
        label: 'Staatsbürgerschaft',
        section: 'schueler',
        type: 'select',
        optionsKey: 'staatsbuergerschaft',
        einschreiber: 'Staatsbürgerschaft'
    },
    {
        id: 'schueler.religion',
        label: 'Religionsbekenntnis',
        section: 'schueler',
        type: 'select',
        optionsKey: 'religion',
        einschreiber: 'Religionsbekenntnis'
    },
    {
        id: 'schueler.alltagssprache',
        label: 'Alltagssprache',
        section: 'schueler',
        type: 'select',
        optionsKey: 'sprache',
        einschreiber: 'Alltagssprache'
    },
    {
        id: 'schueler.erstsprache',
        label: 'Erstsprache',
        section: 'schueler',
        type: 'select',
        optionsKey: 'sprache',
        einschreiber: 'Erstsprache'
    },
    { id: 'schueler.svnr', label: 'Sozialversicherungsnummer', section: 'schueler', einschreiber: 'SVNr m. Formel' },
    { id: 'schueler.mail', label: 'E-Mail (Schüler:in)', section: 'schueler', type: 'email', einschreiber: 'Mailadresse' },
    { id: 'schueler.mobil', label: 'Mobiltelefon (Schüler:in)', section: 'schueler', type: 'tel', einschreiber: 'Mobiltelefon' },

    { id: 'wohnadresse.strasse', label: 'Straße', section: 'adresse', einschreiber: 'Straße' },
    { id: 'wohnadresse.hausnummer', label: 'Hausnummer', section: 'adresse', einschreiber: 'Hausnummer' },
    { id: 'wohnadresse.plz', label: 'PLZ', section: 'adresse', einschreiber: 'PLZ' },
    { id: 'wohnadresse.ort', label: 'Ort', section: 'adresse', einschreiber: 'Ort' },
    { id: 'wohnadresse.land', label: 'Land', section: 'adresse', type: 'select', optionsKey: 'land', einschreiber: 'Land' },
    { id: 'wohnadresse.gemeindenummer', label: 'Gemeindenummer', section: 'adresse', einschreiber: 'Gemeindenummer' },

    {
        id: 'erziehungsberechtigte.0.adressart',
        label: 'Beziehung (1)',
        section: 'erz1',
        type: 'select',
        optionsKey: 'adressart'
    },
    { id: 'erziehungsberechtigte.0.anrede', label: 'Anrede (1)', section: 'erz1' },
    { id: 'erziehungsberechtigte.0.vorname', label: 'Vorname (1)', section: 'erz1' },
    { id: 'erziehungsberechtigte.0.familienname', label: 'Familienname (1)', section: 'erz1' },
    { id: 'erziehungsberechtigte.0.mobil', label: 'Mobiltelefon (1)', section: 'erz1', type: 'tel' },
    { id: 'erziehungsberechtigte.0.mail', label: 'E-Mail (1)', section: 'erz1', type: 'email' },
    {
        id: 'erziehungsberechtigte.0.erziehungsberechtigt',
        label: 'Erziehungsberechtigt (1)',
        section: 'erz1',
        type: 'select',
        optionsKey: 'jaNein'
    },
    {
        id: 'erziehungsberechtigte.0.zahlungspflichtig',
        label: 'Zahlungspflichtig (1)',
        section: 'erz1',
        type: 'select',
        optionsKey: 'jaNein'
    },
    {
        id: 'erziehungsberechtigte.0.entscheideAn',
        label: 'Entscheide an (1)',
        section: 'erz1',
        type: 'select',
        optionsKey: 'jaNein'
    },

    {
        id: 'erziehungsberechtigte.1.adressart',
        label: 'Beziehung (2)',
        section: 'erz2',
        type: 'select',
        optionsKey: 'adressart'
    },
    { id: 'erziehungsberechtigte.1.anrede', label: 'Anrede (2)', section: 'erz2' },
    { id: 'erziehungsberechtigte.1.vorname', label: 'Vorname (2)', section: 'erz2' },
    { id: 'erziehungsberechtigte.1.familienname', label: 'Familienname (2)', section: 'erz2' },
    { id: 'erziehungsberechtigte.1.mobil', label: 'Mobiltelefon (2)', section: 'erz2', type: 'tel' },
    { id: 'erziehungsberechtigte.1.mail', label: 'E-Mail (2)', section: 'erz2', type: 'email' },
    {
        id: 'erziehungsberechtigte.1.erziehungsberechtigt',
        label: 'Erziehungsberechtigt (2)',
        section: 'erz2',
        type: 'select',
        optionsKey: 'jaNein'
    },
    {
        id: 'erziehungsberechtigte.1.zahlungspflichtig',
        label: 'Zahlungspflichtig (2)',
        section: 'erz2',
        type: 'select',
        optionsKey: 'jaNein'
    },
    {
        id: 'erziehungsberechtigte.1.entscheideAn',
        label: 'Entscheide an (2)',
        section: 'erz2',
        type: 'select',
        optionsKey: 'jaNein'
    }
];

export const SECTION_LABELS = Object.freeze({
    schueler: 'Schüler:in',
    adresse: 'Wohnadresse',
    erz1: 'Erziehungsberechtigte:r 1',
    erz2: 'Erziehungsberechtigte:r 2'
});

/** @type {Record<string, Record<string, FieldMode>>} */
export const FIELD_PRESETS = {
    minimal: {
        'schueler.familienname': 'pflicht',
        'schueler.vorname': 'pflicht',
        'schueler.geschlecht': 'pflicht',
        'schueler.geburtsdatum': 'pflicht',
        'schueler.svnr': 'pflicht',
        'wohnadresse.strasse': 'pflicht',
        'wohnadresse.hausnummer': 'pflicht',
        'wohnadresse.plz': 'pflicht',
        'wohnadresse.ort': 'pflicht',
        'erziehungsberechtigte.0.adressart': 'pflicht',
        'erziehungsberechtigte.0.vorname': 'pflicht',
        'erziehungsberechtigte.0.familienname': 'pflicht',
        'erziehungsberechtigte.0.mobil': 'pflicht',
        'erziehungsberechtigte.0.mail': 'sichtbar',
        'erziehungsberechtigte.1.adressart': 'sichtbar',
        'erziehungsberechtigte.1.vorname': 'sichtbar',
        'erziehungsberechtigte.1.familienname': 'sichtbar',
        'erziehungsberechtigte.1.mobil': 'sichtbar',
        'erziehungsberechtigte.1.mail': 'sichtbar'
    },
    kern: {
        'schueler.familienname': 'pflicht',
        'schueler.vorname': 'pflicht',
        'schueler.weitereVornamen': 'sichtbar',
        'schueler.geschlecht': 'pflicht',
        'schueler.geburtsdatum': 'pflicht',
        'schueler.geburtsort': 'sichtbar',
        'schueler.staatsbuergerschaft': 'pflicht',
        'schueler.religion': 'pflicht',
        'schueler.erstsprache': 'pflicht',
        'schueler.alltagssprache': 'sichtbar',
        'schueler.svnr': 'pflicht',
        'schueler.mobil': 'sichtbar',
        'schueler.mail': 'sichtbar',
        'wohnadresse.strasse': 'pflicht',
        'wohnadresse.hausnummer': 'pflicht',
        'wohnadresse.plz': 'pflicht',
        'wohnadresse.ort': 'pflicht',
        'wohnadresse.land': 'sichtbar',
        'erziehungsberechtigte.0.adressart': 'pflicht',
        'erziehungsberechtigte.0.anrede': 'sichtbar',
        'erziehungsberechtigte.0.vorname': 'pflicht',
        'erziehungsberechtigte.0.familienname': 'pflicht',
        'erziehungsberechtigte.0.mobil': 'pflicht',
        'erziehungsberechtigte.0.mail': 'pflicht',
        'erziehungsberechtigte.0.erziehungsberechtigt': 'sichtbar',
        'erziehungsberechtigte.0.zahlungspflichtig': 'sichtbar',
        'erziehungsberechtigte.0.entscheideAn': 'sichtbar',
        'erziehungsberechtigte.1.adressart': 'pflicht',
        'erziehungsberechtigte.1.anrede': 'sichtbar',
        'erziehungsberechtigte.1.vorname': 'pflicht',
        'erziehungsberechtigte.1.familienname': 'pflicht',
        'erziehungsberechtigte.1.mobil': 'pflicht',
        'erziehungsberechtigte.1.mail': 'pflicht',
        'erziehungsberechtigte.1.erziehungsberechtigt': 'sichtbar',
        'erziehungsberechtigte.1.zahlungspflichtig': 'sichtbar',
        'erziehungsberechtigte.1.entscheideAn': 'sichtbar'
    },
    erweitert: Object.fromEntries(FIELD_CATALOG.map((f) => [f.id, 'pflicht']))
};

/**
 * @param {string} [presetId]
 * @param {Record<string, FieldMode>} [overrides]
 */
export function resolveFieldProfile(presetId = 'kern', overrides = {}) {
    const base = { ...(FIELD_PRESETS[presetId] || FIELD_PRESETS.kern), ...(overrides || {}) };
    return FIELD_CATALOG.map((def) => {
        const mode = base[def.id] || 'aus';
        return {
            ...def,
            mode,
            visible: mode === 'sichtbar' || mode === 'pflicht',
            required: mode === 'pflicht'
        };
    });
}

/**
 * @param {ReturnType<typeof resolveFieldProfile>} profile
 */
export function visibleFields(profile) {
    return (profile || []).filter((f) => f.visible);
}

export const PRESET_OPTIONS = [
    { id: 'minimal', label: 'Minimal' },
    { id: 'kern', label: 'Kern (empfohlen)' },
    { id: 'erweitert', label: 'Erweitert' }
];
