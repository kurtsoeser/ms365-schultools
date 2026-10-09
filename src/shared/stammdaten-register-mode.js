/**
 * Stammdaten: Tab-Hashes (#klassen, #intranet, …).
 * Legacy-Hashes (#sync, #import, #pflegen) werden auf Tabs gemappt.
 */

const LEGACY_HASH_TO_TAB = {
    sync: 'tabMainStammdaten',
    synchron: 'tabMainStammdaten',
    synchronisieren: 'tabMainStammdaten',
    intranet: 'tabMainStammdaten',
    import: 'tabMainSchueler',
    einspielen: 'tabMainSchueler',
    importieren: 'tabMainSchueler',
    pflegen: '',
    pflege: '',
    bearbeiten: ''
};

const TAB_HASH_TO_BTN = {
    stammdaten: 'tabMainStammdaten',
    faecher: 'tabMainSubjects',
    subjects: 'tabMainSubjects',
    arge: 'tabMainArges',
    arges: 'tabMainArges',
    verwaltung: 'tabMainVerwaltung',
    sga: 'tabMainSga',
    lehrer: 'tabMainLehrer',
    schueler: 'tabMainSchueler',
    schuelervertretung: 'tabMainSchuelervertretung',
    klassen: 'tabMainKlassen',
    classes: 'tabMainKlassen',
    panelclasses: 'tabMainKlassen',
    datenlandkarte: 'tabMainDatenlandkarte',
    landkarte: 'tabMainDatenlandkarte',
    'daten-landkarte': 'tabMainDatenlandkarte',
    intranet: 'tabMainStammdaten',
    werkzeuge: 'tabMainStammdaten'
};

const BTN_TO_TAB_HASH = {
    tabMainStammdaten: 'stammdaten',
    tabMainSubjects: 'faecher',
    tabMainArges: 'arge',
    tabMainVerwaltung: 'verwaltung',
    tabMainSga: 'sga',
    tabMainLehrer: 'lehrer',
    tabMainSchueler: 'schueler',
    tabMainSchuelervertretung: 'schuelervertretung',
    tabMainKlassen: 'klassen',
    tabMainDatenlandkarte: 'datenlandkarte'
};

/**
 * @param {string} [hash]
 * @returns {{ tabBtnId: string }}
 */
export function parseRegisterHash(hash) {
    const raw = String(hash || '')
        .replace(/^#/, '')
        .trim()
        .toLowerCase();
    if (!raw) {
        return { tabBtnId: '' };
    }
    if (Object.prototype.hasOwnProperty.call(LEGACY_HASH_TO_TAB, raw)) {
        return { tabBtnId: LEGACY_HASH_TO_TAB[raw] || '' };
    }
    const tabBtnId = TAB_HASH_TO_BTN[raw] || '';
    return { tabBtnId };
}

/**
 * @param {string} [tabBtnId]
 * @returns {string} Hash ohne #
 */
export function buildRegisterHash(tabBtnId) {
    const tab = String(tabBtnId || '').trim();
    if (tab && BTN_TO_TAB_HASH[tab]) return BTN_TO_TAB_HASH[tab];
    return '';
}

/** Legacy-Modus-Hashes (werden von Tab-Skript separat aufgelöst). */
export function isRegisterModeHash(raw) {
    const k = String(raw || '')
        .replace(/^#/, '')
        .trim()
        .toLowerCase();
    return Object.prototype.hasOwnProperty.call(LEGACY_HASH_TO_TAB, k);
}

export { TAB_HASH_TO_BTN, BTN_TO_TAB_HASH };
