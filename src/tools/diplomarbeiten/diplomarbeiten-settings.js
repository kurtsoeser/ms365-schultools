/**
 * Diplomarbeiten / VWA – Betriebsmodus und Typ.
 * Der Schlüssel liegt im Schul-Paket „school-apps-settings“ der IT-Bibliothek.
 */
import { notifyAppLocalDataChanged } from '../../shared/app-local-data-notify.js';
import {
    CONFIG_BUNDLE_KIND,
    CONFIG_BUNDLE_SCHEMA_VERSION,
    fingerprintLocalStorageSlice
} from '../../shared/stammdaten-sharepoint-config-bundle.js';

export const DIPLOM_SETTINGS_KEY = 'ms365-diplomarbeiten-settings-v1';
const KEY = DIPLOM_SETTINGS_KEY;

/** @typedef {'diplom'|'vwa'} ArbeitTyp */
/** @typedef {'channels'|'legacy'} Betriebsmodus */

/**
 * @returns {{ arbeitTyp: ArbeitTyp, modus: Betriebsmodus, setupDone: boolean }}
 */
export function defaultDiplomSettings() {
    return {
        arbeitTyp: 'diplom',
        modus: 'channels',
        setupDone: false
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeDiplomSettings(raw) {
    const d = defaultDiplomSettings();
    const r = raw && typeof raw === 'object' ? raw : {};
    const typ = String(r.arbeitTyp || d.arbeitTyp).toLowerCase();
    const modus = String(r.modus || d.modus).toLowerCase();
    return {
        arbeitTyp: typ === 'vwa' ? 'vwa' : 'diplom',
        modus: modus === 'legacy' ? 'legacy' : 'channels',
        setupDone: !!r.setupDone
    };
}

export function loadDiplomSettings() {
    try {
        if (typeof localStorage === 'undefined') return defaultDiplomSettings();
        const raw = localStorage.getItem(KEY);
        if (!raw) return defaultDiplomSettings();
        return normalizeDiplomSettings(JSON.parse(raw));
    } catch {
        return defaultDiplomSettings();
    }
}

/**
 * @param {Partial<ReturnType<typeof defaultDiplomSettings>>} partial
 */
export function saveDiplomSettings(partial) {
    const next = normalizeDiplomSettings({ ...loadDiplomSettings(), ...(partial || {}) });
    try {
        if (typeof localStorage !== 'undefined') {
            localStorage.setItem(KEY, JSON.stringify(next));
        }
    } catch {
        /* ignore */
    }
    notifyAppLocalDataChanged('diplomarbeiten-settings');
    return next;
}

/**
 * Eine Schul-Datei fortschreiben, ohne die übrigen Schlüssel darin zu ersetzen.
 * @param {unknown} part
 * @param {string} settingsJson
 */
export function embedSettingsInSchoolAppsPart(part, settingsJson) {
    const src = part && typeof part === 'object' ? /** @type {Record<string, unknown>} */ (part) : {};
    const prevLs =
        src.localStorage && typeof src.localStorage === 'object'
            ? /** @type {Record<string, string>} */ (src.localStorage)
            : {};
    const localStorageSlice = Object.assign({}, prevLs, {
        [KEY]: String(settingsJson || '')
    });
    return {
        schemaVersion: CONFIG_BUNDLE_SCHEMA_VERSION,
        bundleKind: CONFIG_BUNDLE_KIND,
        bundlePart: 'school-apps-settings',
        exportedAt: new Date().toISOString(),
        schoolName: src.schoolName || null,
        domain: src.domain || null,
        keyCount: Object.keys(localStorageSlice).length,
        fingerprint: fingerprintLocalStorageSlice(localStorageSlice),
        localStorage: localStorageSlice
    };
}
