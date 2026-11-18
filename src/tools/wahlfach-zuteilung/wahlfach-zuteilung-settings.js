/**
 * Einstellungen Wahlfach-Zuteilung (Site-URL, Listen) – lokal, für Setup & Sync.
 */

import { LIST_TITLES } from './wahlfach-zuteilung-schema.js';
import {
    patchSchoolToolSitesSetup,
    resolveSchoolToolSiteUrl
} from '../../shared/school-tool-sites.js';

export const SETTINGS_KEY = 'ms365-wahlfach-zuteilung-settings-v1';

/**
 * @typedef {{
 *   siteUrl: string,
 *   listRunden: string,
 *   listAngebote: string,
 *   listAnmeldungen: string,
 *   listZuteilungen: string,
 *   accessHinweis: string
 * }} WzSettings
 */

/** @returns {WzSettings} */
export function defaultSettings() {
    return {
        siteUrl: '',
        listRunden: LIST_TITLES.runden,
        listAngebote: LIST_TITLES.angebote,
        listAnmeldungen: LIST_TITLES.anmeldungen,
        listZuteilungen: LIST_TITLES.zuteilungen,
        accessHinweis:
            'Ergebnis-CSV für Import in die lokale Schülerakte (Access). Spalten an Auhof-Schema anpassen, sobald Screenshots vorliegen.'
    };
}

/**
 * @returns {WzSettings}
 */
export function loadSettings() {
    const base = defaultSettings();
    try {
        const raw = localStorage.getItem(SETTINGS_KEY);
        if (!raw) {
            try {
                base.siteUrl = String(resolveSchoolToolSiteUrl('wahlfach') || '').trim();
            } catch {
                /* ignore */
            }
            return base;
        }
        const o = JSON.parse(raw) || {};
        let siteUrl = String(o.siteUrl || base.siteUrl || '').trim();
        if (!siteUrl) {
            try {
                siteUrl = String(resolveSchoolToolSiteUrl('wahlfach') || '').trim();
            } catch {
                /* ignore */
            }
        }
        return {
            siteUrl: siteUrl,
            listRunden: String(o.listRunden || base.listRunden).trim() || base.listRunden,
            listAngebote: String(o.listAngebote || base.listAngebote).trim() || base.listAngebote,
            listAnmeldungen:
                String(o.listAnmeldungen || base.listAnmeldungen).trim() || base.listAnmeldungen,
            listZuteilungen:
                String(o.listZuteilungen || base.listZuteilungen).trim() || base.listZuteilungen,
            accessHinweis: String(o.accessHinweis != null ? o.accessHinweis : base.accessHinweis)
        };
    } catch {
        return base;
    }
}

/**
 * @param {Partial<WzSettings>} patch
 * @returns {WzSettings}
 */
export function saveSettings(patch) {
    const next = { ...loadSettings(), ...(patch || {}) };
    next.siteUrl = String(next.siteUrl || '').trim();
    next.listRunden = String(next.listRunden || '').trim() || LIST_TITLES.runden;
    next.listAngebote = String(next.listAngebote || '').trim() || LIST_TITLES.angebote;
    next.listAnmeldungen = String(next.listAnmeldungen || '').trim() || LIST_TITLES.anmeldungen;
    next.listZuteilungen = String(next.listZuteilungen || '').trim() || LIST_TITLES.zuteilungen;
    try {
        localStorage.setItem(SETTINGS_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    try {
        patchSchoolToolSitesSetup({ planerSiteUrls: { wahlfach: next.siteUrl } });
    } catch {
        /* ignore */
    }
    return next;
}
