/**
 * Einstellungen Ausgabenverwaltung (App + SharePoint Site/Listen).
 */
import { LIST_TITLES } from './ausgaben-schema.js';
import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';

const APP_KEY = 'ms365-ausgaben-settings-v1';
export const SP_SETTINGS_KEY = 'ms365-ausgaben-sp-settings-v1';

/**
 * @returns {{ schoolName: string, defaultAusgegebenVon: string, adminEmails: string }}
 */
export function loadAusgabenSettings() {
    try {
        const raw = localStorage.getItem(APP_KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            schoolName: String(o.schoolName || '').trim(),
            defaultAusgegebenVon: String(o.defaultAusgegebenVon || '').trim(),
            adminEmails: String(o.adminEmails || '').trim()
        };
    } catch {
        return { schoolName: '', defaultAusgegebenVon: '', adminEmails: '' };
    }
}

/**
 * @param {Partial<ReturnType<typeof loadAusgabenSettings>>} partial
 */
export function saveAusgabenSettings(partial) {
    const next = { ...loadAusgabenSettings(), ...partial };
    localStorage.setItem(APP_KEY, JSON.stringify(next));
    return next;
}

/**
 * @returns {{ siteUrl: string, listGegenstaende: string, listVorgaenge: string }}
 */
export function loadAusgabenSpSettings() {
    try {
        const raw = localStorage.getItem(SP_SETTINGS_KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            siteUrl: String((o && o.siteUrl) || '').trim(),
            listGegenstaende:
                String((o && o.listGegenstaende) || LIST_TITLES.gegenstaende).trim() ||
                LIST_TITLES.gegenstaende,
            listVorgaenge:
                String((o && o.listVorgaenge) || LIST_TITLES.vorgaenge).trim() || LIST_TITLES.vorgaenge
        };
    } catch {
        return {
            siteUrl: '',
            listGegenstaende: LIST_TITLES.gegenstaende,
            listVorgaenge: LIST_TITLES.vorgaenge
        };
    }
}

/**
 * @param {Partial<{ siteUrl: string, listGegenstaende: string, listVorgaenge: string }>} partial
 */
export function saveAusgabenSpSettings(partial) {
    const cur = loadAusgabenSpSettings();
    const next = {
        siteUrl: partial.siteUrl != null ? String(partial.siteUrl).trim() : cur.siteUrl,
        listGegenstaende:
            partial.listGegenstaende != null
                ? String(partial.listGegenstaende).trim() || cur.listGegenstaende
                : cur.listGegenstaende,
        listVorgaenge:
            partial.listVorgaenge != null
                ? String(partial.listVorgaenge).trim() || cur.listVorgaenge
                : cur.listVorgaenge
    };
    try {
        localStorage.setItem(SP_SETTINGS_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

/**
 * Effektive Site-URL: Tool-Settings → Schul-Tool-Sites → Default/Intranet.
 */
export function resolveAusgabenSiteUrl() {
    return resolveSchoolToolSiteUrl('ausgaben') || '';
}
