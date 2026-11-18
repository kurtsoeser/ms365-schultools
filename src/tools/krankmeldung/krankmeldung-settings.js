/**
 * Krankmeldung-Einstellungen + SharePoint-Site.
 */

import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';
import { LIST_TITLES } from './krankmeldung-schema.js';

const KEY = 'ms365-krankmeldung-settings-v1';
export const SP_SETTINGS_KEY = 'ms365-krankmeldung-sp-settings-v1';

/**
 * @returns {{
 *   schoolName: string,
 *   apiBaseUrl: string,
 *   adminKey: string,
 *   kvHintEmail: string,
 *   publicKlassen: string,
 *   allowFreeKlasse: boolean,
 *   allowAnhang: boolean,
 *   anhangPflicht: boolean
 * }}
 */
export function loadKrankmeldungSettings() {
    try {
        const raw = localStorage.getItem(KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            schoolName: String(o.schoolName || '').trim(),
            apiBaseUrl: String(o.apiBaseUrl || '')
                .trim()
                .replace(/\/+$/, ''),
            adminKey: String(o.adminKey || '').trim(),
            kvHintEmail: String(o.kvHintEmail || '').trim(),
            publicKlassen: String(o.publicKlassen || '').trim(),
            allowFreeKlasse: o.allowFreeKlasse !== false,
            allowAnhang: o.allowAnhang !== false,
            anhangPflicht: o.anhangPflicht === true
        };
    } catch {
        return {
            schoolName: '',
            apiBaseUrl: '',
            adminKey: '',
            kvHintEmail: '',
            publicKlassen: '',
            allowFreeKlasse: true,
            allowAnhang: true,
            anhangPflicht: false
        };
    }
}

/**
 * @param {Partial<ReturnType<typeof loadKrankmeldungSettings>>} partial
 */
export function saveKrankmeldungSettings(partial) {
    const next = { ...loadKrankmeldungSettings(), ...partial };
    localStorage.setItem(KEY, JSON.stringify(next));
    return next;
}

/**
 * @returns {{ siteUrl: string, listMeldungen: string }}
 */
export function loadKrankmeldungSpSettings() {
    try {
        const raw = localStorage.getItem(SP_SETTINGS_KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            siteUrl: String((o && o.siteUrl) || '').trim(),
            listMeldungen:
                String((o && o.listMeldungen) || LIST_TITLES.meldungen).trim() ||
                LIST_TITLES.meldungen
        };
    } catch {
        return { siteUrl: '', listMeldungen: LIST_TITLES.meldungen };
    }
}

/**
 * @param {Partial<{ siteUrl: string, listMeldungen: string }>} partial
 */
export function saveKrankmeldungSpSettings(partial) {
    const cur = loadKrankmeldungSpSettings();
    const next = {
        siteUrl: partial.siteUrl != null ? String(partial.siteUrl).trim() : cur.siteUrl,
        listMeldungen:
            partial.listMeldungen != null
                ? String(partial.listMeldungen).trim() || cur.listMeldungen
                : cur.listMeldungen
    };
    try {
        localStorage.setItem(SP_SETTINGS_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

/**
 * Effektive Site-URL: Tool-Settings → Schul-Tool-Sites → Default.
 */
export function resolveKrankmeldungSiteUrl() {
    return resolveSchoolToolSiteUrl('krankmeldung') || '';
}
