/**
 * Einstellungen Kollegiumsumfragen (Anzeige + SharePoint-Site).
 */
import { LIST_TITLES } from './kollegium-umfragen-schema.js';
import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';

const KEY = 'ms365-kollegium-umfragen-settings-v1';
export const SP_SETTINGS_KEY = 'ms365-kollegium-umfragen-sp-settings-v1';

/**
 * @returns {{ schoolName: string, defaultAnonym: boolean }}
 */
export function loadKollegiumSettings() {
    try {
        const raw = localStorage.getItem(KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            schoolName: String(o.schoolName || '').trim(),
            defaultAnonym: !!o.defaultAnonym
        };
    } catch {
        return { schoolName: '', defaultAnonym: false };
    }
}

/**
 * @param {Partial<ReturnType<typeof loadKollegiumSettings>>} partial
 */
export function saveKollegiumSettings(partial) {
    const next = { ...loadKollegiumSettings(), ...partial };
    localStorage.setItem(KEY, JSON.stringify(next));
    return next;
}

/**
 * @returns {{ siteUrl: string, listUmfragen: string, listAntworten: string }}
 */
export function loadKollegiumSpSettings() {
    try {
        const raw = localStorage.getItem(SP_SETTINGS_KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            siteUrl: String((o && o.siteUrl) || '').trim(),
            listUmfragen:
                String((o && o.listUmfragen) || LIST_TITLES.umfragen).trim() || LIST_TITLES.umfragen,
            listAntworten:
                String((o && o.listAntworten) || LIST_TITLES.antworten).trim() || LIST_TITLES.antworten
        };
    } catch {
        return {
            siteUrl: '',
            listUmfragen: LIST_TITLES.umfragen,
            listAntworten: LIST_TITLES.antworten
        };
    }
}

/**
 * @param {Partial<{ siteUrl: string, listUmfragen: string, listAntworten: string }>} partial
 */
export function saveKollegiumSpSettings(partial) {
    const cur = loadKollegiumSpSettings();
    const next = {
        siteUrl: partial.siteUrl != null ? String(partial.siteUrl).trim() : cur.siteUrl,
        listUmfragen:
            partial.listUmfragen != null
                ? String(partial.listUmfragen).trim() || cur.listUmfragen
                : cur.listUmfragen,
        listAntworten:
            partial.listAntworten != null
                ? String(partial.listAntworten).trim() || cur.listAntworten
                : cur.listAntworten
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
export function resolveKollegiumSiteUrl() {
    return resolveSchoolToolSiteUrl('kollegium') || '';
}
