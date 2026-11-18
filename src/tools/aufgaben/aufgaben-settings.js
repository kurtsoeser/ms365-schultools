/**
 * Site-/Listen-Settings Aufgaben (local-first, Familie betrieb).
 */
import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';

export const SP_SETTINGS_KEY = 'ms365-aufgaben-sp-settings-v1';

/**
 * @returns {{ siteUrl: string, listAufgaben: string }}
 */
export function loadAufgabenSpSettings() {
    try {
        const raw = localStorage.getItem(SP_SETTINGS_KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            siteUrl: String((o && o.siteUrl) || '').trim(),
            listAufgaben: String((o && o.listAufgaben) || 'Aufgaben').trim() || 'Aufgaben'
        };
    } catch {
        return { siteUrl: '', listAufgaben: 'Aufgaben' };
    }
}

/**
 * @param {Partial<{ siteUrl: string, listAufgaben: string }>} partial
 */
export function saveAufgabenSpSettings(partial) {
    const cur = loadAufgabenSpSettings();
    const next = {
        siteUrl: partial.siteUrl != null ? String(partial.siteUrl).trim() : cur.siteUrl,
        listAufgaben:
            partial.listAufgaben != null
                ? String(partial.listAufgaben).trim() || cur.listAufgaben
                : cur.listAufgaben
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
export function resolveAufgabenSiteUrl() {
    return resolveSchoolToolSiteUrl('aufgaben') || '';
}
