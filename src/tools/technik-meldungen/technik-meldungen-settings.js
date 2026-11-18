/**
 * Site-/Listen-Settings Technik-Meldungen (local-first).
 */
import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';

export const SP_SETTINGS_KEY = 'ms365-technik-meldungen-sp-settings-v1';

/**
 * @returns {{ siteUrl: string, listMeldungen: string }}
 */
export function loadTechnikSpSettings() {
    try {
        const raw = localStorage.getItem(SP_SETTINGS_KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            siteUrl: String((o && o.siteUrl) || '').trim(),
            listMeldungen: String((o && o.listMeldungen) || 'TechnikMeldungen').trim() || 'TechnikMeldungen'
        };
    } catch {
        return { siteUrl: '', listMeldungen: 'TechnikMeldungen' };
    }
}

/**
 * @param {Partial<{ siteUrl: string, listMeldungen: string }>} partial
 */
export function saveTechnikSpSettings(partial) {
    const cur = loadTechnikSpSettings();
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
 * Effektive Site-URL: Tool-Settings → Schul-Tool-Sites → Default/Intranet.
 */
export function resolveTechnikSiteUrl() {
    return resolveSchoolToolSiteUrl('technik') || '';
}
