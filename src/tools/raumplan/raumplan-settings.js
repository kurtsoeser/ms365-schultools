/**
 * Site-/Listen-Settings Raumplan (local-first, Familie betrieb).
 */
import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';

export const SP_SETTINGS_KEY = 'ms365-raumplan-sp-settings-v1';

/**
 * @returns {{ siteUrl: string, listRaeume: string }}
 */
export function loadRaumplanSpSettings() {
    try {
        const raw = localStorage.getItem(SP_SETTINGS_KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            siteUrl: String((o && o.siteUrl) || '').trim(),
            listRaeume: String((o && o.listRaeume) || 'Raeume').trim() || 'Raeume'
        };
    } catch {
        return { siteUrl: '', listRaeume: 'Raeume' };
    }
}

/**
 * @param {Partial<{ siteUrl: string, listRaeume: string }>} partial
 */
export function saveRaumplanSpSettings(partial) {
    const cur = loadRaumplanSpSettings();
    const next = {
        siteUrl: partial.siteUrl != null ? String(partial.siteUrl).trim() : cur.siteUrl,
        listRaeume:
            partial.listRaeume != null
                ? String(partial.listRaeume).trim() || cur.listRaeume
                : cur.listRaeume
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
export function resolveRaumplanSiteUrl() {
    return resolveSchoolToolSiteUrl('raumplan') || '';
}
