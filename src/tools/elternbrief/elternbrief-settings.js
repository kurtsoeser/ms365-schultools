/**
 * Einstellungen Elternbrief (Absender, Admins) + SPO-Site.
 */
import {
    resolveElternbriefZentralPostfach,
    resolveSchoolToolSiteUrl
} from '../../shared/school-tool-sites.js';
import { LIST_TITLES } from './elternbrief-schema.js';

const KEY = 'ms365-elternbrief-settings-v1';
export const SP_SETTINGS_KEY = 'ms365-elternbrief-sp-settings-v1';

/**
 * @returns {{
 *   schoolName: string,
 *   absenderDefault: 'eigenes_konto'|'zentrales_postfach',
 *   zentralPostfach: string,
 *   zentralName: string,
 *   adminEmails: string,
 *   expDays: number
 * }}
 */
export function loadElternbriefSettings() {
    try {
        const raw = localStorage.getItem(KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            schoolName: String(o.schoolName || '').trim(),
            absenderDefault:
                String(o.absenderDefault || 'eigenes_konto').trim() === 'zentrales_postfach'
                    ? 'zentrales_postfach'
                    : 'eigenes_konto',
            zentralPostfach: resolveElternbriefZentralPostfach(o.zentralPostfach),
            zentralName: String(o.zentralName || 'Elternbrief Schule').trim(),
            adminEmails: String(o.adminEmails || '').trim(),
            expDays: Math.max(1, Number(o.expDays) || 45)
        };
    } catch {
        return {
            schoolName: '',
            absenderDefault: 'eigenes_konto',
            zentralPostfach: '',
            zentralName: 'Elternbrief Schule',
            adminEmails: '',
            expDays: 45
        };
    }
}

/**
 * @param {Partial<ReturnType<typeof loadElternbriefSettings>>} partial
 */
export function saveElternbriefSettings(partial) {
    const next = { ...loadElternbriefSettings(), ...partial };
    localStorage.setItem(KEY, JSON.stringify(next));
    return next;
}

/**
 * @returns {{ siteUrl: string, listBriefe: string, listLese: string }}
 */
export function loadElternbriefSpSettings() {
    try {
        const raw = localStorage.getItem(SP_SETTINGS_KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            siteUrl: String((o && o.siteUrl) || '').trim(),
            listBriefe:
                String((o && o.listBriefe) || LIST_TITLES.briefe).trim() || LIST_TITLES.briefe,
            listLese:
                String((o && o.listLese) || LIST_TITLES.lesebestaetigungen).trim() ||
                LIST_TITLES.lesebestaetigungen
        };
    } catch {
        return {
            siteUrl: '',
            listBriefe: LIST_TITLES.briefe,
            listLese: LIST_TITLES.lesebestaetigungen
        };
    }
}

/**
 * @param {Partial<{ siteUrl: string, listBriefe: string, listLese: string }>} partial
 */
export function saveElternbriefSpSettings(partial) {
    const cur = loadElternbriefSpSettings();
    const next = {
        siteUrl: partial.siteUrl != null ? String(partial.siteUrl).trim() : cur.siteUrl,
        listBriefe:
            partial.listBriefe != null
                ? String(partial.listBriefe).trim() || cur.listBriefe
                : cur.listBriefe,
        listLese:
            partial.listLese != null
                ? String(partial.listLese).trim() || cur.listLese
                : cur.listLese
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
export function resolveElternbriefSiteUrl() {
    return resolveSchoolToolSiteUrl('elternbrief') || '';
}
