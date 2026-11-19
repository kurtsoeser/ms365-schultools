/**
 * Einsammler-Einstellungen (Secret/API optional) + SharePoint-Site.
 */

import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';
import { LIST_TITLES } from './einsammler-schema.js';

const KEY = 'ms365-einsammler-settings-v1';
export const SP_SETTINGS_KEY = 'ms365-einsammler-sp-settings-v1';

/**
 * @returns {{
 *   schoolName: string,
 *   apiBaseUrl: string,
 *   adminKey: string,
 *   linkSecret: string,
 *   expDays: number,
 *   adminEmails: string
 * }}
 */
export function loadEinsammlerSettings() {
    try {
        const raw = localStorage.getItem(KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            schoolName: String(o.schoolName || '').trim(),
            // Remote-API bewusst leer halten (w2) – nur manuell / IT
            apiBaseUrl: String(o.apiBaseUrl || '').trim().replace(/\/+$/, ''),
            adminKey: String(o.adminKey || '').trim(),
            linkSecret: String(o.linkSecret || '').trim(),
            expDays: Math.max(1, Number(o.expDays) || 45),
            adminEmails: String(o.adminEmails || '').trim()
        };
    } catch {
        return {
            schoolName: '',
            apiBaseUrl: '',
            adminKey: '',
            linkSecret: '',
            expDays: 45,
            adminEmails: ''
        };
    }
}

/**
 * @param {Partial<ReturnType<typeof loadEinsammlerSettings>>} partial
 */
export function saveEinsammlerSettings(partial) {
    const next = { ...loadEinsammlerSettings(), ...partial };
    localStorage.setItem(KEY, JSON.stringify(next));
    return next;
}

/**
 * Remote nur, wenn Basis-URL und Link-Geheimnis gesetzt sind (a5).
 * Demo/Default: beides leer → API gilt als abgeschaltet (w2).
 */
export function remoteConfigured() {
    const s = loadEinsammlerSettings();
    return !!(s.apiBaseUrl && s.linkSecret);
}

/**
 * @returns {{ siteUrl: string, listKoepfe: string, listAntworten: string }}
 */
export function loadEinsammlerSpSettings() {
    try {
        const raw = localStorage.getItem(SP_SETTINGS_KEY);
        const o = raw ? JSON.parse(raw) : {};
        return {
            siteUrl: String((o && o.siteUrl) || '').trim(),
            listKoepfe:
                String((o && o.listKoepfe) || LIST_TITLES.koepfe).trim() || LIST_TITLES.koepfe,
            listAntworten:
                String((o && o.listAntworten) || LIST_TITLES.antworten).trim() ||
                LIST_TITLES.antworten
        };
    } catch {
        return {
            siteUrl: '',
            listKoepfe: LIST_TITLES.koepfe,
            listAntworten: LIST_TITLES.antworten
        };
    }
}

/**
 * @param {Partial<{ siteUrl: string, listKoepfe: string, listAntworten: string }>} partial
 */
export function saveEinsammlerSpSettings(partial) {
    const cur = loadEinsammlerSpSettings();
    const next = {
        siteUrl: partial.siteUrl != null ? String(partial.siteUrl).trim() : cur.siteUrl,
        listKoepfe:
            partial.listKoepfe != null
                ? String(partial.listKoepfe).trim() || cur.listKoepfe
                : cur.listKoepfe,
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
 * Effektive Site-URL: Tool-SP-Settings → Schul-Tool-Sites → Intranet.
 */
export function resolveEinsammlerSiteUrl() {
    return resolveSchoolToolSiteUrl('einsammler') || '';
}
