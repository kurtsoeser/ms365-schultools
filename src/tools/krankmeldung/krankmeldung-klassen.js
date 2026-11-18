/**
 * Klassenliste für Krankmeldung (Admin + öffentliches Formular).
 * Eltern-Browser haben keine Stammdaten → Liste aus Einstellungen und/oder ?klassen=.
 */

import { loadEinsammlerStammdaten } from '../einsammler/einsammler-stammdaten.js';
import { loadKrankmeldungSettings } from './krankmeldung-settings.js';

/**
 * @param {string} text
 * @returns {string[]}
 */
export function parseKlassenList(text) {
    const seen = new Set();
    /** @type {string[]} */
    const out = [];
    for (const part of String(text || '').split(/[;,\n]+/)) {
        const code = part.trim();
        if (!code) continue;
        const key = code.toLowerCase();
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(code);
    }
    return out;
}

/**
 * Klassen aus Tenant-Stammdaten (nur im Schul-/Admin-Browser).
 * @returns {Array<{ code: string, name: string }>}
 */
export function loadStammdatenKlassen() {
    try {
        return loadEinsammlerStammdaten().classes || [];
    } catch {
        return [];
    }
}

/**
 * Veröffentlichten Klassen-Text aus Stammdaten erzeugen.
 */
export function stammdatenKlassenAsText() {
    return loadStammdatenKlassen()
        .map((c) => c.code)
        .join(';');
}

/**
 * Klassen für das öffentliche Formular (Priorität: URL → Einstellungen → Stammdaten).
 * @param {string} [search]
 * @returns {string[]}
 */
export function resolvePublicKlassen(search) {
    const q = new URLSearchParams(
        String(search != null ? search : typeof location !== 'undefined' ? location.search : '')
    );
    const fromParam = parseKlassenList(q.get('klassen') || q.get('klassenCodes') || '');
    if (fromParam.length) return fromParam;
    const fromSettings = parseKlassenList(loadKrankmeldungSettings().publicKlassen || '');
    if (fromSettings.length) return fromSettings;
    return loadStammdatenKlassen().map((c) => c.code);
}

/**
 * Öffentliche Formular-URL inkl. Klassen (für Teilen an Eltern).
 * @param {string} [baseHref]
 * @param {string} [klassenText]
 */
export function buildPublicFormUrl(baseHref, klassenText) {
    const base =
        String(baseHref || 'krankmeldung-formular.html').trim() || 'krankmeldung-formular.html';
    const list = parseKlassenList(klassenText || loadKrankmeldungSettings().publicKlassen || '');
    if (!list.length) return base;
    const join = base.includes('?') ? '&' : '?';
    return base + join + 'klassen=' + encodeURIComponent(list.join(';'));
}

/**
 * Klassen, bei denen die E-Mail als Klassenvorstand hinterlegt ist.
 * @param {string} [email]
 * @returns {string[]}
 */
export function resolveKvKlassenForEmail(email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return [];
    /** @type {string[]} */
    const out = [];
    const seen = new Set();
    const push = (code) => {
        const c = String(code || '').trim();
        if (!c) return;
        const key = c.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(c);
    };
    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const core = window.ms365TenantSettingsLoad();
            const data = (core && core.data) || core || {};
            for (const c of Array.isArray(data.classes) ? data.classes : []) {
                const head = String(
                    (c && (c.headEmail || c.klassenvorstandEmail || c.kvEmail)) || ''
                )
                    .trim()
                    .toLowerCase();
                if (head === em) push(c.code || c.name);
            }
        }
    } catch {
        /* ignore */
    }
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365StammdatenCanonical &&
            typeof window.ms365StammdatenCanonical.listAllClasses === 'function'
        ) {
            for (const c of window.ms365StammdatenCanonical.listAllClasses() || []) {
                const head = String((c && (c.headEmail || c.klassenvorstandEmail)) || '')
                    .trim()
                    .toLowerCase();
                if (head === em) push(c.code || c.name);
            }
        }
    } catch {
        /* ignore */
    }
    return out;
}

/**
 * Aktuelle Microsoft-Anmeldung (falls vorhanden).
 * @returns {{ email: string, name: string, loggedIn: boolean }}
 */
export function getKrankmeldungAccount() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            if (info) {
                const email = String(info.username || info.upn || '')
                    .trim()
                    .toLowerCase();
                return { email, name: String(info.name || '').trim(), loggedIn: !!email };
            }
        }
    } catch {
        /* ignore */
    }
    return { email: '', name: '', loggedIn: false };
}

/**
 * Klassen für Inbox-Filter / Admin-Auswahl: Stammdaten ∪ veröffentlichte ∪ bereits gemeldete.
 * @param {object[]} [meldungen]
 * @returns {string[]}
 */
export function resolveAdminKlassen(meldungen) {
    const seen = new Set();
    /** @type {string[]} */
    const out = [];
    const push = (code) => {
        const c = String(code || '').trim();
        if (!c) return;
        const key = c.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(c);
    };
    for (const c of loadStammdatenKlassen()) push(c.code);
    for (const c of parseKlassenList(loadKrankmeldungSettings().publicKlassen || '')) push(c);
    for (const m of meldungen || []) push(m && m.klasseCode);
    out.sort((a, b) => a.localeCompare(b, 'de'));
    return out;
}
