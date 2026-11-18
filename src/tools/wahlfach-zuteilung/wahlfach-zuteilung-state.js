/**
 * UI-State / Nav / Stammdaten-Helfer Wahlfach-Zuteilung (wahl-portal).
 */

import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';
import { shouldShowItAppChromeSync } from '../../shared/frontend-planner-chrome-policy.js';
import { loadSettings } from './wahlfach-zuteilung-settings.js';
import { resolveWzPhase } from './wahlfach-zuteilung-logic.js';

/** @typedef {'uebersicht'|'angebote'|'anmeldungen'|'zuteilung'|'ergebnis'|'system'} WzView */

export const VIEWS = Object.freeze([
    { id: 'uebersicht', label: 'Übersicht', icon: 'bi-speedometer2', section: 'arbeit' },
    { id: 'angebote', label: 'Angebote', icon: 'bi-collection', section: 'arbeit' },
    { id: 'anmeldungen', label: 'Anmeldungen', icon: 'bi-people', section: 'arbeit' },
    { id: 'zuteilung', label: 'Zuteilung', icon: 'bi-diagram-3', section: 'arbeit' },
    { id: 'ergebnis', label: 'Ergebnis', icon: 'bi-table', section: 'arbeit' },
    { id: 'system', label: 'System', icon: 'bi-gear', section: 'system' }
]);

/**
 * @returns {{ classes: object[], teachers: object[], subjects: object[], students: object[] }}
 */
export function loadWzStammdaten() {
    try {
        const core = window.ms365TenantSettingsLoad && window.ms365TenantSettingsLoad();
        const data = (core && core.data) || core || {};
        return {
            classes: Array.isArray(data.classes) ? data.classes : [],
            teachers: Array.isArray(data.teachers) ? data.teachers : [],
            subjects: Array.isArray(data.subjects) ? data.subjects : [],
            students: Array.isArray(data.students) ? data.students : []
        };
    } catch {
        return { classes: [], teachers: [], subjects: [], students: [] };
    }
}

/**
 * Klassen-Codes für Dropdown (Stammdaten local-first).
 * @param {{ classes?: object[] }|null} stammdaten
 * @returns {string[]}
 */
export function classCodesFromStammdaten(stammdaten) {
    const out = [];
    const seen = new Set();
    for (const c of (stammdaten && stammdaten.classes) || []) {
        const code = String((c && (c.code || c.klasseCode || c.name || c.id)) || '').trim();
        if (!code) continue;
        const key = code.toLowerCase();
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(code);
    }
    return out.sort((a, b) => a.localeCompare(b, 'de'));
}

/**
 * @returns {{ email: string, name: string }}
 */
export function readWzAccount() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const a = window.ms365AuthGetAccountInfo();
            return {
                email: String((a && (a.username || a.upn || a.mail)) || '')
                    .trim()
                    .toLowerCase(),
                name: String((a && (a.name || a.displayName)) || '').trim()
            };
        }
    } catch {
        /* ignore */
    }
    return { email: '', name: '' };
}

/**
 * Site-URL: Settings-Override → school-tool-sites kind `wahlfach`.
 * @returns {string}
 */
export function resolveWzSiteUrl() {
    try {
        const cfg = loadSettings();
        const fromSettings = String((cfg && cfg.siteUrl) || '').trim();
        if (fromSettings) return fromSettings;
    } catch {
        /* ignore */
    }
    try {
        return String(resolveSchoolToolSiteUrl('wahlfach') || '').trim();
    } catch {
        return '';
    }
}

/**
 * IT-/System-Chrome nur bei bewusster IT-Shell.
 * @returns {boolean}
 */
export function showWzSystemChrome() {
    try {
        return shouldShowItAppChromeSync();
    } catch {
        return false;
    }
}

/**
 * Nav-Sektionen: Arbeit immer; System nur wenn IT-Chrome.
 * @param {{ showSystem?: boolean }} [opts]
 */
export function navSectionsForWz(opts) {
    const showSystem = !!(opts && opts.showSystem);
    const arbeit = VIEWS.filter((v) => v.section === 'arbeit');
    const sections = [{ id: 'arbeit', label: '', items: arbeit }];
    if (showSystem) {
        sections.push({
            id: 'system',
            label: 'System',
            items: VIEWS.filter((v) => v.section === 'system')
        });
    }
    return sections;
}

/**
 * @param {string} view
 * @param {{ showSystem?: boolean }} [opts]
 */
export function ensureWzView(view, opts) {
    const id = String(view || 'uebersicht');
    const allowed = navSectionsForWz(opts).flatMap((s) => s.items || []);
    if (allowed.some((v) => v.id === id)) return id;
    return 'uebersicht';
}

/**
 * Demo-Opt-in über URL.
 * @returns {boolean}
 */
export function wantsWzDemoFromUrl() {
    try {
        const q = new URLSearchParams(window.location.search || '');
        return q.get('demo') === '1';
    } catch {
        return false;
    }
}

/**
 * @param {object|null} runde
 * @param {Date|string} [now]
 */
export function phaseForRunde(runde, now) {
    return resolveWzPhase(runde, now);
}

/**
 * Frischer UI-Runtime-State (neben dem Persistenz-Store).
 */
export function createWzUiState() {
    const account = readWzAccount();
    return {
        view: /** @type {WzView} */ ('uebersicht'),
        message: /** @type {{ type: string, text: string }|null} */ (null),
        softError: '',
        hardError: '',
        loading: false,
        localDemoOnly: false,
        siteUrl: resolveWzSiteUrl(),
        siteId: '',
        stammdaten: loadWzStammdaten(),
        stammdatenLive: {
            ok: false,
            error: '',
            classCount: 0,
            teacherCount: 0,
            source: 'cache'
        },
        accountEmail: account.email,
        accountName: account.name,
        access: /** @type {{ ok: boolean, reason: string, demo: boolean }|null} */ (null),
        spoSyncedAt: ''
    };
}
