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
    { id: 'system', label: 'Systemadministration', icon: 'bi-gear', section: 'system' }
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
 * Klassen, in denen die E-Mail als Klassenvorstand hinterlegt ist.
 * Dieselbe Stammdaten-Quelle wie sRDP (headEmail), keine eigene Rollenliste.
 * @param {{ classes?: object[] }|null|undefined} stammdaten
 * @param {string} email
 * @returns {string[]}
 */
export function kvKlassenForEmail(stammdaten, email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em || em.indexOf('@') === -1) return [];
    const out = [];
    const seen = new Set();
    for (const c of (stammdaten && stammdaten.classes) || []) {
        const head = String(
            (c && (c.headEmail || c.klassenvorstandEmail || c.kvEmail || c.classTeacherEmail)) || ''
        )
            .trim()
            .toLowerCase();
        if (!head || head !== em) continue;
        const code = String((c && (c.code || c.name)) || '').trim();
        if (!code) continue;
        const key = code.toLowerCase();
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(code);
    }
    return out;
}

/**
 * @param {string} klasse
 * @param {string[]} klassen
 */
function klasseKey(klasse) {
    return String(klasse || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '');
}

export function klasseInList(klasse, klassen) {
    const key = klasseKey(klasse);
    if (!key) return false;
    return (klassen || []).some((c) => klasseKey(c) === key);
}

/**
 * Schreibt „3ak“ und „3 AK“ auf den Klassen-Code aus den Stammdaten.
 * @param {string} klasse
 * @param {string[]} [codes]
 */
export function canonicalKlasse(klasse, codes) {
    const compact = String(klasse || '')
        .trim()
        .replace(/\s+/g, '');
    if (!compact) return '';
    const hit = (codes || []).find((c) => klasseKey(c) === klasseKey(compact));
    return hit ? String(hit).trim() : compact.toUpperCase();
}

/**
 * Letzte nicht stornierte Anmeldung dieser E-Mail.
 * @param {{ anmeldungen?: object[] }|null} state
 * @param {string} email
 */
export function findOwnAnmeldung(state, email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return null;
    const list = (state && state.anmeldungen) || [];
    for (let i = list.length - 1; i >= 0; i--) {
        const row = list[i];
        if (!row || row.status === 'storniert') continue;
        if (String(row.schuelerEmail || '').trim().toLowerCase() === em) return row;
    }
    return null;
}

/**
 * Dieselbe Person aktualisiert ihre Anmeldung, statt eine zweite Zeile anzulegen.
 * @param {object[]} anmeldungen
 * @param {{ id?: string, schuelerEmail?: string }} item
 */
export function matchAnmeldungIndex(anmeldungen, item) {
    const list = Array.isArray(anmeldungen) ? anmeldungen : [];
    const id = String((item && item.id) || '').trim();
    if (id) {
        const byId = list.findIndex((a) => a && a.id === id);
        if (byId >= 0) return byId;
    }
    const email = String((item && item.schuelerEmail) || '')
        .trim()
        .toLowerCase();
    if (!email) return -1;
    for (let i = list.length - 1; i >= 0; i--) {
        const row = list[i];
        if (!row || row.status === 'storniert') continue;
        if (String(row.schuelerEmail || '').trim().toLowerCase() === email) return i;
    }
    return -1;
}

/**
 * Portal-Kachel Wahlfach-Zuteilung: Direktion, Klassenvorstand, Schul-IT.
 * Eine bekannte Lehrkraft ohne diese Rolle sieht den Eintrag nicht.
 * @param {{ isIt?: boolean, demoRole?: string, email?: string, kvKlassen?: string[], localAdmin?: boolean, cachedRole?: string }} [opts]
 */
export function shouldShowWahlfachZuteilungMenu(opts) {
    const o = opts || {};
    if (o.isIt) return true;
    if (o.demoRole === 'lehrer') return false;
    if (o.demoRole === 'direktion' || o.demoRole === 'kv') return true;
    if (o.kvKlassen && o.kvKlassen.length) return true;
    if (o.localAdmin) return true;
    if (o.cachedRole === 'direktion') return true;
    if (!o.email) return true;
    return false;
}

function studentKlasseCode(student) {
    return String((student && (student.klasse || student.klasseCode || student.classCode)) || '').trim();
}

function studentEmail(student) {
    return String((student && (student.email || student.mail || student.upn)) || '')
        .trim()
        .toLowerCase();
}

/**
 * Anmeldungen der KV-Klassen, plus Schüler der Klasse ohne Anmeldung.
 * @param {{ anmeldungen?: object[] }|null} state
 * @param {string[]} klassen
 * @param {object[]} [students]
 * @returns {Array<{ name: string, klasse: string, email: string, anmeldung: object|null }>}
 */
export function classResultRows(state, klassen, students) {
    const codes = Array.isArray(klassen) ? klassen : [];
    /** @type {Array<{ name: string, klasse: string, email: string, anmeldung: object|null }>} */
    const rows = [];
    const seen = new Set();
    for (const a of (state && state.anmeldungen) || []) {
        if (!a || !klasseInList(a.klasse, codes)) continue;
        const email = String(a.schuelerEmail || '')
            .trim()
            .toLowerCase();
        const name = String(a.schuelerName || '').trim();
        const key = (email || name.toLowerCase()) + '|' + String(a.klasse || '').trim().toLowerCase();
        if (seen.has(key)) continue;
        seen.add(key);
        rows.push({
            name: name || email || 'Schüler/in',
            klasse: String(a.klasse || '').trim(),
            email,
            anmeldung: a
        });
    }
    for (const student of students || []) {
        const klasse = studentKlasseCode(student);
        if (!klasseInList(klasse, codes)) continue;
        const email = studentEmail(student);
        const name = String((student && (student.name || student.displayName)) || '').trim();
        const key = (email || name.toLowerCase()) + '|' + klasse.toLowerCase();
        if (seen.has(key)) continue;
        const already = rows.some((row) => {
            const sameClass = String(row.klasse || '').trim().toLowerCase() === klasse.toLowerCase();
            if (!sameClass) return false;
            if (email && row.email === email) return true;
            return name && row.name.toLowerCase() === name.toLowerCase();
        });
        if (already) continue;
        seen.add(key);
        rows.push({
            name: name || email || 'Schüler/in',
            klasse,
            email,
            anmeldung: null
        });
    }
    rows.sort((a, b) => {
        const byClass = a.klasse.localeCompare(b.klasse, 'de', { numeric: true });
        if (byClass) return byClass;
        return a.name.localeCompare(b.name, 'de');
    });
    return rows;
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
    const sections = [{ id: 'arbeit', label: 'Arbeit', items: arbeit }];
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
        systemTab: 'site',
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
        zuteilKlasse: '',
        zuteilNurOffen: false,
        access: /** @type {{ ok: boolean, reason: string, demo: boolean }|null} */ (null),
        spoSyncedAt: ''
    };
}
