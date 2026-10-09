/**
 * State / Rollen: Lehrkraft vs. Direktion.
 */
import { toIsoDateOnly } from './lfr-logic.js';
import { KATEGORIE_CHOICES } from './lfr-schema.js';
import { loadLfrPermissions, direktionEmailFromFreistellungSetup } from './lfr-permissions.js';

const ROLE_KEY = 'ms365-lfr-planer-role-v1';
const SITE_KEY = 'ms365-lfr-planer-site-v1';
const SETUP_KEY = 'ms365-lfr-setup-v1';
const DEMO_KEY = 'ms365-lfr-demo-items-v1';

/** @typedef {'lehrer'|'direktion'} LfrRole */

export const VIEWS = [
    { id: 'dashboard', label: 'Übersicht', icon: 'bi-speedometer2', roles: ['lehrer', 'direktion'] },
    { id: 'meine', label: 'Meine Anträge', icon: 'bi-person', roles: ['lehrer', 'direktion'] },
    { id: 'antrag', label: 'Antrag stellen', icon: 'bi-plus-circle', roles: ['lehrer', 'direktion'] },
    { id: 'liste', label: 'Alle Anträge', icon: 'bi-list-ul', roles: ['direktion'] },
    { id: 'freigabe', label: 'Freigabe', icon: 'bi-check2-square', roles: ['direktion'] },
    { id: 'kalender', label: 'Kalender', icon: 'bi-calendar3', roles: ['lehrer', 'direktion'] },
    { id: 'export', label: 'Kalender-Export', icon: 'bi-download', roles: ['direktion'] },
    { id: 'einrichtung', label: 'Administration', icon: 'bi-gear', roles: ['direktion'] }
];

/**
 * @param {LfrRole|string} role
 */
export function viewsForRole(role) {
    const r = role === 'direktion' ? 'direktion' : 'lehrer';
    return VIEWS.filter((v) => v.roles.includes(r));
}

export function createInitialState() {
    const now = new Date();
    return {
        role: resolveRole(),
        view: 'dashboard',
        siteUrl: loadSavedSiteUrl(),
        listName: loadSetupField('listName') || 'Lehrer-Freistellungen',
        listId: loadSetupField('listId') || '',
        loading: false,
        error: '',
        roleHint: '',
        accountEmail: '',
        accountName: '',
        items: [],
        stammdaten: { teachers: [] },
        filters: { status: '', kategorie: '', q: '' },
        form: emptyForm(),
        editingItemId: null,
        detailId: null,
        calYear: now.getFullYear(),
        calMonth: now.getMonth() + 1,
        ctx: null,
        localDemoOnly: false,
        demoRoleOverride: false,
        calendarExportOnlyApproved: true,
        outlookCalendarUser: loadSetupField('outlookCalendarUser'),
        outlookCalendarId: loadSetupField('outlookCalendarId')
    };
}

export function emptyForm(partial) {
    const t = new Date();
    const iso = toIsoDateOnly(t) || '';
    const local = iso + 'T08:00';
    const end = iso + 'T16:00';
    return {
        antragId: '',
        titel: '',
        beginn: local,
        ende: end,
        kategorie: KATEGORIE_CHOICES[0],
        beschreibung: '',
        lehrerName: '',
        lehrerEmail: '',
        status: 'Ausstehend',
        bemerkungDirektion: '',
        ...(partial || {})
    };
}

/**
 * @param {LfrRole|string} [preferred]
 */
export function resolveRole(preferred) {
    if (preferred === 'direktion' || preferred === 'lehrer') return preferred;
    try {
        const s = String(localStorage.getItem(ROLE_KEY) || '').toLowerCase();
        if (s === 'direktion' || s === 'lehrer') return s;
    } catch {
        /* ignore */
    }
    return 'lehrer';
}

export function persistRole(role) {
    try {
        localStorage.setItem(ROLE_KEY, role === 'direktion' ? 'direktion' : 'lehrer');
    } catch {
        /* ignore */
    }
}

export function loadSetupCfg() {
    try {
        return JSON.parse(localStorage.getItem(SETUP_KEY) || '{}');
    } catch {
        return {};
    }
}

function loadSetupField(key) {
    const c = loadSetupCfg();
    return c && c[key] != null ? String(c[key]).trim() : '';
}

export function persistSetupCfg(patch) {
    const cur = loadSetupCfg();
    const next = Object.assign({}, cur, patch || {});
    try {
        localStorage.setItem(SETUP_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

export function loadSavedSiteUrl() {
    try {
        const local = String(localStorage.getItem(SITE_KEY) || '').trim();
        if (local) return local;
    } catch {
        /* ignore */
    }
    const setup = loadSetupCfg();
    if (setup.siteUrl) return String(setup.siteUrl).trim();
    try {
        const fr = JSON.parse(localStorage.getItem('ms365-freistellung-setup-v1') || '{}');
        if (fr.siteUrl) return String(fr.siteUrl).trim();
    } catch {
        /* ignore */
    }
    try {
        const app =
            window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup()
                : null;
        if (app && app.intranetSiteUrl) return String(app.intranetSiteUrl).trim();
    } catch {
        /* ignore */
    }
    return '';
}

export function persistSiteUrl(url) {
    try {
        localStorage.setItem(SITE_KEY, String(url || '').trim());
    } catch {
        /* ignore */
    }
    persistSetupCfg({ siteUrl: String(url || '').trim() });
}

export function loadStammdaten() {
    const teachers = [];
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getTeachers === 'function') {
            const rows = window.ms365AppDataV2.getTeachers() || [];
            rows.forEach((t) => {
                teachers.push({
                    code: String(t.code || t.id || '').trim(),
                    name: String(t.name || t.displayName || '').trim(),
                    email: String(t.email || t.mail || '').trim().toLowerCase()
                });
            });
        }
    } catch {
        /* ignore */
    }
    return { teachers };
}

/**
 * @param {object[]} teachers
 * @param {string} email
 */
export function matchTeacherByEmail(teachers, email) {
    const em = String(email || '').trim().toLowerCase();
    if (!em) return null;
    return (teachers || []).find((t) => String(t.email || '').trim().toLowerCase() === em) || null;
}

/**
 * @param {object} state
 */
export function scopeFromState(state, opts) {
    const o = opts || {};
    if (state.role === 'direktion' || o.scopeAll) {
        return { scopeAll: true, accountEmail: state.accountEmail };
    }
    return { scopeAll: false, accountEmail: state.accountEmail };
}

export function canDecide(state) {
    return state && state.role === 'direktion';
}

export function loadDemoItems() {
    try {
        const raw = JSON.parse(localStorage.getItem(DEMO_KEY) || '[]');
        return Array.isArray(raw) ? raw : [];
    } catch {
        return [];
    }
}

export function saveDemoItems(items) {
    try {
        localStorage.setItem(DEMO_KEY, JSON.stringify(items || []));
    } catch {
        /* ignore */
    }
}

/**
 * Rollen-Hinweis aus Stammdaten / Direktions-Mail (ohne Graph-Gruppen-Check im MVP).
 * @param {object} state
 */
export function inferRoleHint(state) {
    const perms = loadLfrPermissions();
    const dirMail = perms.emailDirektion || direktionEmailFromFreistellungSetup();
    const em = String(state.accountEmail || '').trim().toLowerCase();
    if (dirMail && em && dirMail === em) return 'direktion';
    const match = matchTeacherByEmail(state.stammdaten.teachers, em);
    if (match) return 'lehrer';
    return '';
}
