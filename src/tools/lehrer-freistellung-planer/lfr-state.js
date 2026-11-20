/**
 * State / Rollen: Lehrkraft vs. Direktion.
 */
import { filterItems, toIsoDateOnly } from './lfr-logic.js';
import { KATEGORIE_CHOICES } from './lfr-schema.js';
import {
    loadLfrPermissions,
    accountIsLfrDirektionAccess,
    accountIsLfrLehrerAccess
} from './lfr-permissions.js';
import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';
import { loadEffectiveTeachersForPersona } from '../../shared/live-school-stammdaten-cache.js';

const ROLE_KEY = 'ms365-lfr-planer-role-v1';
const SITE_KEY = 'ms365-lfr-planer-site-v1';
const SETUP_KEY = 'ms365-lfr-setup-v1';
const DEMO_KEY = 'ms365-lfr-demo-items-v1';
/** @type {import('../../shared/school-tool-sites.js').SchoolToolSiteKind} */
export const LFR_SITE_KIND = 'lehrerFreistellung';

/** @typedef {'lehrer'|'direktion'} LfrRole */

export const VIEWS = [
    { id: 'dashboard', label: 'Übersicht', icon: 'bi-speedometer2', roles: ['lehrer', 'direktion'] },
    { id: 'meine', label: 'Meine Anträge', icon: 'bi-person', roles: ['lehrer', 'direktion'] },
    { id: 'antrag', label: 'Antrag stellen', icon: 'bi-plus-circle', roles: ['lehrer', 'direktion'] },
    { id: 'liste', label: 'Alle Anträge', icon: 'bi-list-ul', roles: ['direktion'] },
    { id: 'freigabe', label: 'Entscheidungen', icon: 'bi-check2-square', roles: ['direktion'] },
    { id: 'kalender', label: 'Kalender', icon: 'bi-calendar3', roles: ['lehrer', 'direktion'] },
    { id: 'export', label: 'Kalender-Export', icon: 'bi-download', roles: ['direktion'] },
    { id: 'einrichtung', label: 'Systemadministration', icon: 'bi-gear', roles: ['direktion'] }
];

function viewById(id) {
    return VIEWS.find((v) => v.id === id) || null;
}

/**
 * Gruppierte Navigation (approval-planer wie Schulaktivitäten / Schüler-FR).
 * Lehrkraft: flach. Direktion: Arbeit · Auswertung · System.
 * @param {LfrRole|string} role
 * @returns {{ id: string, label: string, items: typeof VIEWS }[]}
 */
export function navSectionsForRole(role) {
    const r = role === 'direktion' ? 'direktion' : 'lehrer';
    if (r === 'direktion') {
        return [
            {
                id: 'arbeit',
                label: 'Arbeit',
                items: [
                    viewById('dashboard'),
                    viewById('freigabe'),
                    viewById('liste'),
                    viewById('kalender'),
                    viewById('antrag')
                ].filter(Boolean)
            },
            {
                id: 'auswertung',
                label: 'Auswertung',
                items: [viewById('export')].filter(Boolean)
            },
            {
                id: 'system',
                label: 'System',
                items: [viewById('einrichtung')].filter(Boolean)
            }
        ].filter((s) => (s.items || []).length);
    }
    return [
        {
            id: 'arbeit',
            label: '',
            items: [
                viewById('dashboard'),
                viewById('meine'),
                viewById('kalender'),
                viewById('antrag')
            ].filter(Boolean)
        }
    ];
}

/**
 * @param {LfrRole|string} role
 */
export function viewsForRole(role) {
    return navSectionsForRole(role).flatMap((s) => s.items || []);
}

/**
 * Offene Direktions-Entscheidungen.
 * @param {object[]} items
 */
export function countOpenAusstehend(items) {
    return (Array.isArray(items) ? items : []).filter((i) => {
        const s = String(i && i.status || '')
            .trim()
            .toLowerCase();
        return s === 'ausstehend';
    }).length;
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
        calMode: 'month',
        calWeekAnchor: now.toISOString(),
        ctx: null,
        localDemoOnly: false,
        demoRoleOverride: false,
        entraGroupsConfigured: false,
        planerRoles: /** @type {import('./lfr-entra-role.js').LfrRole[]} */ ([]),
        planerRoleSources: /** @type {Record<string, string>} */ ({}),
        planerAccessDenied: false,
        calendarExportOnlyApproved: true,
        outlookCalendarUser: loadSetupField('outlookCalendarUser'),
        outlookCalendarId: loadSetupField('outlookCalendarId'),
        systemTab: 'site'
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
 * Lehrkraft-Identität aus dem angemeldeten Konto (und Stammdaten-Match) ins Formular.
 * Felder bleiben readonly – Antragsteller ist immer der eingeloggte User.
 * @param {object} state
 * @param {{ force?: boolean }} [opts]
 */
export function prefillLehrerFormIdentity(state, opts) {
    if (!state || !state.form) return state;
    const force = !!(opts && opts.force);
    const match = matchTeacherByEmail(
        (state.stammdaten && state.stammdaten.teachers) || [],
        state.accountEmail
    );
    const email = String(state.accountEmail || '').trim().toLowerCase();
    const name = String(
        state.accountName || (match && match.name) || state.form.lehrerName || ''
    ).trim();
    if (force || !String(state.form.lehrerEmail || '').trim()) {
        if (email) state.form.lehrerEmail = email;
    }
    if (force || !String(state.form.lehrerName || '').trim()) {
        if (name) state.form.lehrerName = name;
    }
    return state;
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
        // Stammdaten planerSiteUrls / Default / Intranet (ohne erneuten Local-Key)
        const via = resolveSchoolToolSiteUrl(LFR_SITE_KIND, undefined, { ignoreLocal: true });
        if (via) return via;
    } catch {
        /* ignore */
    }
    // Migration: dieselbe Site wie Schüler-Freistellung / Setup-JSON
    try {
        const fr = resolveSchoolToolSiteUrl('freistellung');
        if (fr) return fr;
    } catch {
        /* ignore */
    }
    try {
        const fr = JSON.parse(localStorage.getItem('ms365-freistellung-setup-v1') || '{}');
        if (fr.siteUrl) return String(fr.siteUrl).trim();
    } catch {
        /* ignore */
    }
    return '';
}

export function persistSiteUrl(url) {
    const next = String(url || '').trim();
    try {
        localStorage.setItem(SITE_KEY, next);
    } catch {
        /* ignore */
    }
    persistSetupCfg({ siteUrl: next });
}

export function loadStammdaten() {
    /** @type {{ code: string, name: string, email: string }[]} */
    const teachers = [];
    const seen = new Set();
    const pushRows = (rows) => {
        (rows || []).forEach((t) => {
            const email = String(t.email || t.mail || '')
                .trim()
                .toLowerCase();
            const code = String(t.code || t.id || '').trim();
            const name = String(t.name || t.displayName || '').trim();
            const key = (email || code).toLowerCase();
            if (!key || seen.has(key)) return;
            seen.add(key);
            teachers.push({ code, name, email });
        });
    };
    try {
        pushRows(loadEffectiveTeachersForPersona());
    } catch {
        /* ignore */
    }
    if (!teachers.length) {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getTeachers === 'function') {
                pushRows(window.ms365AppDataV2.getTeachers() || []);
            }
        } catch {
            /* ignore */
        }
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
 * Sichtbarkeits-Scope für Listen/Kalender.
 * Explizites `scopeAll: false` gewinnt immer (auch bei Rolle Direktion → „Meine Anträge“).
 * Lehrkraft ohne Option: immer nur eigene Anträge.
 * @param {object} state
 * @param {{ scopeAll?: boolean }} [opts]
 */
export function scopeFromState(state, opts) {
    const o = opts || {};
    const email = state && state.accountEmail;
    const accountName = state && state.accountName;
    if (Object.prototype.hasOwnProperty.call(o, 'scopeAll')) {
        return { scopeAll: !!o.scopeAll, accountEmail: email, accountName };
    }
    if (state && state.role === 'direktion') {
        return { scopeAll: true, accountEmail: email, accountName };
    }
    return { scopeAll: false, accountEmail: email, accountName };
}

/**
 * Anträge für die aktuelle Rolle/Ansicht (Lehrkraft: nur eigene).
 * @param {object} state
 * @param {object} [filters]
 * @param {{ scopeAll?: boolean }} [scopeOpts]
 */
export function visibleItemsForState(state, filters, scopeOpts) {
    return filterItems(state && state.items, filters || {}, scopeFromState(state, scopeOpts));
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

export function clearDemoItems() {
    try {
        localStorage.removeItem(DEMO_KEY);
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
    const em = String(state.accountEmail || '').trim().toLowerCase();
    // Nur LFR-Perms/Setup – Schüler-FR-Genehmiger-Mail darf hier keine Direktion setzen
    if (accountIsLfrDirektionAccess(em, perms)) {
        return 'direktion';
    }
    if (accountIsLfrLehrerAccess(em, perms)) return 'lehrer';
    const match = matchTeacherByEmail(state.stammdaten.teachers, em);
    if (match) return 'lehrer';
    return '';
}
