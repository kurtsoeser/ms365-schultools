/**
 * State-Helfer & Filter für den Schularbeiten-Planer.
 */
import { DEFAULT_RULES } from './schularbeiten-planer-logic.js';

export const ROLE_STORAGE_KEY = 'ms365-sa-demo-role';
export const SITE_STORAGE_KEY = 'ms365-sa-site-url';
export const SETTINGS_STORAGE_KEY = 'ms365-sa-settings-v1';
export const DEMO_KLASSE_STORAGE_KEY = 'ms365-sa-demo-klasse';

/** @typedef {'lehrer'|'admin'|'schueler'} PlanerRole */

export const VIEWS = [
    { id: 'dashboard', label: 'Dashboard', icon: 'bi-speedometer2', adminOnly: false, staffOnly: false },
    { id: 'kalender', label: 'Kalender', icon: 'bi-calendar3', adminOnly: false, staffOnly: false },
    { id: 'neu', label: 'Neue Schularbeit', icon: 'bi-plus-circle', adminOnly: false, staffOnly: true },
    { id: 'meine', label: 'Meine Schularbeiten', icon: 'bi-card-checklist', adminOnly: false, staffOnly: true },
    { id: 'admin', label: 'Administration', icon: 'bi-shield-check', adminOnly: true, staffOnly: true },
    { id: 'regeln', label: 'Regelwerk', icon: 'bi-book', adminOnly: false, staffOnly: true },
    { id: 'export', label: 'Export', icon: 'bi-download', adminOnly: false, staffOnly: false }
];

/**
 * @param {PlanerRole|string} role
 */
export function viewsForRole(role) {
    const r = String(role || 'lehrer');
    const isAdmin = r === 'admin';
    const isSchueler = r === 'schueler';
    return VIEWS.filter((v) => {
        if (v.adminOnly && !isAdmin) return false;
        if (v.staffOnly && isSchueler) return false;
        return true;
    });
}

export function createInitialState() {
    const now = new Date();
    return {
        siteUrl: '',
        ctx: null,
        view: 'dashboard',
        role: 'lehrer',
        filters: { klasse: '', fach: '', lehrer: '', status: '' },
        stammdaten: { subjects: [], classes: [], teachers: [], students: [] },
        items: [],
        windows: [],
        fachMeta: [],
        rules: { ...DEFAULT_RULES, name: 'Standard', itemId: '', regelwerkId: 'rw-1', aktiv: true },
        settings: loadPlanerSettings(),
        loading: false,
        bootstrapped: false,
        error: '',
        accountEmail: '',
        accountName: '',
        teacherMatch: null,
        studentMatch: null,
        /** Demo-Override, wenn Anmeldung keinem Schüler zugeordnet ist */
        demoKlasseCode: loadDemoKlasseCode(),
        calYear: now.getFullYear(),
        calMonth: now.getMonth(),
        form: emptyForm(),
        editingItemId: null,
        detailId: null
    };
}

export function emptyForm(overrides) {
    return {
        thema: '',
        fachCode: '',
        klasseCode: '',
        lehrerCode: '',
        lehrerEmail: '',
        datum: '',
        dauerMinuten: 100,
        semester: 'WS',
        notiz: '',
        schularbeitId: '',
        ...(overrides || {})
    };
}

/**
 * Stammdaten aus tenant-settings laden.
 */
export function loadStammdaten() {
    let settings = null;
    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            settings = window.ms365TenantSettingsLoad();
        }
    } catch {
        settings = null;
    }
    const subjects = Array.isArray(settings && settings.subjects) ? settings.subjects : [];
    const classes = Array.isArray(settings && settings.classes) ? settings.classes : [];
    const teachers = Array.isArray(settings && settings.teachers) ? settings.teachers : [];
    const students = Array.isArray(settings && settings.students) ? settings.students : [];
    return {
        subjects: subjects
            .map((s) => ({
                code: String(s.code || '').trim(),
                name: String(s.name || s.code || '').trim()
            }))
            .filter((s) => s.code),
        classes: classes
            .map((c) => ({
                code: String(c.code || '').trim(),
                name: String(c.name || c.code || '').trim(),
                year: c.year
            }))
            .filter((c) => c.code),
        teachers: teachers
            .map((t) => ({
                code: String(t.code || '').trim(),
                name: String(t.name || t.code || '').trim(),
                email: String(t.email || '').trim().toLowerCase()
            }))
            .filter((t) => t.code),
        students: students
            .map((s) => ({
                klasse: String(s.klasse || s.class || '').trim(),
                name: String(s.name || '').trim(),
                email: String(s.email || '').trim().toLowerCase()
            }))
            .filter((s) => s.klasse || s.email)
    };
}

/**
 * @param {string} email
 * @param {{ code: string, name: string, email: string }[]} teachers
 */
export function matchTeacherByEmail(email, teachers) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return null;
    const list = Array.isArray(teachers) ? teachers : [];
    return list.find((t) => t.email && t.email === em) || null;
}

/**
 * @param {string} email
 * @param {{ klasse: string, name: string, email: string }[]} students
 */
export function matchStudentByEmail(email, students) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return null;
    const list = Array.isArray(students) ? students : [];
    return list.find((s) => s.email && s.email === em) || null;
}

/**
 * Effektive Klassen-Codes für Schüler-Ansicht (Stammdaten-Match oder Demo-Override).
 * @param {{ studentMatch?: { klasse: string }|null, demoKlasseCode?: string }} stateOrScope
 */
export function resolveStudentKlasseCode(stateOrScope) {
    const fromMatch =
        stateOrScope && stateOrScope.studentMatch && stateOrScope.studentMatch.klasse
            ? String(stateOrScope.studentMatch.klasse).trim()
            : '';
    if (fromMatch) return fromMatch;
    const demo = stateOrScope && stateOrScope.demoKlasseCode ? String(stateOrScope.demoKlasseCode).trim() : '';
    return demo || '';
}

export function loadDemoKlasseCode() {
    try {
        return String(localStorage.getItem(DEMO_KLASSE_STORAGE_KEY) || '').trim();
    } catch {
        return '';
    }
}

export function persistDemoKlasseCode(code) {
    try {
        const c = String(code || '').trim();
        if (c) localStorage.setItem(DEMO_KLASSE_STORAGE_KEY, c);
        else localStorage.removeItem(DEMO_KLASSE_STORAGE_KEY);
    } catch {
        /* ignore */
    }
}

/**
 * Scope-Objekt für Filter/Rechte aus App-State.
 * @param {object} state
 * @param {{ onlyMine?: boolean, scopeAll?: boolean }} [opts]
 */
export function scopeFromState(state, opts) {
    const o = opts || {};
    const role = o.scopeAll ? 'admin' : (state && state.role) || 'lehrer';
    return {
        role,
        teacherMatch: (state && state.teacherMatch) || null,
        studentMatch: (state && state.studentMatch) || null,
        demoKlasseCode: (state && state.demoKlasseCode) || '',
        accountEmail: (state && state.accountEmail) || '',
        onlyMine: !!o.onlyMine
    };
}

/**
 * Demo-Rolle aus localStorage; produktiv später Entra-Gruppen.
 * @param {PlanerRole|null|undefined} preferred
 */
export function resolveRole(preferred) {
    if (preferred === 'admin' || preferred === 'lehrer' || preferred === 'schueler') return preferred;
    try {
        const stored = String(localStorage.getItem(ROLE_STORAGE_KEY) || '').toLowerCase();
        if (stored === 'admin' || stored === 'lehrer' || stored === 'schueler') return stored;
    } catch {
        /* ignore */
    }
    return 'lehrer';
}

/**
 * @param {PlanerRole|string} role
 */
export function persistRole(role) {
    try {
        const r = role === 'admin' ? 'admin' : role === 'schueler' ? 'schueler' : 'lehrer';
        localStorage.setItem(ROLE_STORAGE_KEY, r);
    } catch {
        /* ignore */
    }
}

export function loadSavedSiteUrl() {
    try {
        const local = String(localStorage.getItem(SITE_STORAGE_KEY) || '').trim();
        if (local) return local;
    } catch {
        /* ignore */
    }
    try {
        const setup =
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup()
                : null;
        if (setup && setup.intranetSiteUrl) return String(setup.intranetSiteUrl).trim();
    } catch {
        /* ignore */
    }
    return '';
}

export function persistSiteUrl(url) {
    try {
        localStorage.setItem(SITE_STORAGE_KEY, String(url || '').trim());
    } catch {
        /* ignore */
    }
}

/**
 * @returns {{ syncSchultermine: boolean, schultermineList: string }}
 */
export function loadPlanerSettings() {
    const defaults = { syncSchultermine: false, schultermineList: 'Schultermine' };
    try {
        const raw = JSON.parse(localStorage.getItem(SETTINGS_STORAGE_KEY) || '{}') || {};
        return {
            syncSchultermine: !!raw.syncSchultermine,
            schultermineList: String(raw.schultermineList || defaults.schultermineList).trim() || defaults.schultermineList
        };
    } catch {
        return { ...defaults };
    }
}

/**
 * @param {{ syncSchultermine?: boolean, schultermineList?: string }} patch
 */
export function persistPlanerSettings(patch) {
    const next = { ...loadPlanerSettings(), ...(patch || {}) };
    try {
        localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

/**
 * @param {object[]} fachMeta
 * @param {string} fachCode
 */
export function fachMetaForCode(fachMeta, fachCode) {
    const code = String(fachCode || '').trim();
    if (!code) return null;
    const list = Array.isArray(fachMeta) ? fachMeta : [];
    return list.find((m) => String(m.fachCode || '').trim() === code) || null;
}

/**
 * Absolute Planer-URL (für Intranet-Link).
 */
export function planerPublicUrl() {
    try {
        return String(window.location.href).split('#')[0].split('?')[0];
    } catch {
        return '';
    }
}

/**
 * Einfaches Embed-Snippet für SharePoint-Seiten.
 * @param {string} [url]
 */
export function buildEmbedSnippet(url) {
    const href = String(url || planerPublicUrl() || 'https://…/tools/schularbeiten-planer.html');
    return (
        '<!-- Schularbeiten-Planer – Link oder Einbettung -->\n' +
        '<div style="padding:12px;border:1px solid #e2e8f0;border-radius:8px;font-family:Segoe UI,sans-serif;">\n' +
        '  <p style="margin:0 0 8px;"><strong>Schularbeiten planen</strong></p>\n' +
        '  <p style="margin:0 0 10px;color:#475569;font-size:14px;">Anträge, Kalender und Freigabe (SchUG/LBVO).</p>\n' +
        '  <a href="' +
        href +
        '" target="_blank" rel="noopener">Planer öffnen</a>\n' +
        '</div>\n' +
        '<!-- Optional iframe (Höhe anpassen): -->\n' +
        '<!-- <iframe src="' +
        href +
        '" style="width:100%;height:900px;border:0;" title="Schularbeiten-Planer"></iframe> -->'
    );
}

/**
 * @param {object[]} items
 * @param {object} filters
 * @param {{ role: string, teacherMatch: object|null, studentMatch?: object|null, demoKlasseCode?: string, accountEmail: string, onlyMine?: boolean }} scope
 */
export function filterSchularbeiten(items, filters, scope) {
    const f = filters || {};
    const list = Array.isArray(items) ? items : [];
    const role = (scope && scope.role) || 'lehrer';
    const onlyMine = !!(scope && scope.onlyMine);
    const teacher = scope && scope.teacherMatch;
    const email = String((scope && scope.accountEmail) || '').toLowerCase();

    if (role === 'schueler') {
        const klasse = resolveStudentKlasseCode(scope);
        return list.filter((sa) => {
            if (!klasse || sa.klasseCode !== klasse) return false;
            if (String(sa.status || '').toLowerCase() !== 'fixiert') return false;
            if (f.fach && sa.fachCode !== f.fach) return false;
            if (f.lehrer && sa.lehrerCode !== f.lehrer) return false;
            return true;
        });
    }

    // Ohne Lehrkraft-Zuordnung (typisch Demo): sonst wäre die UI leer
    const canScopeToSelf = !!(teacher || email);
    const restrictToOwn = (onlyMine || role !== 'admin') && canScopeToSelf;

    return list.filter((sa) => {
        if (restrictToOwn) {
            const codeOk = teacher && sa.lehrerCode && sa.lehrerCode === teacher.code;
            const mailOk =
                (sa.lehrerEmail && email && sa.lehrerEmail === email) ||
                (teacher && teacher.email && sa.lehrerEmail === teacher.email);
            const beantragtOk = sa.beantragtVon && email && String(sa.beantragtVon).toLowerCase() === email;
            if (!codeOk && !mailOk && !beantragtOk) return false;
        }
        if (f.klasse && sa.klasseCode !== f.klasse) return false;
        if (f.fach && sa.fachCode !== f.fach) return false;
        if (f.lehrer && sa.lehrerCode !== f.lehrer) return false;
        if (f.status && sa.status !== f.status) return false;
        return true;
    });
}

/**
 * Labels aus Stammdaten.
 */
export function labelMaps(stammdaten) {
    const fach = {};
    const klasse = {};
    const lehrer = {};
    (stammdaten.subjects || []).forEach((s) => {
        fach[s.code] = s.name || s.code;
    });
    (stammdaten.classes || []).forEach((c) => {
        klasse[c.code] = c.name || c.code;
    });
    (stammdaten.teachers || []).forEach((t) => {
        lehrer[t.code] = t.name || t.code;
    });
    return { fach, klasse, lehrer };
}

export function formatDeDate(iso) {
    if (!iso) return '–';
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso));
    if (!m) return String(iso);
    return `${m[3]}.${m[2]}.${m[1]}`;
}

export function statusLabel(status) {
    const s = String(status || '').toLowerCase();
    if (s === 'fixiert') return 'Fixiert';
    if (s === 'abgelehnt') return 'Abgelehnt';
    return 'Beantragt';
}

export function roleLabel(role) {
    if (role === 'admin') return 'Admin';
    if (role === 'schueler') return 'Schüler';
    return 'Lehrer';
}

/**
 * Darf der User den Antrag bearbeiten/löschen?
 * @param {object} sa
 * @param {{ role: string, teacherMatch: object|null, accountEmail: string }} scope
 */
export function canEditSchularbeit(sa, scope) {
    if (!sa || String(sa.status || '').toLowerCase() !== 'beantragt') return false;
    if ((scope && scope.role) === 'schueler') return false;
    if ((scope && scope.role) === 'admin') return true;
    return isOwnSchularbeit(sa, scope);
}

/**
 * Admin-Aktionen Fixieren/Ablehnen
 */
export function canAdminDecide(sa, scope) {
    if (!sa || String(sa.status || '').toLowerCase() !== 'beantragt') return false;
    if ((scope && scope.role) === 'schueler') return false;
    return (scope && scope.role) === 'admin';
}

/**
 * @param {object} sa
 * @param {{ teacherMatch: object|null, accountEmail: string, role?: string }} scope
 */
export function isOwnSchularbeit(sa, scope) {
    if (!sa) return false;
    if ((scope && scope.role) === 'schueler') return false;
    const email = String((scope && scope.accountEmail) || '').toLowerCase();
    const teacher = scope && scope.teacherMatch;
    if (teacher && sa.lehrerCode && sa.lehrerCode === teacher.code) return true;
    if (sa.lehrerEmail && email && sa.lehrerEmail === email) return true;
    if (teacher && teacher.email && sa.lehrerEmail === teacher.email) return true;
    if (sa.beantragtVon && email && String(sa.beantragtVon).toLowerCase() === email) return true;
    return false;
}
