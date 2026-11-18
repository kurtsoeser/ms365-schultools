/**
 * State / Rollen / Views / Filter für sRDP-Anmeldung App.
 */
import { STORAGE_KEY, getProfile, listTitleForYear } from './srdp-anmeldung-schema.js';
import {
    filterAbschlussklassen,
    teacherChoiceLabels,
    parseChoiceLines,
    filterItemsForTeacher,
    filterItemsForKlassen,
    parseCodeList,
    joinCodeList,
    countNacharbeitPending,
    klassenOhneAnmeldung
} from './srdp-anmeldung-logic.js';
import {
    resolveSchoolToolSiteUrl,
    patchSchoolToolSitesSetup
} from '../../shared/school-tool-sites.js';

const ROLE_KEY = 'ms365-srdp-anmeldung-role-v1';
const SITE_OVERRIDE_KEY = 'ms365-srdp-planer-site-v1';

/** @typedef {'schueler'|'lehrer'|'admin'} SrdpRole */

export const VIEWS = [
    { id: 'dashboard', label: 'Übersicht', icon: 'bi-house', roles: ['lehrer', 'admin'] },
    { id: 'anmelden', label: 'Anmelden', icon: 'bi-pencil-square', roles: ['schueler'] },
    { id: 'meine', label: 'Meine Anmeldung', icon: 'bi-person-check', roles: ['schueler'] },
    { id: 'meine-pruefungen', label: 'Meine Prüfungen', icon: 'bi-mortarboard', roles: ['lehrer'] },
    { id: 'meine-klasse', label: 'Meine Klasse', icon: 'bi-people', roles: ['lehrer'] },
    { id: 'liste', label: 'Alle Anmeldungen', icon: 'bi-list-ul', roles: ['lehrer', 'admin'] },
    { id: 'kalender', label: 'Kalender', icon: 'bi-calendar3', roles: ['lehrer', 'admin'] },
    { id: 'auswertung', label: 'Auswertung', icon: 'bi-bar-chart', roles: ['admin'] },
    { id: 'system', label: 'Systemadministration', icon: 'bi-gear', roles: ['admin'] }
];

/** @param {string} id */
function viewById(id) {
    return VIEWS.find((v) => v.id === id) || null;
}

/** Setup-/Site-Toolbar nur für Admin – KV bleibt Lehrer-Capability. */
export function isSrdpStaffRole(role) {
    return role === 'admin';
}

/**
 * Gruppierte Navigation nach Rolle.
 * - Schüler / Lehrer: flach
 * - Admin: Arbeit · Auswertung · System
 *
 * @param {SrdpRole|string} role
 * @param {string[]|null|undefined} [_availableRoles]
 * @param {{ kvKlassen?: string[] }|null} [opts]
 * @returns {{ id: string, label: string, items: typeof VIEWS }[]}
 */
export function navSectionsForRole(role, _availableRoles, opts) {
    const r = role === 'admin' ? 'admin' : role === 'lehrer' ? 'lehrer' : 'schueler';
    const kvKlassen = Array.isArray(opts && opts.kvKlassen) ? opts.kvKlassen : [];

    if (r === 'admin') {
        return [
            {
                id: 'arbeit',
                label: 'Arbeit',
                items: [viewById('dashboard'), viewById('liste'), viewById('kalender')].filter(Boolean)
            },
            {
                id: 'auswertung',
                label: 'Auswertung',
                items: [viewById('auswertung')].filter(Boolean)
            },
            {
                id: 'system',
                label: 'System',
                items: [viewById('system')].filter(Boolean)
            }
        ];
    }

    if (r === 'lehrer') {
        const items = [
            viewById('dashboard'),
            viewById('meine-pruefungen'),
            viewById('kalender'),
            ...(kvKlassen.length ? [viewById('meine-klasse')] : []),
            viewById('liste')
        ].filter(Boolean);
        return [{ id: 'arbeit', label: '', items }];
    }

    return [
        {
            id: 'arbeit',
            label: '',
            items: [viewById('anmelden'), viewById('meine')].filter(Boolean)
        }
    ];
}

/**
 * Flache View-Liste (Allowed-Checks / Tests).
 * @param {SrdpRole|string} role
 * @param {{ kvKlassen?: string[] }|null} [opts]
 */
export function viewsForRole(role, opts) {
    return navSectionsForRole(role, null, opts).flatMap((s) => s.items || []);
}

export function roleLabel(role) {
    if (role === 'admin') return 'Admin / Sekretariat';
    if (role === 'lehrer') return 'Lehrkraft';
    return 'Schüler:in';
}

/** KV = Capability über Stammdaten, kein Dauer-App-Modus. */
export function hasSrdpKvCapability(state) {
    if (!state) return false;
    if (state.role === 'admin' && !state.demoPersonaEmail) return true;
    return kvKlasseCodesForState(state).length > 0;
}

export function loadSavedRole() {
    try {
        const r = String(localStorage.getItem(ROLE_KEY) || '').toLowerCase();
        if (r === 'admin' || r === 'lehrer' || r === 'schueler') return r;
    } catch {
        /* ignore */
    }
    return 'schueler';
}

export function persistRole(role) {
    try {
        const r = role === 'admin' ? 'admin' : role === 'lehrer' ? 'lehrer' : 'schueler';
        localStorage.setItem(ROLE_KEY, r);
    } catch {
        /* ignore */
    }
}

export function loadWizardConfig() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return {};
        const o = JSON.parse(raw);
        return o && typeof o === 'object' ? o : {};
    } catch {
        return {};
    }
}

/**
 * @param {object} patch
 */
export function persistWizardConfig(patch) {
    const prev = loadWizardConfig();
    const next = Object.assign({}, prev, patch && typeof patch === 'object' ? patch : {});
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

export function loadSavedSiteUrl() {
    return resolveSchoolToolSiteUrl('srdp') || '';
}

export function persistSiteUrl(url) {
    const u = String(url || '').trim();
    try {
        if (u) localStorage.setItem(SITE_OVERRIDE_KEY, u);
        else localStorage.removeItem(SITE_OVERRIDE_KEY);
    } catch {
        /* ignore */
    }
    try {
        patchSchoolToolSitesSetup({ planerSiteUrls: { srdp: u } });
    } catch {
        /* ignore */
    }
}

export function loadStammdaten() {
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

export function matchTeacherByEmail(teachers, email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return null;
    const list = Array.isArray(teachers) ? teachers : [];
    return (
        list.find((t) => String(t.email || '').trim().toLowerCase() === em) ||
        list.find((t) => String(t.upn || '').trim().toLowerCase() === em) ||
        null
    );
}

export function teacherDisplayLabel(teacher) {
    if (!teacher) return '';
    return String(teacher.name || teacher.code || '').trim();
}

export function kvKlasseCodesForEmail(stammdaten, email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return [];
    const classes = (stammdaten && stammdaten.classes) || [];
    return classes
        .filter((c) => {
            const head = String(c.headEmail || c.kvEmail || c.classTeacherEmail || '')
                .trim()
                .toLowerCase();
            return head && head === em;
        })
        .map((c) => String(c.code || c.name || '').trim())
        .filter(Boolean);
}

export function effectiveAccountEmail(state) {
    return String((state && (state.demoPersonaEmail || state.accountEmail)) || '')
        .trim()
        .toLowerCase();
}

export function kvKlasseCodesForState(state) {
    if (!state) return [];
    if (Array.isArray(state.kvKlassen) && state.kvKlassen.length) return state.kvKlassen;
    return kvKlasseCodesForEmail(state.stammdaten, effectiveAccountEmail(state));
}

export function matchStudentByEmail(students, email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return null;
    const list = Array.isArray(students) ? students : [];
    return (
        list.find((s) => String(s.email || '').trim().toLowerCase() === em) ||
        list.find((s) => String(s.upn || '').trim().toLowerCase() === em) ||
        null
    );
}

/**
 * Optionen aus Wizard + Profil.
 */
export function resolveFormOptions(wizard, stammdaten, profileId) {
    const wiz = wizard || {};
    const profile = getProfile(profileId || wiz.profileId || 'hak');
    const jahr = String(wiz.jahr || wiz.terminJahr || new Date().getFullYear()).trim();
    const klassenFromSd = filterAbschlussklassen(stammdaten && stammdaten.classes, jahr);
    const klassenWiz = Array.isArray(wiz.klassen) ? wiz.klassen.filter(Boolean) : [];
    const klassen = klassenWiz.length ? klassenWiz : klassenFromSd;
    const lehrer = teacherChoiceLabels((stammdaten && stammdaten.teachers) || []);
    const lfs = parseChoiceLines(wiz.lfsText || '').length
        ? parseChoiceLines(wiz.lfsText)
        : Array.isArray(wiz.lfs)
          ? wiz.lfs
          : (profile.lfsHints || []).slice();
    const wahlfaecher = parseChoiceLines(wiz.wahlfachText || '').length
        ? parseChoiceLines(wiz.wahlfachText)
        : Array.isArray(wiz.wahlfaecher)
          ? wiz.wahlfaecher
          : (profile.defaultWahlfaecher || []).slice();
    const seminare = parseChoiceLines(wiz.seminarText || '').length
        ? parseChoiceLines(wiz.seminarText)
        : Array.isArray(wiz.seminare)
          ? wiz.seminare
          : (profile.defaultSeminare || []).slice();
    return {
        profileId: profile.id,
        profileLabel: profile.label || profile.id,
        jahr,
        listTitle: listTitleForYear(jahr, profile.id),
        klassen,
        lehrer,
        lfs,
        wahlfaecher,
        seminare,
        variants: profile.variants || [],
        anmeldeFrist: String(wiz.anmeldeFrist || '').trim(),
        haupttermin: String(wiz.haupttermin || '').trim()
    };
}

export function emptyForm(seed) {
    const s = seed || {};
    return {
        klasse: String(s.klasse || '').trim(),
        nachname: String(s.nachname || '').trim(),
        vorname: String(s.vorname || '').trim(),
        titelDiplomarbeit: String(s.titelDiplomarbeit || '').trim(),
        betreuungslehrerDA: String(s.betreuungslehrerDA || '').trim(),
        variante: String(s.variante || '').trim(),
        lfs: String(s.lfs || '').trim(),
        lehrerLFS: String(s.lehrerLFS || '').trim(),
        lehrerBKO: String(s.lehrerBKO || '').trim(),
        wahlfachMuendlich: String(s.wahlfachMuendlich || s.wahlfach || '').trim(),
        seminar: String(s.seminar || '').trim(),
        lehrerWahlfach: String(s.lehrerWahlfach || '').trim(),
        bestaetigung: !!s.bestaetigung,
        wizardStep: Number(s.wizardStep) > 0 ? Number(s.wizardStep) : 1
    };
}

export function createInitialState() {
    const wizard = loadWizardConfig();
    const stammdaten = loadStammdaten();
    const opts = resolveFormOptions(wizard, stammdaten, wizard.profileId);
    const today = new Date();
    return {
        role: loadSavedRole(),
        view: 'anmelden',
        planerRoles: /** @type {string[]} */ ([]),
        demoRoleOverride: false,
        entraGroupsConfigured: true,
        planerAccessDenied: false,
        roleHint: '',
        accountEmail: '',
        accountName: '',
        demoPersonaEmail: '',
        demoPersonaName: '',
        siteUrl: loadSavedSiteUrl(),
        listId: '',
        listWebUrl: '',
        siteId: '',
        loading: false,
        saving: false,
        error: '',
        softError: '',
        localDemoOnly: false,
        /** @type {{ source: string, lastPullAt: string, listTitle: string, connected: boolean }} */
        syncMeta: {
            source: 'none',
            lastPullAt: '',
            listTitle: '',
            connected: false
        },
        items: /** @type {object[]} */ ([]),
        stammdaten,
        stammdatenLive: {
            ok: false,
            error: '',
            classCount: 0,
            teacherCount: 0,
            missing: /** @type {string[]} */ ([]),
            source: 'cache'
        },
        wizard,
        options: opts,
        form: emptyForm(),
        formErrors: /** @type {string[]} */ ([]),
        submitSuccess: null,
        filterKlasse: '',
        filterVariante: '',
        filterLehrer: '',
        filterQuery: '',
        filterMultiOpen: /** @type {Record<string, boolean>} */ ({}),
        filterMultiQuery: /** @type {Record<string, string>} */ ({}),
        detailItemId: '',
        teacherLabel: '',
        kvKlassen: /** @type {string[]} */ ([]),
        calMode: 'month',
        calYear: today.getFullYear(),
        calMonth: today.getMonth() + 1,
        calWeekAnchor: today.toISOString().slice(0, 10)
    };
}

/**
 * @param {object[]} items
 * @param {object} state
 */
export function filterItemsForView(items, state) {
    const list = Array.isArray(items) ? items : [];
    const q = String(state.filterQuery || '')
        .trim()
        .toLowerCase();
    const klassen = parseCodeList(state.filterKlasse);
    const varianten = parseCodeList(state.filterVariante);
    const lehrer = parseCodeList(state.filterLehrer);
    return list.filter((it) => {
        if (klassen.length && !klassen.includes(String(it.klasse || '').trim())) return false;
        if (varianten.length && !varianten.includes(String(it.variante || '').trim())) return false;
        if (lehrer.length) {
            const blob = [
                it.betreuungslehrerDA,
                it.lehrerLFS,
                it.lehrerBKO,
                it.lehrerWahlfach
            ]
                .map((x) => String(x || '').trim())
                .filter(Boolean);
            if (!lehrer.some((l) => blob.includes(l))) return false;
        }
        if (!q) return true;
        const hay = [
            it.nachname,
            it.vorname,
            it.klasse,
            it.variante,
            it.wahlfachMuendlich,
            it.lfs,
            it.title,
            it.schuelerEmail,
            it.createdBy
        ]
            .join(' ')
            .toLowerCase();
        return hay.includes(q);
    });
}

/**
 * Filter-Jump: gewünschte Filter setzen, Rest leeren.
 * @param {object} state
 * @param {{ klasse?: string|string[], variante?: string|string[], lehrer?: string|string[], query?: string }} jump
 */
export function applyFilterJump(state, jump) {
    const j = jump || {};
    state.filterKlasse = joinCodeList(j.klasse || '');
    state.filterVariante = joinCodeList(j.variante || '');
    state.filterLehrer = joinCodeList(j.lehrer || '');
    state.filterQuery = String(j.query || '').trim();
    state.filterMultiOpen = {};
    state.filterMultiQuery = {};
}

export function itemsForCurrentRole(state) {
    const items = Array.isArray(state.items) ? state.items : [];
    if (state.role === 'admin') return filterItemsForView(items, state);
    if (state.role === 'lehrer') {
        if (state.view === 'meine-klasse') {
            return filterItemsForKlassen(items, kvKlasseCodesForState(state));
        }
        if (state.view === 'meine-pruefungen' || state.view === 'dashboard') {
            return filterItemsForTeacher(items, state.teacherLabel);
        }
        return filterItemsForView(items, state);
    }
    const em = effectiveAccountEmail(state);
    const name = String(state.demoPersonaName || state.accountName || '')
        .trim()
        .toLowerCase();
    return items.filter((it) => {
        const owned = String(it.schuelerEmail || it.beantragtVon || it.createdBy || '')
            .trim()
            .toLowerCase();
        if (em && owned && owned === em) return true;
        const title = String(it.title || '').trim().toLowerCase();
        if (name && title && title.indexOf(name.split(' ')[0]) !== -1) return true;
        const nn = String(it.nachname || '').trim().toLowerCase();
        const vn = String(it.vorname || '').trim().toLowerCase();
        if (name && nn && name.includes(nn) && vn && name.includes(vn)) return true;
        return false;
    });
}

export function findItemById(state, id) {
    const want = String(id || '').trim();
    if (!want) return null;
    return (state.items || []).find((it) => String(it.itemId || '') === want) || null;
}

export { filterItemsForTeacher, filterItemsForKlassen, STORAGE_KEY, countNacharbeitPending, klassenOhneAnmeldung, parseCodeList, joinCodeList };
