/**
 * State, Filter, Rollen für Projektwochen.
 */
import { toIsoDateOnly, effectiveBuchungAb, isBookingOpen } from './projektwochen-logic.js';

export const ROLE_STORAGE_KEY = 'ms365-pw-demo-role';
export const SITE_STORAGE_KEY = 'ms365-pw-site-url';

/** @typedef {'lehrer'|'admin'|'schueler'} PwRole */

export const VIEWS = [
    { id: 'dashboard', label: 'Dashboard', icon: 'bi-speedometer2', adminOnly: false, staffOnly: false },
    { id: 'plan', label: 'Wochenplan', icon: 'bi-grid-3x3-gap', adminOnly: false, staffOnly: false },
    { id: 'kalender', label: 'Kalender', icon: 'bi-calendar3', adminOnly: false, staffOnly: false },
    { id: 'liste', label: 'Liste', icon: 'bi-table', adminOnly: false, staffOnly: false },
    { id: 'neu', label: 'Neues Angebot', icon: 'bi-plus-circle', adminOnly: false, staffOnly: true },
    { id: 'meine', label: 'Meine Angebote', icon: 'bi-card-checklist', adminOnly: false, staffOnly: true },
    { id: 'admin', label: 'Administration', icon: 'bi-shield-check', adminOnly: true, staffOnly: true },
    { id: 'einstellungen', label: 'Einstellungen', icon: 'bi-gear', adminOnly: true, staffOnly: true },
    { id: 'bookings', label: 'Bookings', icon: 'bi-box-arrow-up-right', adminOnly: true, staffOnly: true },
    { id: 'teilnehmer', label: 'Teilnehmer', icon: 'bi-people', adminOnly: false, staffOnly: true },
    { id: 'export', label: 'Export', icon: 'bi-download', adminOnly: false, staffOnly: false }
];

/**
 * @param {PwRole|string} role
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
        role: loadRole(),
        filters: { status: '', kategorie: '', tag: '', lehrer: '', klasse: '', q: '', buchung: '' },
        stammdaten: { classes: [], teachers: [], students: [] },
        aktionen: [],
        /** aktive Aktion (offen bevorzugt) */
        aktion: null,
        items: [],
        loading: false,
        bootstrapped: false,
        localDemoOnly: false,
        error: '',
        roleHint: '',
        accountEmail: '',
        accountName: '',
        teacherMatch: null,
        calYear: now.getFullYear(),
        calMonth: now.getMonth(),
        form: emptyForm(),
        editingItemId: null,
        detailId: null,
        detailReturnView: 'dashboard',
        /** Bookings Phase 3 */
        bookingsLog: '',
        bookingsBusy: false,
        attendeeRows: [],
        occupancy: {},
        attendeeFilter: { q: '', angebotId: '', klasse: '' },
        attendeeLoadedAt: ''
    };
}

export function emptyForm(overrides) {
    return {
        title: '',
        beschreibung: '',
        hinweisEltern: '',
        ort: '',
        treffpunkt: '',
        tag: '',
        datum: '',
        slot: 'vormittag',
        startzeit: '08:00',
        endzeit: '12:00',
        kapazitaet: 20,
        preisEuro: 0,
        kostenHinweis: '',
        zielklassen: 'alle',
        lehrerCode: '',
        lehrerEmail: '',
        begleitung: '',
        kategorie: 'workshop',
        buchungAb: '',
        notizIntern: '',
        angebotId: '',
        ...(overrides || {})
    };
}

export function loadStammdaten() {
    let settings = null;
    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            settings = window.ms365TenantSettingsLoad();
        }
    } catch {
        settings = null;
    }
    return {
        classes: Array.isArray(settings && settings.classes) ? settings.classes : [],
        teachers: Array.isArray(settings && settings.teachers) ? settings.teachers : [],
        students: Array.isArray(settings && settings.students) ? settings.students : []
    };
}

export function loadRole() {
    try {
        const r = localStorage.getItem(ROLE_STORAGE_KEY);
        if (r === 'admin' || r === 'lehrer' || r === 'schueler') return r;
    } catch {
        /* ignore */
    }
    return 'lehrer';
}

export function persistRole(role) {
    try {
        localStorage.setItem(ROLE_STORAGE_KEY, String(role || 'lehrer'));
    } catch {
        /* ignore */
    }
}

export function loadSavedSiteUrl() {
    try {
        const u = localStorage.getItem(SITE_STORAGE_KEY);
        if (u) return u;
    } catch {
        /* ignore */
    }
    try {
        if (typeof window !== 'undefined' && window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
            const setup = window.ms365AppDataV2.getSetup();
            if (setup && setup.intranetSiteUrl) return String(setup.intranetSiteUrl).trim();
        }
    } catch {
        /* ignore */
    }
    return '';
}

export function persistSiteUrl(url) {
    const u = String(url || '').trim();
    if (!u) return;
    try {
        localStorage.setItem(SITE_STORAGE_KEY, u);
    } catch {
        /* ignore */
    }
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.patchSetup === 'function'
        ) {
            window.ms365AppDataV2.patchSetup({ intranetSiteUrl: u });
        }
    } catch {
        /* ignore */
    }
}

/**
 * @param {string} email
 * @param {object[]} teachers
 */
export function matchTeacherByEmail(email, teachers) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return null;
    const list = Array.isArray(teachers) ? teachers : [];
    for (let i = 0; i < list.length; i++) {
        const t = list[i];
        if (String((t && t.email) || '')
            .trim()
            .toLowerCase() === em) {
            return t;
        }
    }
    return null;
}

/**
 * Offene Aktion bevorzugen, sonst erste.
 * @param {object[]} aktionen
 */
/**
 * Aktive Projektwoche: bevorzugt „offen“ mit den meisten Angeboten
 * (sonst bleibt eine leere Seed-Aktion aus dem Listen-Setup aktiv).
 * @param {object[]} aktionen
 * @param {object[]} [angebote]
 */
export function pickActiveAktion(aktionen, angebote) {
    const list = (Array.isArray(aktionen) ? aktionen : []).filter((a) => a);
    if (!list.length) return null;
    const items = Array.isArray(angebote) ? angebote : [];
    function countFor(a) {
        const id = String((a && a.aktionId) || '').trim();
        if (!id) return 0;
        let n = 0;
        for (let i = 0; i < items.length; i++) {
            if (items[i] && String(items[i].aktionId || '').trim() === id) n += 1;
        }
        return n;
    }
    const offen = list.filter((a) => a.status === 'offen');
    const pool = offen.length ? offen.slice() : list.slice();
    pool.sort((a, b) => {
        const dc = countFor(b) - countFor(a);
        if (dc !== 0) return dc;
        return String(b.startdatum || '').localeCompare(String(a.startdatum || ''));
    });
    return pool[0] || null;
}

/**
 * @param {object} state
 */
export function scopeFromState(state) {
    return {
        role: state.role || 'lehrer',
        teacherMatch: state.teacherMatch,
        accountEmail: String(state.accountEmail || '')
            .trim()
            .toLowerCase(),
        onlyMine: state.role === 'lehrer'
    };
}

/**
 * @param {object[]} items
 * @param {object} filters
 * @param {object} scope
 * @param {object|null} aktion
 */
export function filterAngebote(items, filters, scope, aktion) {
    const f = filters || {};
    const list = Array.isArray(items) ? items : [];
    const role = (scope && scope.role) || 'lehrer';
    const q = String(f.q || '')
        .trim()
        .toLowerCase();
    const now = new Date();

    return list.filter((a) => {
        if (!a) return false;
        if (aktion && a.aktionId && aktion.aktionId && a.aktionId !== aktion.aktionId) return false;

        if (role === 'schueler') {
            if (a.status !== 'freigegeben') return false;
        } else if (role === 'lehrer' && scope && scope.onlyMine) {
            const code = String((scope.teacherMatch && scope.teacherMatch.code) || '')
                .trim()
                .toLowerCase();
            const em = String(scope.accountEmail || '').toLowerCase();
            const matchCode = code && String(a.lehrerCode || '').toLowerCase() === code;
            const matchMail =
                em &&
                (String(a.lehrerEmail || '').toLowerCase() === em ||
                    String(a.beantragtVon || '').toLowerCase() === em);
            if (scope.teacherMatch || em) {
                if (!matchCode && !matchMail) return false;
            }
        }

        if (f.status && a.status !== f.status) return false;
        if (f.kategorie && a.kategorie !== f.kategorie) return false;
        if (f.tag) {
            const tag = a.tag || '';
            if (tag !== f.tag) return false;
        }
        if (f.lehrer) {
            if (String(a.lehrerCode || '') !== f.lehrer && String(a.lehrerEmail || '') !== f.lehrer) return false;
        }
        if (f.klasse) {
            const z = String(a.zielklassen || '').toLowerCase();
            if (z !== 'alle' && z.split(/[,;]/).map((s) => s.trim()).indexOf(String(f.klasse).toLowerCase()) === -1) {
                return false;
            }
        }
        if (f.buchung === 'offen') {
            if (a.status !== 'freigegeben' || !isBookingOpen(effectiveBuchungAb(a, aktion), now)) return false;
        }
        if (f.buchung === 'gesperrt') {
            if (a.status !== 'freigegeben' || isBookingOpen(effectiveBuchungAb(a, aktion), now)) return false;
        }
        if (q) {
            const hay = [a.title, a.ort, a.beschreibung, a.lehrerCode, a.kategorie].join(' ').toLowerCase();
            if (hay.indexOf(q) === -1) return false;
        }
        return true;
    });
}

export function canEditAngebot(item, scope) {
    if (!item) return false;
    if ((scope && scope.role) === 'admin') return true;
    if ((scope && scope.role) !== 'lehrer') return false;
    const st = item.status;
    if (st !== 'entwurf' && st !== 'beantragt') return false;
    const code = String((scope.teacherMatch && scope.teacherMatch.code) || '')
        .trim()
        .toLowerCase();
    const em = String(scope.accountEmail || '').toLowerCase();
    if (code && String(item.lehrerCode || '').toLowerCase() === code) return true;
    if (em && (String(item.lehrerEmail || '').toLowerCase() === em || String(item.beantragtVon || '').toLowerCase() === em)) {
        return true;
    }
    return !scope.teacherMatch && !em;
}

export function canAdminDecide(scope) {
    return (scope && scope.role) === 'admin';
}

/**
 * @param {object} stammdaten
 */
export function labelMaps(stammdaten) {
    const teachers = {};
    const classes = {};
    ((stammdaten && stammdaten.teachers) || []).forEach((t) => {
        if (t && t.code) teachers[t.code] = t.name || t.code;
    });
    ((stammdaten && stammdaten.classes) || []).forEach((c) => {
        if (c && c.code) classes[c.code] = c.name || c.code;
    });
    return { teachers, classes };
}

/**
 * Form aus Angebot befüllen.
 * @param {object} item
 */
export function formFromItem(item) {
    if (!item) return emptyForm();
    return emptyForm({
        title: item.title || '',
        beschreibung: item.beschreibung || '',
        hinweisEltern: item.hinweisEltern || '',
        ort: item.ort || '',
        treffpunkt: item.treffpunkt || '',
        tag: item.tag || '',
        datum: toIsoDateOnly(item.datum) || '',
        slot: item.slot || 'vormittag',
        startzeit: item.startzeit || '08:00',
        endzeit: item.endzeit || '12:00',
        kapazitaet: item.kapazitaet != null ? item.kapazitaet : 20,
        preisEuro: item.preisEuro != null ? item.preisEuro : 0,
        kostenHinweis: item.kostenHinweis || '',
        zielklassen: item.zielklassen || 'alle',
        lehrerCode: item.lehrerCode || '',
        lehrerEmail: item.lehrerEmail || '',
        begleitung: item.begleitung || '',
        kategorie: item.kategorie || 'workshop',
        buchungAb: item.buchungAb || '',
        notizIntern: item.notizIntern || '',
        angebotId: item.angebotId || ''
    });
}
