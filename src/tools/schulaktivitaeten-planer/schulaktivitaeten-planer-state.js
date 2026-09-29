/**
 * State / Filter / Rollen für Schulaktivitäten-Planer.
 */
import { toIsoDateOnly } from './schulaktivitaeten-planer-logic.js';
import { TYP_CHOICES } from './schulaktivitaeten-planer-schema.js';

const ROLE_KEY = 'ms365-akt-planer-role-v1';
const SITE_KEY = 'ms365-akt-planer-site-v1';

/** @typedef {'lehrer'|'admin'} AktRole */

export const VIEWS = [
    { id: 'dashboard', label: 'Übersicht', icon: 'bi-speedometer2', roles: ['lehrer', 'admin'] },
    { id: 'liste', label: 'Liste', icon: 'bi-list-ul', roles: ['lehrer', 'admin'] },
    { id: 'kalender', label: 'Kalender', icon: 'bi-calendar3', roles: ['lehrer', 'admin'] },
    { id: 'antrag', label: 'Neuer Antrag', icon: 'bi-plus-circle', roles: ['lehrer', 'admin'] },
    { id: 'freigabe', label: 'Freigabe', icon: 'bi-check2-square', roles: ['admin'] },
    { id: 'regeln', label: 'Regelwerk', icon: 'bi-sliders', roles: ['admin'] }
];

export function viewsForRole(role) {
    const r = role === 'admin' ? 'admin' : 'lehrer';
    return VIEWS.filter((v) => v.roles.includes(r));
}

export function createInitialState() {
    return {
        role: resolveRole(),
        view: 'dashboard',
        siteUrl: loadSavedSiteUrl(),
        loading: false,
        error: '',
        roleHint: '',
        accountEmail: '',
        accountName: '',
        items: [],
        rules: null,
        stammdaten: { classes: [], teachers: [] },
        filters: { klasse: '', typ: '', status: '', lehrer: '' },
        form: emptyForm(),
        editingItemId: null,
        detailId: null,
        calYear: new Date().getFullYear(),
        calMonth: new Date().getMonth() + 1,
        ctx: null,
        localDemoOnly: false
    };
}

export function emptyForm(partial) {
    const t = new Date();
    const iso = toIsoDateOnly(t) || '';
    return {
        aktivitaetId: '',
        titel: '',
        typ: 'Exkursion',
        klasseCode: '',
        lehrerCode: '',
        lehrerEmail: '',
        begleitung: '',
        ort: '',
        startdatum: iso,
        enddatum: iso,
        startZeit: '',
        endZeit: '',
        status: 'beantragt',
        notiz: '',
        ablehnungsGrund: '',
        verkehrsmittel: '',
        kostenHinweis: '',
        ...(partial || {})
    };
}

export function resolveRole(preferred) {
    if (preferred === 'admin' || preferred === 'lehrer') return preferred;
    try {
        const s = String(localStorage.getItem(ROLE_KEY) || '').toLowerCase();
        if (s === 'admin' || s === 'lehrer') return s;
    } catch {
        /* ignore */
    }
    return 'lehrer';
}

export function persistRole(role) {
    try {
        localStorage.setItem(ROLE_KEY, role === 'admin' ? 'admin' : 'lehrer');
    } catch {
        /* ignore */
    }
}

export function loadSavedSiteUrl() {
    try {
        const local = String(localStorage.getItem(SITE_KEY) || '').trim();
        if (local) return local;
    } catch {
        /* ignore */
    }
    try {
        const setup =
            window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function'
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
        localStorage.setItem(SITE_KEY, String(url || '').trim());
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
            teachers: Array.isArray(data.teachers) ? data.teachers : []
        };
    } catch {
        return { classes: [], teachers: [] };
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

export function labelMaps(stammdaten) {
    const klasse = {};
    const lehrer = {};
    (stammdaten.classes || []).forEach((c) => {
        if (c && c.code) klasse[c.code] = c.name || c.code;
    });
    (stammdaten.teachers || []).forEach((t) => {
        if (t && t.code) lehrer[t.code] = t.name || t.code;
    });
    return { klasse, lehrer };
}

export function statusLabel(status) {
    const s = String(status || '').toLowerCase();
    if (s === 'genehmigt') return 'Genehmigt';
    if (s === 'abgelehnt') return 'Abgelehnt';
    return 'Beantragt';
}

export function typLabel(typ) {
    const t = String(typ || '');
    if (t === 'Schulaktivitaet') return 'Schulaktivität';
    return t || '–';
}

export function roleLabel(role) {
    return role === 'admin' ? 'Admin / Direktion' : 'Lehrkraft';
}

export function formatDeDate(iso) {
    const s = toIsoDateOnly(iso);
    if (!s) return '–';
    const [y, m, d] = s.split('-');
    return `${d}.${m}.${y}`;
}

export function typChoices() {
    return TYP_CHOICES.map((c) => ({ code: c, name: typLabel(c) }));
}

/**
 * @param {object[]} items
 * @param {object} filters
 * @param {{ role: string, accountEmail?: string, onlyMine?: boolean, onlyOpen?: boolean }} scope
 */
export function filterItems(items, filters, scope) {
    const f = filters || {};
    const list = Array.isArray(items) ? items : [];
    const email = String((scope && scope.accountEmail) || '')
        .trim()
        .toLowerCase();
    const onlyMine = !!(scope && scope.onlyMine);
    const onlyOpen = !!(scope && scope.onlyOpen);
    return list.filter((it) => {
        if (onlyOpen && String(it.status || '').toLowerCase() !== 'beantragt') return false;
        if (onlyMine && email) {
            const em = String(it.lehrerEmail || '').toLowerCase();
            const von = String(it.beantragtVon || '').toLowerCase();
            if (em !== email && von !== email) return false;
        }
        if (f.klasse && it.klasseCode !== f.klasse) return false;
        if (f.typ && it.typ !== f.typ) return false;
        if (f.status && String(it.status || '').toLowerCase() !== String(f.status).toLowerCase()) return false;
        if (f.lehrer && it.lehrerCode !== f.lehrer) return false;
        return true;
    });
}

export function canEditItem(item, state) {
    if (!item) return false;
    if (state.role === 'admin') return true;
    const st = String(item.status || '').toLowerCase();
    if (st !== 'beantragt') return false;
    const email = String(state.accountEmail || '')
        .toLowerCase()
        .trim();
    if (!email) return true;
    return (
        String(item.lehrerEmail || '').toLowerCase() === email ||
        String(item.beantragtVon || '').toLowerCase() === email
    );
}

export function canDecide(state) {
    return state.role === 'admin';
}

export function scopeFromState(state, opts) {
    const o = opts || {};
    return {
        role: state.role,
        accountEmail: state.accountEmail,
        onlyMine: !!o.onlyMine || (state.role !== 'admin' && !o.scopeAll),
        onlyOpen: !!o.onlyOpen
    };
}
