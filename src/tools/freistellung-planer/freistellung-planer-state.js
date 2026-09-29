/**
 * State / Filter / Rollen für Freistellungs-Planer.
 */
import { toIsoDateOnly, filterFreistellungen } from './freistellung-planer-logic.js';
import { KATEGORIE_CHOICES, LIST_TITLE_DEFAULT } from './freistellung-planer-schema.js';

export { filterFreistellungen as filterItems };

const ROLE_KEY = 'ms365-freistellung-planer-role-v1';
const SITE_KEY = 'ms365-freistellung-planer-site-v1';
const SETUP_KEY = 'ms365-freistellung-setup-v1';

/** @typedef {'schueler'|'kv'|'direktion'} FrRole */

export const VIEWS = [
    { id: 'dashboard', label: 'Übersicht', icon: 'bi-speedometer2', roles: ['schueler', 'kv', 'direktion'] },
    { id: 'liste', label: 'Liste', icon: 'bi-list-ul', roles: ['schueler', 'kv', 'direktion'] },
    { id: 'antrag', label: 'Neuer Antrag', icon: 'bi-plus-circle', roles: ['schueler', 'kv', 'direktion'] },
    { id: 'meine', label: 'Meine Anträge', icon: 'bi-person', roles: ['schueler'] },
    { id: 'freigabe', label: 'Offene Genehmigungen', icon: 'bi-check2-square', roles: ['kv', 'direktion'] },
    { id: 'bericht', label: 'Berichte', icon: 'bi-bar-chart', roles: ['kv', 'direktion'] }
];

export function viewsForRole(role) {
    const r = role === 'kv' || role === 'direktion' ? role : 'schueler';
    return VIEWS.filter((v) => v.roles.includes(r));
}

export function createInitialState() {
    const setup = loadSetupCfg();
    return {
        role: resolveRole(),
        view: 'dashboard',
        siteUrl: loadSavedSiteUrl(setup),
        listName: (setup && setup.listName) || LIST_TITLE_DEFAULT,
        listId: (setup && setup.listId) || '',
        emailDirektion: (setup && setup.emailDirektion) || '',
        loading: false,
        error: '',
        info: '',
        accountEmail: '',
        accountName: '',
        items: [],
        stammdaten: { classes: [], teachers: [] },
        filters: { klasse: '', status: '', kategorie: '', multiDay: '', q: '' },
        form: emptyForm(),
        editingItemId: null,
        detailId: null,
        ctx: null,
        localDemoOnly: false
    };
}

export function emptyForm(partial) {
    const iso = toIsoDateOnly(new Date()) || '';
    return {
        titel: '',
        schuelerName: '',
        klasse: '',
        beginn: iso,
        ende: iso,
        kategorie: KATEGORIE_CHOICES[0],
        beschreibung: '',
        kvEmail: '',
        kvName: '',
        status: 'Ausstehend',
        bemerkungen: '',
        ...(partial || {})
    };
}

export function resolveRole(preferred) {
    if (preferred === 'schueler' || preferred === 'kv' || preferred === 'direktion') return preferred;
    try {
        const s = String(localStorage.getItem(ROLE_KEY) || '').toLowerCase();
        if (s === 'schueler' || s === 'kv' || s === 'direktion') return s;
    } catch {
        /* ignore */
    }
    return 'schueler';
}

export function persistRole(role) {
    const r = role === 'kv' || role === 'direktion' ? role : 'schueler';
    try {
        localStorage.setItem(ROLE_KEY, r);
    } catch {
        /* ignore */
    }
}

export function loadSetupCfg() {
    try {
        return JSON.parse(localStorage.getItem(SETUP_KEY) || '{}') || {};
    } catch {
        return {};
    }
}

export function loadSavedSiteUrl(setup) {
    try {
        const local = String(localStorage.getItem(SITE_KEY) || '').trim();
        if (local) return local;
    } catch {
        /* ignore */
    }
    if (setup && setup.siteUrl) return String(setup.siteUrl).trim();
    try {
        const s =
            window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup()
                : null;
        if (s && s.intranetSiteUrl) return String(s.intranetSiteUrl).trim();
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

/**
 * KV aus Stammdaten-Klasse ableiten.
 * @param {object[]} classes
 * @param {string} klasseCode
 */
export function resolveKvForClass(classes, klasseCode) {
    const code = String(klasseCode || '').trim();
    if (!code) return null;
    const list = Array.isArray(classes) ? classes : [];
    const c =
        list.find((x) => String(x.code || '').trim() === code) ||
        list.find((x) => String(x.name || '').trim() === code) ||
        null;
    if (!c) return null;
    const email = String(c.headEmail || c.klassenvorstandEmail || c.kvEmail || '')
        .trim()
        .toLowerCase();
    const name = String(c.headName || c.klassenvorstandName || c.kvName || '').trim();
    if (!email) return null;
    return { email, name, classCode: String(c.code || code) };
}

export function roleLabel(role) {
    if (role === 'kv') return 'Klassenvorstand';
    if (role === 'direktion') return 'Direktion / Admin';
    return 'Schüler/in';
}

export function statusLabel(status) {
    const s = String(status || '').trim();
    if (s === 'Genehmigt') return 'Genehmigt';
    if (s === 'Abgelehnt') return 'Abgelehnt';
    return 'Ausstehend';
}

export function formatDeDate(iso) {
    const s = toIsoDateOnly(iso);
    if (!s) return '–';
    const [y, m, d] = s.split('-');
    return `${d}.${m}.${y}`;
}

export function scopeFromState(state, opts) {
    const o = opts || {};
    const role = state.role;
    if (o.scopeAll || role === 'direktion') {
        return { accountEmail: state.accountEmail };
    }
    if (role === 'kv') {
        return { onlyKv: true, accountEmail: state.accountEmail };
    }
    return { onlyMine: true, accountEmail: state.accountEmail };
}

export function canDecide(state) {
    return state.role === 'kv' || state.role === 'direktion';
}

export function buildAntragTitle(form) {
    const name = String(form.schuelerName || form.titel || '').trim() || 'Freistellung';
    const klasse = String(form.klasse || '').trim();
    return klasse ? `${name} (${klasse})` : name;
}
