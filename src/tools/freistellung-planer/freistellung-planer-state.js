/**
 * State / Filter / Rollen für Freistellungs-Planer.
 */
import { toIsoDateOnly, filterFreistellungen } from './freistellung-planer-logic.js';
import { LIST_TITLE_DEFAULT } from './freistellung-planer-schema.js';
import { mergeKategorieChoices, loadExtraKategorien, KATEGORIE_CHOICES } from './freistellung-planer-kategorien.js';
import { accountIsDirektionPlannerUser } from './freistellung-planer-direktion-users.js';
import { loadPermissionsConfig } from './freistellung-planer-permissions.js';
import { enrichClassesFromLinkedGroups } from '../../shared/class-list-enrich.js';
import {
    collectAllClassRows,
    collectAllStudentRows,
    loadClassTeamsContext,
    classRowsFromTeamLinks,
    classTeamLinksFromAppData
} from './freistellung-planer-class-context.js';
import { loadEffectivePermissionsConfig } from './freistellung-planer-permissions.js';

export { filterFreistellungen as filterItems };

const ROLE_KEY = 'ms365-freistellung-planer-role-v1';
export const ROLE_STORAGE_KEY = ROLE_KEY;
const SITE_KEY = 'ms365-freistellung-planer-site-v1';
const SETUP_KEY = 'ms365-freistellung-setup-v1';
const STUDENT_KLASSE_PICK_KEY = 'ms365-freistellung-student-klasse-pick-v1';

/** @typedef {'schueler'|'kv'|'direktion'} FrRole */

export const VIEWS = [
    { id: 'dashboard', label: 'Übersicht', icon: 'bi-speedometer2', roles: ['kv', 'direktion'] },
    { id: 'liste', label: 'Liste', icon: 'bi-list-ul', roles: ['kv', 'direktion'] },
    { id: 'meine', label: 'Meine Anträge', icon: 'bi-person', roles: ['schueler'] },
    { id: 'antrag', label: 'Antrag stellen', icon: 'bi-plus-circle', roles: ['kv', 'direktion', 'schueler'] },
    { id: 'freigabe', label: 'Offene Genehmigungen', icon: 'bi-check2-square', roles: ['kv', 'direktion'] },
    { id: 'bericht', label: 'Berichte', icon: 'bi-bar-chart', roles: ['kv', 'direktion'] },
    { id: 'administration', label: 'Administration', icon: 'bi-gear-wide-connected', roles: ['direktion'] }
];

/**
 * @param {FrRole|string} role
 * @param {object} [state]
 */
export function viewsForRole(role, state) {
    const r = role === 'kv' || role === 'direktion' ? role : 'schueler';
    if (state && useKvPlanerChrome(state)) {
        return VIEWS.filter((v) => v.id === 'dashboard');
    }
    return VIEWS.filter((v) => v.roles.includes(r));
}

/**
 * Lehrkraft / Verwaltung (auch wenn die Planer-Rolle noch nicht aufgelöst ist).
 * @param {object} state
 */
export function isLikelyFreistellungStaffAccount(state) {
    if (!state) return false;
    if (state.kvMatch || state.direktionMatch) return true;
    const em = String(state.accountEmail || '')
        .trim()
        .toLowerCase();
    if (!em) return false;
    if (accountIsDirektionPlannerUser(em, loadPermissionsConfig().direktionUsers)) return true;
    const teachers = (state.stammdaten && state.stammdaten.teachers) || [];
    return teachers.some((t) => String(t.email || '').trim().toLowerCase() === em);
}

/**
 * Schüler-UI ohne IT-Leiste (nur echte Schüler, nicht KV/Direktion/Lehrkraft).
 */
export function useStudentPlanerChrome(state) {
    if (!state || state.demoRoleOverride) return false;
    if (state.role === 'kv' || state.role === 'direktion') return false;
    const roles = state.planerRoles || [];
    if (roles.includes('kv') || roles.includes('direktion')) return false;
    if (isLikelyFreistellungStaffAccount(state)) return false;
    return true;
}

/**
 * Klassenvorstand ohne IT-Leiste (SharePoint-URL, Setup, CSV) – nur Schüler-Übersicht.
 */
export function useKvPlanerChrome(state) {
    if (!state || state.demoRoleOverride) return false;
    if (state.role !== 'kv') return false;
    try {
        const p = new URLSearchParams(typeof window !== 'undefined' ? window.location.search || '' : '');
        if (p.get('demoRole') === '1') return false;
    } catch {
        /* ignore */
    }
    return true;
}

/**
 * Direktion ohne IT-Leiste in Übersicht/Liste – Technik nur unter „Administration“.
 */
export function useDirektionPlanerChrome(state) {
    if (!state || state.demoRoleOverride) return false;
    if (state.role !== 'direktion') return false;
    try {
        const p = new URLSearchParams(typeof window !== 'undefined' ? window.location.search || '' : '');
        if (p.get('demoRole') === '1') return false;
    } catch {
        /* ignore */
    }
    return true;
}

/** Schüler-, KV- oder Direktion-Arbeitsansicht ohne Technik-Chrome oben. */
export function useMinimalPlanerChrome(state) {
    return (
        useStudentPlanerChrome(state) || useKvPlanerChrome(state) || useDirektionPlanerChrome(state)
    );
}

export function createInitialState() {
    const setup = loadSetupCfg();
    return {
        role: resolveRole(),
        view: 'meine',
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
        localDemoOnly: false,
        entraGroupsConfigured: false,
        demoRoleOverride: false,
        planerRoles: [],
        planerRoleSources: {},
        planerAccessDenied: false,
        roleSource: 'demo',
        roleHint: '',
        roleHintStaff: '',
        roleHintPublic: '',
        studentMatch: null,
        kvMatch: null,
        direktionMatch: false,
        demoKlasseCode: '',
        kategorieChoices: mergeKategorieChoices(loadExtraKategorien()),
        /** Aus SharePoint-Liste „Freistellungen“, Spalte Klasse (Choice). */
        klasseColumnChoices: []
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
        kategorie: mergeKategorieChoices(loadExtraKategorien())[0] || KATEGORIE_CHOICES[0],
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

/** @param {Record<string, unknown>} patch */
export function persistSetupFields(patch) {
    const p = patch && typeof patch === 'object' ? patch : {};
    const setup = loadSetupCfg();
    const next = Object.assign({}, setup, p);
    try {
        localStorage.setItem(SETUP_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

/** Mandanten-Stammweb ohne /sites/… – dort liegt die Freistellungsliste selten. */
export function isLikelySharePointTenantRoot(url) {
    try {
        const u = new URL(String(url || '').trim());
        const path = u.pathname.replace(/\/+$/, '');
        return path === '' || path === '/';
    } catch {
        return false;
    }
}

export function loadSavedSiteUrl(setup) {
    const cfg = setup || loadSetupCfg();
    try {
        const local = String(localStorage.getItem(SITE_KEY) || '').trim();
        if (local && !isLikelySharePointTenantRoot(local)) return local;
        if (local && isLikelySharePointTenantRoot(local)) {
            /* Altes Auto-Fallback – Setup/Discovery bevorzugen */
        }
    } catch {
        /* ignore */
    }
    if (cfg && cfg.siteUrl) return String(cfg.siteUrl).trim();
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

/** Site für Planer/Setup (lokal, Setup-JSON, Stammdaten-Intranet). */
export function resolveFreistellungSiteUrl(current) {
    const direct = String(current || '').trim();
    if (direct) return direct;
    return loadSavedSiteUrl(loadSetupCfg());
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
        if (
            typeof window !== 'undefined' &&
            window.ms365StammdatenCanonical &&
            typeof window.ms365StammdatenCanonical.reconcileStammdatenStorage === 'function'
        ) {
            try {
                window.ms365StammdatenCanonical.reconcileStammdatenStorage();
            } catch {
                /* ignore */
            }
        }
        const core = window.ms365TenantSettingsLoad && window.ms365TenantSettingsLoad();
        const data = (core && core.data) || core || {};
        let baseClasses = Array.isArray(data.classes) ? data.classes : [];
        if (
            typeof window !== 'undefined' &&
            window.ms365StammdatenCanonical &&
            typeof window.ms365StammdatenCanonical.listAllClasses === 'function'
        ) {
            const canon = window.ms365StammdatenCanonical.listAllClasses();
            if (canon && canon.length) baseClasses = canon;
        }
        return {
            classes: collectAllClassRows({ stammdaten: { classes: baseClasses } }),
            teachers: Array.isArray(data.teachers) ? data.teachers : [],
            students: collectAllStudentRows()
        };
    } catch {
        return { classes: [], teachers: [], students: [] };
    }
}

/**
 * Nach SharePoint-Stammdaten-Pull: Schülerzeile (Klasse) neu zuordnen.
 * @param {object} state
 */
export function refreshStudentMatchFromStammdaten(state) {
    if (!state || state.role !== 'schueler') return false;
    state.stammdaten = loadStammdaten();
    const hit = matchStudentByEmail(state.accountEmail, state.stammdaten.students);
    if (!hit || !studentKlasseFromRecord(hit)) return false;
    state.studentMatch = {
        ...hit,
        email: state.accountEmail,
        klasseSource: 'stammdaten'
    };
    state.demoKlasseCode = '';
    persistDemoKlasseCode('');
    prefillStudentFreistellungForm(state);
    return true;
}

/** @param {object} s */
export function studentKlasseFromRecord(s) {
    if (!s || typeof s !== 'object') return '';
    const k = s.klasse || s.class || s.classCode || s.ktKlasse || s.Klasse || '';
    return String(k).trim();
}

/** @param {string} codeOrName */
export function findClassInStammdaten(classes, codeOrName) {
    const q = String(codeOrName || '').trim();
    if (!q) return null;
    const list = Array.isArray(classes) ? classes : [];
    const ql = q.toLowerCase();
    return (
        list.find((c) => String(c.code || '').trim() === q) ||
        list.find((c) => String(c.name || '').trim() === q) ||
        list.find((c) => String(c.code || '').trim().toLowerCase() === ql) ||
        list.find((c) => String(c.name || '').trim().toLowerCase() === ql) ||
        null
    );
}

/**
 * Klassen für Schüler-Dropdown: Stammdaten plus ggf. Klassen aus bereits geladenen Anträgen.
 * @param {{ stammdaten?: { classes?: object[] }, items?: { klasse?: string }[] }} state
 */
function mergeClassPickerRows(rows) {
    const out = [];
    const seen = new Set();
    (rows || []).forEach((c) => {
        if (!c || typeof c !== 'object') return;
        const code = String(c.code || c.name || '').trim();
        if (!code || seen.has(code.toLowerCase())) return;
        seen.add(code.toLowerCase());
        out.push({
            code,
            name: String(c.name || c.code || code).trim()
        });
    });
    return out;
}

function dropLegacyKlasseRows(rows) {
    return (rows || []).filter((r) => !/^[1-5]AHW$/i.test(String((r && (r.code || r.name)) || '')));
}

export function classesForStudentPicker(state) {
    const perms = loadEffectivePermissionsConfig();
    // 1) Veröffentlichtes Klassenverzeichnis (aus Stammdaten: 1A, 1B, …)
    const fromCatalog = dropLegacyKlasseRows(mergeClassPickerRows(perms.classCatalog || []));
    if (fromCatalog.length) {
        return fromCatalog;
    }
    // 2) Lokale Stammdaten (IT-Browser)
    const fromStamm = dropLegacyKlasseRows(mergeClassPickerRows(collectAllClassRows(state || {})));
    if (fromStamm.length) {
        return fromStamm;
    }
    // 3) SharePoint-Spalte „Klasse“ – ohne alte Schema-Defaults 1AHW…
    const fromList = dropLegacyKlasseRows(
        mergeClassPickerRows(state && state.klasseColumnChoices ? state.klasseColumnChoices : [])
    );
    if (fromList.length) {
        return fromList;
    }
    const fromTeams = classRowsFromTeamLinks(
        (perms.classTeamLinks || []).concat(classTeamLinksFromAppData())
    );
    const merged = dropLegacyKlasseRows(mergeClassPickerRows(fromTeams));
    const seen = new Set(merged.map((c) => String(c.code).toLowerCase()));
    const extra = [];
    const pushCode = (k) => {
        const code = String(k || '').trim();
        if (!code || seen.has(code.toLowerCase()) || /^[1-5]AHW$/i.test(code)) return;
        seen.add(code.toLowerCase());
        extra.push({ code, name: code });
    };
    pushCode(resolveStudentKlasseCode(state));
    (state && state.items ? state.items : []).forEach((it) => {
        pushCode(it.klasse);
    });
    return merged.concat(extra);
}

export function loadStudentKlassePick(email) {
    const em = String(email || '').trim().toLowerCase();
    if (!em) return '';
    try {
        const raw = JSON.parse(localStorage.getItem(STUDENT_KLASSE_PICK_KEY) || '{}');
        return String((raw && raw[em]) || '').trim();
    } catch {
        return '';
    }
}

export function clearStudentKlassePick(email) {
    const em = String(email || '').trim().toLowerCase();
    if (!em) return;
    try {
        const raw = JSON.parse(localStorage.getItem(STUDENT_KLASSE_PICK_KEY) || '{}') || {};
        if (!raw[em]) return;
        delete raw[em];
        localStorage.setItem(STUDENT_KLASSE_PICK_KEY, JSON.stringify(raw));
    } catch {
        /* ignore */
    }
}

export function persistStudentKlassePick(email, klasseCode) {
    const em = String(email || '').trim().toLowerCase();
    const code = String(klasseCode || '').trim();
    if (!em || !code) return;
    try {
        const raw = JSON.parse(localStorage.getItem(STUDENT_KLASSE_PICK_KEY) || '{}') || {};
        raw[em] = code;
        localStorage.setItem(STUDENT_KLASSE_PICK_KEY, JSON.stringify(raw));
    } catch {
        /* ignore */
    }
}

/**
 * Gespeicherte Klassenwahl (einmal wählen, danach fix) wenn Stammdaten/Entra fehlen.
 * @param {object} state
 */
export function applyStudentKlasseFromLocalPick(state) {
    if (!state || state.role !== 'schueler') return false;
    if (studentKlasseFromRecord(state.studentMatch)) return true;
    const code = loadStudentKlassePick(state.accountEmail);
    if (!code) return false;
    state.studentMatch = {
        ...(state.studentMatch && typeof state.studentMatch === 'object' ? state.studentMatch : {}),
        email: state.accountEmail,
        name: state.accountName || '',
        klasse: code,
        klasseSource: 'local-pick'
    };
    state.demoKlasseCode = '';
    persistDemoKlasseCode('');
    prefillStudentFreistellungForm(state);
    return true;
}

function studentRecordEmails(s) {
    const out = [];
    ['email', 'mail', 'upn', 'userPrincipalName'].forEach((key) => {
        const v = String((s && s[key]) || '')
            .trim()
            .toLowerCase();
        if (v) out.push(v);
    });
    return out;
}

/**
 * @param {string} email
 * @param {{ klasse?: string, name?: string, email?: string }[]} students
 */
export function matchStudentByEmail(email, students) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return null;
    const list = Array.isArray(students) ? students : [];
    return list.find((s) => studentRecordEmails(s).includes(em)) || null;
}

/**
 * Klassenvorstand, wenn E-Mail in Stammdaten einer Klasse als headEmail hinterlegt ist.
 * @param {string} email
 * @param {object[]} classes
 */
export function matchKvByClassHeadEmail(email, classes) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return null;
    const list = Array.isArray(classes) ? classes : [];
    const hit = list.find((c) => {
        const head = String(c.headEmail || c.klassenvorstandEmail || c.kvEmail || '')
            .trim()
            .toLowerCase();
        return head && head === em;
    });
    if (!hit) return null;
    return {
        classCode: String(hit.code || hit.name || '').trim(),
        className: String(hit.name || hit.code || '').trim(),
        email: em
    };
}

/**
 * @param {{ studentMatch?: { klasse?: string }|null, demoKlasseCode?: string }} stateOrScope
 */
/** Schüler: Klasse nur bei Stammdaten/Entra sperren – nicht bei freier Auswahl im Formular. */
export function isStudentKlasseLocked(state) {
    if (!state || state.role !== 'schueler') return false;
    const sm = state.studentMatch;
    if (!sm || !studentKlasseFromRecord(sm)) return false;
    const src = String(sm.klasseSource || 'stammdaten');
    return (
        src === 'stammdaten' || src === 'entra-class-group' || src === 'entra-group-label'
    );
}

export function resolveStudentKlasseCode(stateOrScope) {
    const fromMatch =
        stateOrScope && stateOrScope.studentMatch
            ? studentKlasseFromRecord(stateOrScope.studentMatch)
            : '';
    if (fromMatch) return fromMatch;
    const demo = stateOrScope && stateOrScope.demoKlasseCode ? String(stateOrScope.demoKlasseCode).trim() : '';
    return demo;
}

const DEMO_KLASSE_KEY = 'ms365-freistellung-planer-demo-klasse-v1';

export function persistDemoKlasseCode(code) {
    try {
        localStorage.setItem(DEMO_KLASSE_KEY, String(code || '').trim());
    } catch {
        /* ignore */
    }
}

export function loadDemoKlasseCode() {
    try {
        return String(localStorage.getItem(DEMO_KLASSE_KEY) || '').trim();
    } catch {
        return '';
    }
}

/**
 * KV aus Stammdaten-Klasse ableiten.
 * @param {object[]} classes
 * @param {string} klasseCode
 */
/**
 * KV-E-Mail/Name aus gewählter Klasse in state.form setzen.
 * @param {object} state
 */
export function applyKvFromClass(state) {
    if (!state || !state.form) return;
    const kv = resolveKvForClass(state, state.form.klasse);
    if (kv) {
        state.form.kvEmail = kv.email;
        state.form.kvName = kv.name;
    }
}

/**
 * Schüler-Antrag: Name, Klasse und KV aus Stammdaten vorbefüllen.
 * @param {object} state
 */
export function prefillStudentFreistellungForm(state) {
    if (!state || state.role !== 'schueler') return;
    if (!state.form) state.form = emptyForm();
    if (state.accountName && !String(state.form.schuelerName || '').trim()) {
        state.form.schuelerName = state.accountName;
    }
    const fromStamm = studentKlasseFromRecord(state.studentMatch);
    const klasseCode = fromStamm || resolveStudentKlasseCode(state);
    if (fromStamm) {
        state.demoKlasseCode = '';
        persistDemoKlasseCode('');
    }
    if (klasseCode) {
        state.form.klasse = klasseCode;
        applyKvFromClass(state);
    }
}

function classRowHasKvEmail(row) {
    const email = String(
        (row && (row.headEmail || row.klassenvorstandEmail || row.kvEmail)) || ''
    )
        .trim()
        .toLowerCase();
    return email.includes('@');
}

function kvFromClassRow(row, teachers, klasseCode) {
    if (!row) return null;
    let email = String(row.headEmail || row.klassenvorstandEmail || row.kvEmail || '')
        .trim()
        .toLowerCase();
    let name = String(row.headName || row.klassenvorstandName || row.kvName || '').trim();
    if (!email.includes('@') && name && Array.isArray(teachers)) {
        const nl = name.toLowerCase();
        const t = teachers.find(
            (x) =>
                String(x.name || '').trim().toLowerCase() === nl ||
                String(x.code || '').trim().toLowerCase() === nl
        );
        if (t && t.email) email = String(t.email).trim().toLowerCase();
    }
    if (!email.includes('@')) return null;
    return { email, name, classCode: String(row.code || klasseCode || '').trim() };
}

/**
 * KV aus Klassen-Stammdaten, Schuljahr-Buckets und verknüpften Klassenteams.
 * @param {object} state
 * @param {string} klasseCode
 */
export function resolveKvForClass(state, klasseCode) {
    const code = String(klasseCode || '').trim();
    if (!code) return null;

    const sm = state && state.studentMatch;
    if (sm && (sm.kvEmail || sm.klassenvorstandEmail)) {
        const em = String(sm.kvEmail || sm.klassenvorstandEmail || '')
            .trim()
            .toLowerCase();
        if (em.includes('@')) {
            return {
                email: em,
                name: String(sm.kvName || sm.klassenvorstandName || '').trim(),
                classCode: code
            };
        }
    }

    const teachers = (state && state.stammdaten && state.stammdaten.teachers) || [];
    const allRows = collectAllClassRows(state || {});
    let row = findClassInStammdaten(allRows, code);

    if (!classRowHasKvEmail(row)) {
        const { classTeams, setup } = loadClassTeamsContext();
        const enriched = enrichClassesFromLinkedGroups([row || { code, name: code }], {
            classTeams,
            classGroupMatchByKey: setup.classGroupMatchByKey || {}
        });
        row = enriched.classes && enriched.classes[0] ? enriched.classes[0] : row;
    }

    return kvFromClassRow(row, teachers, code);
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
    const base = { accountEmail: state.accountEmail };
    const jahrgCodes =
        state.jahrgangScope && state.jahrgangScope.classCodes && state.jahrgangScope.classCodes.size
            ? state.jahrgangScope.classCodes
            : null;

    if (o.scopeAll || role === 'direktion') {
        return base;
    }
    if (role === 'kv') {
        const src = state.planerRoleSources && state.planerRoleSources.kv;
        const entraKv = src === 'entra' || src === 'global-admin' || src === 'kv-user';
        const scope = Object.assign({}, base);
        if (entraKv) scope.onlyKv = true;
        if (jahrgCodes) {
            scope.jahrgangClassCodes = jahrgCodes;
            if (scope.onlyKv) scope.jahrgangOrKv = true;
        }
        if (!scope.onlyKv && !scope.jahrgangClassCodes) scope.onlyKv = true;
        return scope;
    }
    return Object.assign({}, base, { onlyMine: true });
}

/**
 * @param {object} state
 * @param {object} item
 * @param {object} [opts] wie scopeFromState
 */
export function itemVisibleForRole(state, item, opts) {
    if (!item) return false;
    const scope = scopeFromState(state, opts);
    return filterFreistellungen([item], {}, scope).length > 0;
}

export function canDecide(state) {
    return state.role === 'kv' || state.role === 'direktion';
}

export function buildAntragTitle(form) {
    const name = String(form.schuelerName || form.titel || '').trim() || 'Freistellung';
    const klasse = String(form.klasse || '').trim();
    return klasse ? `${name} (${klasse})` : name;
}
