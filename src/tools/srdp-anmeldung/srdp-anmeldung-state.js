/**
 * State / Rollen / Views / Filter für sRDP-Anmeldung App.
 */
import { STORAGE_KEY, getProfile, listTitleForYear } from './srdp-anmeldung-schema.js';
import {
    filterAbschlussklassen,
    teacherChoiceLabels,
    parseChoiceLines,
    looksLikeDemoLfs,
    wizardConfiguredChoices,
    filterItemsForTeacher,
    filterItemsForKlassen,
    parseCodeList,
    joinCodeList,
    countNacharbeitPending,
    klassenOhneAnmeldung,
    preferredSrdpTerminJahr
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
    { id: 'termine', label: 'sRDP-Termine', icon: 'bi-calendar-event', roles: ['schueler', 'lehrer', 'admin'] },
    { id: 'meine-pruefungen', label: 'Meine Prüfungen', icon: 'bi-mortarboard', roles: ['lehrer'] },
    { id: 'meine-klasse', label: 'Meine Klasse', icon: 'bi-people', roles: ['lehrer'] },
    { id: 'liste', label: 'Alle Anmeldungen', icon: 'bi-list-ul', roles: ['lehrer', 'admin'] },
    { id: 'kalender', label: 'Kalender', icon: 'bi-calendar3', roles: ['schueler', 'lehrer', 'admin'] },
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
                items: [viewById('dashboard'), viewById('liste'), viewById('termine'), viewById('kalender')].filter(
                    Boolean
                )
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
            viewById('termine'),
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
            items: [viewById('anmelden'), viewById('meine'), viewById('termine'), viewById('kalender')].filter(
                Boolean
            )
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

export function roleLabel(role, source) {
    if (role === 'admin') {
        const s = String(source || '');
        if (s === 'designated-it' || s === 'global-admin') return 'Schul-IT';
        return 'Direktion';
    }
    if (role === 'lehrer') return 'Lehrkraft';
    return 'Schüler:in';
}

/** Jahr aus Listentitel „sRDP-Anmeldungen HAK 2027“. */
export function terminJahrFromListTitle(title) {
    const m = String(title || '').match(/\b(20\d{2})\b/);
    return m ? m[1] : '';
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

function studentKlasseFilled(row) {
    return !!String((row && (row.klasse || row.klasseCode || row.classCode)) || '').trim();
}

function studentMergeKey(row) {
    const em = String((row && (row.email || row.mail || row.upn)) || '')
        .trim()
        .toLowerCase();
    if (em) return em;
    return 'id:' + String((row && (row.id || row.name)) || '');
}

/** Zeile mit Klasse behalten, wenn dieselbe E-Mail in einem anderen Schuljahr ohne Klasse steht. */
function mergeStudentsPreferKlasse(lists) {
    const byKey = new Map();
    const order = [];
    (lists || []).forEach((list) => {
        (Array.isArray(list) ? list : []).forEach((row) => {
            if (!row || typeof row !== 'object') return;
            const key = studentMergeKey(row);
            if (!key || key === 'id:') return;
            const prev = byKey.get(key);
            if (!prev) {
                byKey.set(key, row);
                order.push(key);
                return;
            }
            if (!studentKlasseFilled(prev) && studentKlasseFilled(row)) byKey.set(key, row);
        });
    });
    return order.map((key) => byKey.get(key)).filter(Boolean);
}

/** Schüler aus jedem Schuljahr, nicht nur aus years.current. */
function studentsFromAllYearBuckets() {
    try {
        const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
        if (!api || typeof api.getContainer !== 'function') return [];
        const container = api.getContainer();
        const by = container && container.years && container.years.byLabel;
        if (!by || typeof by !== 'object') return [];
        const out = [];
        Object.keys(by).forEach((label) => {
            const bucket = by[label];
            (bucket && Array.isArray(bucket.students) ? bucket.students : []).forEach((row) => {
                out.push(row);
            });
        });
        return out;
    } catch {
        return [];
    }
}

/** Klassen → Microsoft-365-Gruppe (classTeams und classGroupMatchByKey). */
export function classGroupLinksFromAppData() {
    try {
        const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
        if (!api || typeof api.getContainer !== 'function') return [];
        const container = api.getContainer();
        const teams =
            typeof api.normalizeCoreClassTeams === 'function'
                ? api.normalizeCoreClassTeams((container.core && container.core.classTeams) || [])
                : (container.core && container.core.classTeams) || [];
        const out = [];
        (Array.isArray(teams) ? teams : []).forEach((t) => {
            if (!t) return;
            const code = String(t.classCode || t.code || '').trim();
            if (!code) return;
            out.push({
                code,
                groupId: String(t.graphGroupId || t.groupId || '').trim(),
                displayName: String(t.displayName || '').trim(),
                mail: String(t.mail || t.gruppenmail || '')
                    .trim()
                    .toLowerCase()
            });
        });
        const setup = typeof api.getSetup === 'function' ? api.getSetup() : null;
        const map = (setup && setup.classGroupMatchByKey) || {};
        Object.keys(map).forEach((code) => {
            const entry = map[code];
            if (!entry || typeof entry !== 'object') return;
            out.push({
                code: String(code || '').trim(),
                groupId: String(entry.groupId || entry.graphGroupId || '').trim(),
                displayName: String(entry.displayName || entry.name || '').trim(),
                mail: String(entry.mail || entry.email || '')
                    .trim()
                    .toLowerCase()
            });
        });
        return out.filter((row) => row.code);
    } catch {
        return [];
    }
}

export function loadStammdaten() {
    try {
        const core = window.ms365TenantSettingsLoad && window.ms365TenantSettingsLoad();
        const data = (core && core.data) || core || {};
        const currentStudents = Array.isArray(data.students) ? data.students : [];
        return {
            classes: Array.isArray(data.classes) ? data.classes : [],
            teachers: Array.isArray(data.teachers) ? data.teachers : [],
            subjects: Array.isArray(data.subjects) ? data.subjects : [],
            students: mergeStudentsPreferKlasse([studentsFromAllYearBuckets(), currentStudents])
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
    const hits = list.filter((s) => {
        const mail = String((s && (s.email || s.mail)) || '')
            .trim()
            .toLowerCase();
        const upn = String((s && (s.upn || s.userPrincipalName)) || '')
            .trim()
            .toLowerCase();
        return (mail && mail === em) || (upn && upn === em);
    });
    if (!hits.length) return null;
    return hits.find((s) => studentKlasseFilled(s)) || hits[0];
}

/**
 * Demo-Klassen (5AHAK …) vs. echte Abschlussklassen aus Stammdaten/Liste erkennen.
 * @param {string[]} klassen
 */
export function looksLikeDemoKlassen(klassen) {
    const list = (Array.isArray(klassen) ? klassen : [])
        .map((k) => String(k || '').trim())
        .filter(Boolean);
    if (!list.length) return false;
    // Demo-Welt: 5AHAK / 5BHAK / 5CHAK – echte Schulen nutzen i. d. R. 5A / 5B o. Ä.
    return list.every((k) => /^5[A-C]HAK$/i.test(k));
}

/** Demo-Wizard: nur DIGBIZ und Wirtschaftsrecht. */
export function looksLikeDemoSeminare(value) {
    const list = (Array.isArray(value) ? value : parseChoiceLines(value))
        .map((s) => String(s || '').trim().toLowerCase())
        .filter(Boolean)
        .sort();
    return list.length === 2 && list[0] === 'digbiz' && list[1] === 'wirtschaftsrecht';
}

/**
 * Optionen aus Wizard + Profil.
 * Live-Stammdaten/Listen-Choices schlagen veraltete Demo-Wizard-Klassen.
 */
export function resolveFormOptions(wizard, stammdaten, profileId, opts) {
    const wiz = wizard || {};
    const allowDemoKlassen = !!(opts && opts.allowDemoKlassen);
    const profile = getProfile(profileId || wiz.profileId || 'hak');
    const jahr = preferredSrdpTerminJahr(new Date(), wiz.jahr || wiz.terminJahr);
    const klassenFromSd = filterAbschlussklassen(stammdaten && stammdaten.classes, jahr);
    const klassenWizRaw = Array.isArray(wiz.klassen)
        ? wiz.klassen.map((k) => String(k || '').trim()).filter(Boolean)
        : [];
    const klassenWiz =
        allowDemoKlassen || !looksLikeDemoKlassen(klassenWizRaw) ? klassenWizRaw : [];
    const klassenFromList = Array.isArray(wiz.listKlassen)
        ? wiz.listKlassen.map((k) => String(k || '').trim()).filter(Boolean)
        : [];
    let klassen;
    if (klassenFromSd.length && (!klassenWiz.length || looksLikeDemoKlassen(klassenWiz))) {
        klassen = klassenFromSd.slice();
    } else if (klassenFromList.length && (!klassenWiz.length || looksLikeDemoKlassen(klassenWiz))) {
        klassen = klassenFromList.slice();
    } else if (klassenWiz.length) {
        klassen = klassenWiz.slice();
    } else {
        klassen = klassenFromSd.length ? klassenFromSd.slice() : klassenFromList.slice();
    }
    klassen = [...new Set(klassen)];
    const lehrerFromStamm = teacherChoiceLabels((stammdaten && stammdaten.teachers) || []);
    const lehrerFromList = (Array.isArray(wiz.listLehrer) ? wiz.listLehrer : [])
        .map((k) => String(k || '').trim())
        .filter((k) => k && k !== '—' && k !== '–');
    const lehrer = lehrerFromStamm.length
        ? lehrerFromStamm
        : [...new Set(lehrerFromList)].sort((a, b) => a.localeCompare(b, 'de'));
    const listLfs = (Array.isArray(wiz.listLfs) ? wiz.listLfs : [])
        .map((k) => String(k || '').trim())
        .filter((k) => k && k !== '—' && k !== '–');
    const listWahl = (Array.isArray(wiz.listWahlfaecher) ? wiz.listWahlfaecher : [])
        .map((k) => String(k || '').trim())
        .filter((k) => k && k !== '—' && k !== '–' && !/bitte wählen/i.test(k));
    const seminarFromList = (Array.isArray(wiz.listSeminare) ? wiz.listSeminare : [])
        .map((k) => String(k || '').trim())
        .filter((k) => k && k !== '—' && k !== '–' && !/bitte wählen/i.test(k));
    let lfsLines = wizardConfiguredChoices(wiz, 'lfsText', 'lfs');
    let wahlLines = wizardConfiguredChoices(wiz, 'wahlfachText', 'wahlfach');
    let seminarLines = wizardConfiguredChoices(wiz, 'seminarText', 'seminar');
    if (!allowDemoKlassen && looksLikeDemoLfs(lfsLines)) lfsLines = [];
    if (!allowDemoKlassen && looksLikeDemoSeminare(seminarLines)) seminarLines = [];
    const lfs = lfsLines.length
        ? lfsLines
        : listLfs.length
          ? listLfs
          : allowDemoKlassen
            ? (profile.lfsHints || []).slice()
            : [];
    const wahlfaecher = wahlLines.length
        ? wahlLines
        : listWahl.length
          ? listWahl
          : allowDemoKlassen
            ? (profile.defaultWahlfaecher || []).slice()
            : [];
    const seminare = seminarLines.length
        ? seminarLines
        : seminarFromList.length
          ? seminarFromList
          : Array.isArray(wiz.seminare) && wiz.seminare.length
            ? wiz.seminare
            : allowDemoKlassen
              ? (profile.defaultSeminare || []).slice()
              : [];
    return {
        profileId: profile.id,
        profileLabel: profile.label || profile.id,
        jahr,
        listTitle: listTitleForYear(jahr, profile.id),
        klassen,
        lehrer,
        lfs,
        lfsFieldLabel: profile.lfsFieldLabel || 'LFS',
        wahlfaecher,
        seminare,
        variants: profile.variants || [],
        anmeldeFrist: String(wiz.anmeldeFrist || '').trim(),
        haupttermin: String(wiz.haupttermin || '').trim(),
        bfkDatum: String(wiz.bfkDatum || '').match(/20\d{2}-\d{2}-\d{2}/)
            ? String(wiz.bfkDatum).match(/20\d{2}-\d{2}-\d{2}/)[0]
            : ''
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
        hatABA: s.hatABA === true || s.hatABA === 'true' || s.hatABA === 1,
        titelABA: String(s.titelABA || '').trim(),
        betreuungslehrerABA: String(s.betreuungslehrerABA || '').trim(),
        abaPrefill: !!s.abaPrefill,
        variante: String(s.variante || '').trim(),
        lfs: String(s.lfs || '').trim(),
        lehrerLFS: String(s.lehrerLFS || '').trim(),
        lehrerBKO: String(s.lehrerBKO || '').trim(),
        lehrerAM: String(s.lehrerAM || '').trim(),
        wahlfachMuendlich: String(s.wahlfachMuendlich || s.wahlfach || '').trim(),
        seminar: String(s.seminar || '').trim(),
        lehrerWahlfach: String(s.lehrerWahlfach || '').trim(),
        schwerpunktFach: String(s.schwerpunktFach || '').trim(),
        lehrerSchwerpunkt: String(s.lehrerSchwerpunkt || '').trim(),
        fachkolloquium: String(s.fachkolloquium || '').trim(),
        lehrerFachkolloquium: String(s.lehrerFachkolloquium || '').trim(),
        fachtheorie: String(s.fachtheorie || '').trim(),
        lehrerFachtheorie: String(s.lehrerFachtheorie || '').trim(),
        klausurKombi: String(s.klausurKombi || '').trim(),
        muendlichExtra: String(s.muendlichExtra || '').trim(),
        lehrerMuendlichExtra: String(s.lehrerMuendlichExtra || '').trim(),
        vierteKlausur: String(s.vierteKlausur || '').trim(),
        muendlich1: String(s.muendlich1 || '').trim(),
        muendlich2: String(s.muendlich2 || '').trim(),
        muendlich3: String(s.muendlich3 || '').trim(),
        lehrerMuendlich1: String(s.lehrerMuendlich1 || '').trim(),
        lehrerMuendlich2: String(s.lehrerMuendlich2 || '').trim(),
        lehrerMuendlich3: String(s.lehrerMuendlich3 || '').trim(),
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
        planerRoleSources: {},
        planerAccessDenied: false,
        roleHint: '',
        accountEmail: '',
        accountName: '',
        memberGroups: [],
        demoPersonaEmail: '',
        demoPersonaName: '',
        siteUrl: loadSavedSiteUrl(),
        listId: String(wizard.listId || '').trim(),
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
        /** @type {object[]} ABA-Portal Import (explodiert pro Schüler/in) */
        abschlussarbeiten: [],
        /** @type {object[]} aktuelle Treffer für Prefill-Auswahl */
        abschlussarbeitHits: [],
        abschlussarbeitenMeta: { listTitle: '', count: 0, lastImportAt: '' },
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
        groupBy: 'klasse',
        filterMultiOpen: /** @type {Record<string, boolean>} */ ({}),
        filterMultiQuery: /** @type {Record<string, string>} */ ({}),
        detailItemId: '',
        teacherLabel: '',
        kvKlassen: /** @type {string[]} */ ([]),
        calMode: 'month',
        calYear: today.getFullYear(),
        calMonth: today.getMonth() + 1,
        calWeekAnchor: today.toISOString().slice(0, 10),
        systemTab: 'site'
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
