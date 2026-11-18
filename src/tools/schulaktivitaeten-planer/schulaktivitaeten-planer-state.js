/**
 * State / Filter / Rollen für Schulaktivitäten-Planer.
 */
import {
    toIsoDateOnly,
    joinCodeList,
    klasseCodesOf,
    parseCodeList,
    initialStatusForSubmit
} from './schulaktivitaeten-planer-logic.js';
import { TYP_CHOICES, TEILNAHME_MODUS_CHOICES } from './schulaktivitaeten-planer-schema.js';
import {
    resolveSchoolToolSiteUrl,
    patchSchoolToolSitesSetup
} from '../../shared/school-tool-sites.js';

const ROLE_KEY = 'ms365-akt-planer-role-v1';
const SITE_KEY = 'ms365-akt-planer-site-v1';

/** @typedef {'lehrer'|'kv'|'admin'} AktRole */

export const VIEWS = [
    { id: 'dashboard', label: 'Übersicht', icon: 'bi-house', roles: ['lehrer', 'kv', 'admin'] },
    { id: 'liste', label: 'Meine Aktivitäten', icon: 'bi-list-ul', roles: ['lehrer', 'kv', 'admin'] },
    { id: 'kalender', label: 'Kalender', icon: 'bi-calendar3', roles: ['lehrer', 'kv', 'admin'] },
    { id: 'meine-klasse', label: 'Meine Klasse', icon: 'bi-people', roles: ['kv', 'admin'] },
    { id: 'limits', label: 'Kostenschwellen', icon: 'bi-cash-coin', roles: ['admin'] },
    { id: 'berichte', label: 'Berichte', icon: 'bi-bar-chart', roles: ['admin'] },
    { id: 'antrag', label: 'Neue Aktivität', icon: 'bi-plus-lg', roles: ['lehrer', 'kv', 'admin'] },
    { id: 'freigabe', label: 'Entscheidungen', icon: 'bi-check2-square', roles: ['admin'] },
    { id: 'regeln', label: 'Regelwerk', icon: 'bi-sliders', roles: ['admin'] },
    { id: 'system', label: 'Systemadministration', icon: 'bi-gear', roles: ['admin'] }
];

/** @param {string} id */
function viewById(id) {
    return VIEWS.find((v) => v.id === id) || null;
}

/** Setup-/Site-Toolbar nur für Direktion/Admin – KV bleibt in der Lehrerwelt. */
export function isAktStaffRole(role) {
    return role === 'admin';
}

/**
 * Offene Direktions-Entscheidungen (beantragt).
 * @param {object[]} items
 */
export function countOpenEntscheidungen(items) {
    return (Array.isArray(items) ? items : []).filter(
        (i) => String(i.status || '').toLowerCase() === 'beantragt'
    ).length;
}

/**
 * Genehmigt, Untis noch offen.
 * @param {object[]} items
 */
export function countUntisPending(items) {
    return (Array.isArray(items) ? items : []).filter(
        (i) => String(i.status || '').toLowerCase() === 'genehmigt' && !i.untisErledigt
    ).length;
}

/**
 * Gruppierte Navigation nach Rolle.
 * - Lehrer: persönliche Arbeit (+ Meine Klasse bei KV)
 * - Admin/Direktion: Arbeit · Auswertung · Regeln · System
 *
 * @param {string} role
 * @param {string[]|null|undefined} [availableRoles]
 * @param {{ kvKlassen?: string[] }|null} [opts]
 * @returns {{ id: string, label: string, items: typeof VIEWS }[]}
 */
export function navSectionsForRole(role, availableRoles, opts) {
    const r = role === 'admin' ? 'admin' : role === 'kv' ? 'kv' : 'lehrer';

    if (r === 'admin') {
        const listeAll = { ...viewById('liste'), label: 'Alle Aktivitäten' };
        const entscheidungen = { ...viewById('freigabe'), label: 'Entscheidungen' };
        const kosten = { ...viewById('limits'), label: 'Kostenschwellen' };
        return [
            {
                id: 'arbeit',
                label: 'Arbeit',
                items: [viewById('dashboard'), entscheidungen, listeAll, viewById('kalender')].filter(
                    Boolean
                )
            },
            {
                id: 'auswertung',
                label: 'Auswertung',
                items: [viewById('berichte')].filter(Boolean)
            },
            {
                id: 'regeln',
                label: 'Regeln',
                items: [viewById('regeln'), kosten].filter(Boolean)
            },
            {
                id: 'system',
                label: 'System',
                items: [viewById('system')].filter(Boolean)
            }
        ];
    }

    const avail = Array.isArray(availableRoles)
        ? availableRoles.map((x) => String(x || '').toLowerCase())
        : [];
    const kvKlassen = Array.isArray(opts && opts.kvKlassen) ? opts.kvKlassen : [];
    const showMeineKlasse = r === 'kv' || avail.includes('kv') || kvKlassen.length > 0;
    const items = [
        viewById('dashboard'),
        viewById('liste'),
        viewById('kalender'),
        ...(showMeineKlasse ? [viewById('meine-klasse')] : []),
        viewById('antrag')
    ].filter(Boolean);
    return [{ id: 'arbeit', label: '', items }];
}

/**
 * Flache View-Liste (für Allowed-Checks / Tests).
 * @param {string} role
 * @param {string[]|null|undefined} [availableRoles]
 * @param {{ kvKlassen?: string[] }|null} [opts]
 */
export function viewsForRole(role, availableRoles, opts) {
    return navSectionsForRole(role, availableRoles, opts).flatMap((s) => s.items || []);
}

/** Klassenvorstand-Fähigkeit über Stammdaten/Rollen – nicht als App-Modus. */
export function hasAktKvCapability(state) {
    if (!state) return false;
    if (state.role === 'admin' && !state.demoPersonaEmail) return true;
    if (kvKlasseCodesForState(state).length) return true;
    if (state.role === 'kv') return true;
    return (state.planerRoles || []).map((r) => String(r || '').toLowerCase()).includes('kv');
}

/**
 * Persönliche Beziehungen zur Aktivität (mehrere gleichzeitig möglich).
 * @param {object} item
 * @param {{ email?: string, teacherCode?: string }} who
 * @returns {Array<{ type: 'applicant'|'responsible'|'companion', label: string }>}
 */
export function resolveUserRelations(item, who) {
    if (!item) return [];
    const email = String((who && who.email) || '')
        .trim()
        .toLowerCase();
    const code = String((who && who.teacherCode) || '')
        .trim()
        .toUpperCase();
    /** @type {Array<{ type: 'applicant'|'responsible'|'companion', label: string }>} */
    const rels = [];

    if (email) {
        const von = String(item.beantragtVon || item.authorEmail || '')
            .trim()
            .toLowerCase();
        if (von && von === email) {
            rels.push({ type: 'applicant', label: 'Von dir eingereicht' });
        }
    }

    const em = String(item.lehrerEmail || '')
        .trim()
        .toLowerCase();
    const ownCode = String(item.lehrerCode || '')
        .trim()
        .toUpperCase();
    const isResponsible =
        (email && em && em === email) || (code && ownCode && ownCode === code);
    if (isResponsible) {
        rels.push({ type: 'responsible', label: 'Du bist verantwortlich' });
    }

    // Fallback: ohne beantragtVon gilt Verantwortliche / Autor oft als Antragsteller
    if (!rels.some((r) => r.type === 'applicant')) {
        const von = String(item.beantragtVon || '')
            .trim()
            .toLowerCase();
        const author = String(item.authorEmail || '')
            .trim()
            .toLowerCase();
        if (isResponsible && !von) {
            rels.unshift({ type: 'applicant', label: 'Von dir eingereicht' });
        } else if (email && author && author === email && !von) {
            rels.unshift({ type: 'applicant', label: 'Von dir eingereicht' });
        }
    }

    if (code) {
        const begleitung = parseCodeList(item.begleitungCodes || item.begleitung).map((c) =>
            String(c).trim().toUpperCase()
        );
        if (begleitung.includes(code)) {
            rels.push({ type: 'companion', label: 'Du begleitest' });
        }
    }

    // Dedup by type
    const seen = new Set();
    return rels.filter((r) => {
        if (seen.has(r.type)) return false;
        seen.add(r.type);
        return true;
    });
}

/** @deprecated Kompatibilität – primäre Beziehung */
export function userInvolvement(item, who) {
    const rels = resolveUserRelations(item, who);
    if (rels.some((r) => r.type === 'companion') && !rels.some((r) => r.type === 'applicant' || r.type === 'responsible')) {
        return 'begleitung';
    }
    if (rels.length) return 'organizer';
    return null;
}

export function involvementLabel(kind) {
    if (kind === 'organizer') return 'Von dir eingereicht';
    if (kind === 'begleitung') return 'Du begleitest';
    return '';
}

export function formatRelationSummary(rels) {
    return (rels || []).map((r) => r.label).join(' · ');
}

export function hasPersonalInvolvement(item, who) {
    return resolveUserRelations(item, who).length > 0;
}

/** Aktivität betrifft eine der KV-Klassen (auch Teilgruppe). */
export function itemTouchesKvKlassen(item, kvKlassen) {
    const wanted = (kvKlassen || []).map((c) => String(c).trim().toUpperCase()).filter(Boolean);
    if (!wanted.length) return false;
    return klasseCodesOf(item).some((c) => wanted.includes(String(c).trim().toUpperCase()));
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
        /** Live-Hydrate-Meta (Klassen/Lehrer aus SharePoint) */
        stammdatenLive: {
            ok: false,
            error: '',
            classCount: 0,
            teacherCount: 0,
            missing: /** @type {string[]} */ ([]),
            source: ''
        },
        /** Pull-Meta Aktivitäten-Liste (graph / spo-rest / demo) */
        aktivitaetenLive: {
            ok: false,
            count: 0,
            source: '',
            error: '',
            pulledAt: ''
        },
        filters: { klasse: '', typ: '', status: '', lehrer: '', beteiligung: '', verarbeitung: '' },
        /** Mehrfachfilter-Dropdowns offen halten über paint() */
        filterMultiOpen: { klasse: false, typ: false, status: false, lehrer: false },
        /** Suchtext in Mehrfachfiltern über paint() behalten */
        filterMultiQuery: { klasse: '', typ: '', status: '', lehrer: '' },
        /** Auswahl in „Meine Klasse“ – getrennt vom Listen-/Kalenderfilter */
        meineKlasseCode: '',
        form: emptyForm(),
        editingItemId: null,
        detailId: null,
        calYear: new Date().getFullYear(),
        calMonth: new Date().getMonth() + 1,
        /** @type {'agenda'|'week'|'month'} */
        calMode: 'month',
        /** Montag der angezeigten Woche */
        calWeekAnchor: (() => {
            const d = new Date();
            d.setHours(0, 0, 0, 0);
            const day = (d.getDay() + 6) % 7;
            d.setDate(d.getDate() - day);
            return d.toISOString();
        })(),
        /** Startjahr des Schuljahrs (2026 → 2026/27) */
        reportYear: (() => {
            const n = new Date();
            return n.getMonth() + 1 >= 9 ? n.getFullYear() : n.getFullYear() - 1;
        })(),
        /** Nach paint: #aktUntisQueue anscrollen (legacy) */
        focusUntisQueue: false,
        /** Formular-Dirty für Regeln / Kostenschwellen */
        rulesDirty: false,
        limitsDirty: false,
        /** Kurzfeedback nach Speichern */
        settingsSavedFlash: '',
        ctx: null,
        localDemoOnly: false,
        feedMeta: null,
        feedBusy: false,
        demoRoleOverride: false,
        /** Lokale Demo: Filter/KV als diese Persona, obwohl IT angemeldet bleibt */
        demoPersonaEmail: '',
        demoPersonaName: '',
        entraGroupsConfigured: false,
        planerRoles: /** @type {import('./schulaktivitaeten-planer-entra-role.js').AktRole[]} */ ([]),
        planerRoleSources: /** @type {Record<string, string>} */ ({}),
        planerAccessDenied: false,
        /** @type {1|2|3|4} Guided Flow Schritt (Neue Aktivität) */
        formStep: 1,
        /** Validierung erst nach Blur/Weiter/Submit */
        formAttempted: false,
        /** @type {Record<string, boolean>} */
        formTouched: {},
        /** @type {'ue'|'clock'} Zeitmodell Organisation */
        timeMode: 'ue',
        begleitungPickerOpen: false,
        weitereAngabenOpen: false,
        /** Mehrtägige Aktivität (Schritt 1 Progressive Disclosure) */
        multiDay: false,
        /** Nach erfolgreichem Einreichen */
        submitSuccess: /** @type {null|{ id: string, titel: string }} */ (null)
    };
}

export function emptyForm(partial) {
    const t = new Date();
    const iso = toIsoDateOnly(t) || '';
    const base = {
        aktivitaetId: '',
        titel: '',
        typ: 'Exkursion',
        klasseCode: '',
        klasseCodes: '',
        teilnahmeModus: 'klasse',
        gruppenText: '',
        teilnehmendeText: '',
        nichtTeilnehmende: '',
        lehrerCode: '',
        lehrerEmail: '',
        begleitung: '',
        begleitungCodes: '',
        externeVortragende: '',
        ort: '',
        ablauf: '',
        unterrichtsbezug: '',
        startdatum: iso,
        enddatum: iso,
        startUe: '',
        endUe: '',
        startZeit: '',
        endZeit: '',
        status: 'entwurf',
        notiz: '',
        ablehnungsGrund: '',
        rueckfrageText: '',
        aenderungsLog: '',
        verkehrsmittel: '',
        kostenHinweis: '',
        untisErledigt: false,
        untisErledigtVon: '',
        untisErledigtAm: ''
    };
    const merged = { ...base, ...(partial || {}) };
    const codes = klasseCodesOf(merged);
    merged.klasseCodes = joinCodeList(codes);
    merged.klasseCode = codes[0] || '';
    if (!TEILNAHME_MODUS_CHOICES.includes(String(merged.teilnahmeModus || '').toLowerCase())) {
        merged.teilnahmeModus = codes.length > 1 ? 'klassen' : 'klasse';
    }
    merged.begleitungCodes = joinCodeList(merged.begleitungCodes || parseCodeList(merged.begleitung));
    return merged;
}

export function resolveRole(preferred) {
    if (preferred === 'admin' || preferred === 'kv' || preferred === 'lehrer') return preferred;
    try {
        const s = String(localStorage.getItem(ROLE_KEY) || '').toLowerCase();
        if (s === 'admin' || s === 'kv' || s === 'lehrer') return s;
    } catch {
        /* ignore */
    }
    return 'lehrer';
}

export function persistRole(role) {
    try {
        const r = role === 'admin' ? 'admin' : role === 'kv' ? 'kv' : 'lehrer';
        localStorage.setItem(ROLE_KEY, r);
    } catch {
        /* ignore */
    }
}

export function loadSavedSiteUrl() {
    return resolveSchoolToolSiteUrl('aktivitaeten') || '';
}

/**
 * Site-URL lokal + schulweit im Setup merken (damit IT-Sicherung / andere Browser sie finden).
 * @param {string} url
 */
export function persistSiteUrl(url) {
    const u = String(url || '').trim();
    if (!u) return;
    try {
        localStorage.setItem(SITE_KEY, u);
    } catch {
        /* ignore */
    }
    try {
        patchSchoolToolSitesSetup({
            planerDefaultSiteUrl: u,
            planerSiteUrls: { aktivitaeten: u },
            applyLocalEmpty: true
        });
    } catch {
        /* ignore */
    }
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.patchSetup === 'function'
        ) {
            const setup =
                typeof window.ms365AppDataV2.getSetup === 'function'
                    ? window.ms365AppDataV2.getSetup() || {}
                    : {};
            /** @type {Record<string, string>} */
            const patch = { planerDefaultSiteUrl: u };
            if (!String(setup.intranetSiteUrl || '').trim()) patch.intranetSiteUrl = u;
            window.ms365AppDataV2.patchSetup(patch);
        }
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
            students: Array.isArray(data.students) ? data.students : []
        };
    } catch {
        return { classes: [], teachers: [], students: [] };
    }
}

/**
 * Schüler:innen-Anzahl für eine Klasse (falls Stammdaten vorhanden).
 * @param {object} stammdaten
 * @param {string} klasseCode
 */
export function studentCountForKlasse(stammdaten, klasseCode) {
    const code = String(klasseCode || '').trim();
    if (!code) return 0;
    const students = (stammdaten && stammdaten.students) || [];
    if (!students.length) {
        const cls = ((stammdaten && stammdaten.classes) || []).find(
            (c) => String(c.code || '').trim() === code
        );
        const n = Number(cls && (cls.studentCount || cls.schuelerAnzahl || cls.size));
        return Number.isFinite(n) && n > 0 ? n : 0;
    }
    return students.filter((s) => {
        const k = String(s.klasseCode || s.classCode || s.klasse || s.class || '').trim();
        return k === code;
    }).length;
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

/** Effektive E-Mail für Meine Aktivitäten / Meine Klasse (inkl. lokaler Demo-Persona). */
export function effectiveAccountEmail(state) {
    return String((state && (state.demoPersonaEmail || state.accountEmail)) || '')
        .trim()
        .toLowerCase();
}

export function effectiveTeacherCode(state) {
    const email = effectiveAccountEmail(state);
    const t = matchTeacherByEmail((state && state.stammdaten && state.stammdaten.teachers) || [], email);
    return t && t.code ? String(t.code) : '';
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
    if (s === 'entwurf') return 'Entwurf';
    if (s === 'beantragt') return 'Eingereicht';
    if (s === 'kv_stellungnahme') return 'In Prüfung';
    if (s === 'nachforderung') return 'Ergänzung erforderlich';
    if (s === 'genehmigt') return 'Genehmigt';
    if (s === 'abgelehnt') return 'Abgelehnt';
    if (s === 'abgesagt') return 'Abgesagt';
    return 'Eingereicht';
}

/**
 * Organisatorischer Verarbeitungsstatus (getrennt vom Genehmigungsstatus).
 * @param {object} item
 * @returns {'not_required'|'pending'|'completed'}
 */
export function scheduleStatusOf(item) {
    if (!item) return 'not_required';
    const st = String(item.status || '').toLowerCase();
    if (st !== 'genehmigt') return 'not_required';
    return item.untisErledigt ? 'completed' : 'pending';
}

/** @param {'not_required'|'pending'|'completed'|string} code */
export function scheduleStatusLabel(code) {
    const c = String(code || '').toLowerCase();
    if (c === 'pending') return 'Eintragung ausständig';
    if (c === 'completed') return 'In Untis verarbeitet';
    return '';
}

/** Statusoptionen für Filter (Code + Label). */
export function statusFilterChoices() {
    return [
        { code: 'entwurf', name: statusLabel('entwurf') },
        { code: 'beantragt', name: statusLabel('beantragt') },
        { code: 'kv_stellungnahme', name: statusLabel('kv_stellungnahme') },
        { code: 'nachforderung', name: statusLabel('nachforderung') },
        { code: 'genehmigt', name: statusLabel('genehmigt') },
        { code: 'abgelehnt', name: statusLabel('abgelehnt') },
        { code: 'abgesagt', name: statusLabel('abgesagt') }
    ];
}

export function teilnahmeModusLabel(modus) {
    const m = String(modus || '').toLowerCase();
    if (m === 'klassen') return 'Mehrere Klassen';
    if (m === 'gruppe') return 'Unterrichtsgruppe';
    if (m === 'schueler') return 'Einzelne Schüler:innen';
    return 'Ganze Klasse';
}

export function teilnahmeModusChoices() {
    return TEILNAHME_MODUS_CHOICES.map((c) => ({ code: c, name: teilnahmeModusLabel(c) }));
}

/**
 * Anzeige-Text für Klassen / Gruppen in Listen.
 * @param {object} item
 * @param {Record<string,string>} [klasseMap]
 */
export function teilnahmeKurz(item, klasseMap) {
    const map = klasseMap || {};
    const codes = klasseCodesOf(item);
    const labels = codes.map((c) => map[c] || c);
    const modus = String((item && item.teilnahmeModus) || '').toLowerCase();
    if (modus === 'gruppe' && item.gruppenText) {
        return item.gruppenText + (labels.length ? ` (${labels.join(', ')})` : '');
    }
    if (modus === 'schueler') {
        return (labels.length ? labels.join(', ') + ' · ' : '') + 'Einzelpersonen';
    }
    return labels.join(', ') || (item && item.klasseCode) || '–';
}

export function typLabel(typ) {
    const t = String(typ || '');
    if (t === 'Schulaktivitaet') return 'Schulaktivität';
    if (t === 'Klassenabwesenheit') return 'Klassenabwesenheit';
    return t || '–';
}

export function roleLabel(role) {
    if (role === 'admin') return 'Admin / Direktion';
    if (role === 'kv') return 'Klassenvorstand';
    return 'Lehrkraft';
}

/**
 * Klassen-Codes, für die der aktuelle Account Klassenvorstand ist (headEmail).
 * @param {object} stammdaten
 * @param {string} email
 * @returns {string[]}
 */
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
        .map((c) => String(c.code || '').trim())
        .filter(Boolean);
}

/** KV-Klassen der effektiven Persona (Demo oder Login). */
export function kvKlasseCodesForState(state) {
    return kvKlasseCodesForEmail(state && state.stammdaten, effectiveAccountEmail(state));
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
 * @param {{ role: string, accountEmail?: string, onlyMine?: boolean, onlyOpen?: boolean, onlyApprovedPublic?: boolean, kvKlassen?: string[] }} scope
 */
export function filterItems(items, filters, scope) {
    const f = filters || {};
    const list = Array.isArray(items) ? items : [];
    const email = String((scope && scope.accountEmail) || '')
        .trim()
        .toLowerCase();
    const onlyMine = !!(scope && scope.onlyMine);
    const onlyOpen = !!(scope && scope.onlyOpen);
    const onlyApprovedPublic = !!(scope && scope.onlyApprovedPublic);
    const kvKlassen = Array.isArray(scope && scope.kvKlassen) ? scope.kvKlassen : [];
    return list.filter((it) => {
        const st = String(it.status || '').toLowerCase();
        if (onlyOpen) {
            if (st !== 'beantragt' && st !== 'kv_stellungnahme') return false;
        }
        if (st === 'abgesagt' && onlyApprovedPublic) return false;
        const teacherCode = String((scope && scope.teacherCode) || '').trim();
        const who = { email, teacherCode };
        const rels = resolveUserRelations(it, who);
        if (onlyMine) {
            if (!rels.length) return false;
        }
        if (onlyApprovedPublic) {
            if (!rels.length && st !== 'genehmigt') return false;
        }
        if (kvKlassen.length) {
            if (!itemTouchesKvKlassen(it, kvKlassen)) return false;
        }
        const wantKlassen = parseCodeList(f.klasse).map((c) => String(c).toUpperCase());
        if (wantKlassen.length) {
            const codes = klasseCodesOf(it).map((c) => String(c).toUpperCase());
            if (!wantKlassen.some((w) => codes.includes(w))) return false;
        }
        const wantTypen = parseCodeList(f.typ);
        if (wantTypen.length && !wantTypen.includes(String(it.typ || ''))) return false;
        const wantStatus = parseCodeList(f.status).map((s) => String(s).toLowerCase());
        if (wantStatus.length && !wantStatus.includes(st)) return false;
        const wantLehrer = parseCodeList(f.lehrer).map((c) => String(c).toUpperCase());
        if (
            wantLehrer.length &&
            !wantLehrer.includes(String(it.lehrerCode || '').trim().toUpperCase())
        ) {
            return false;
        }
        if (f.beteiligung) {
            const want = String(f.beteiligung).toLowerCase();
            if (!rels.some((r) => r.type === want)) return false;
        }
        if (f.verarbeitung) {
            const want = String(f.verarbeitung).toLowerCase();
            const sched = scheduleStatusOf(it);
            if (want === 'offen' && sched !== 'pending') return false;
            if (want === 'erledigt' && sched !== 'completed') return false;
        }
        return true;
    });
}

/**
 * Chronologisch nach Startdatum (Admin-Sammelliste / Untis).
 * @param {object[]} items
 */
export function sortByStartdatum(items) {
    return (Array.isArray(items) ? items.slice() : []).sort((a, b) => {
        const sa = toIsoDateOnly(a && a.startdatum) || '';
        const sb = toIsoDateOnly(b && b.startdatum) || '';
        if (sa < sb) return -1;
        if (sa > sb) return 1;
        return String((a && a.titel) || '').localeCompare(String((b && b.titel) || ''), 'de');
    });
}

export function canEditItem(item, state) {
    if (!item) return false;
    if (state.role === 'admin' && !state.demoPersonaEmail) return true;
    const st = String(item.status || '').toLowerCase();
    if (st !== 'entwurf' && st !== 'beantragt' && st !== 'nachforderung' && st !== 'kv_stellungnahme')
        return false;
    const email = effectiveAccountEmail(state);
    if (!email) return true;
    return (
        String(item.lehrerEmail || '').toLowerCase() === email ||
        String(item.beantragtVon || '').toLowerCase() === email
    );
}

export function canDecide(state) {
    return state.role === 'admin';
}

/** KV darf Stellungnahme zu Anträgen der eigenen Klasse abgeben. */
export function canKvStellungnahme(item, state) {
    if (!item || !hasAktKvCapability(state)) return false;
    if (String(item.status || '').toLowerCase() !== 'kv_stellungnahme') return false;
    const kvKlassen = kvKlasseCodesForState(state);
    // Produktiv: ohne KV-Klassen keine Stellungnahme; Demo (?demoRole=1) bleibt offen
    if (!kvKlassen.length) return !!state.demoRoleOverride;
    return klasseCodesOf(item).some((c) => kvKlassen.includes(c));
}

/** Admin: genehmigte Einträge als in Untis verarbeitet markieren. */
export function canMarkUntis(item, state) {
    if (!canDecide(state) || !item) return false;
    return String(item.status || '').toLowerCase() === 'genehmigt';
}

/**
 * Status nach Speichern durch Lehrkraft / Admin.
 * Nachforderung → erneut in den Genehmigungspfad; genehmigt bleibt bei Admin-Edit genehmigt.
 * @param {object} prev
 * @param {object} state
 * @param {object} [rules]
 */
export function nextStatusAfterSave(prev, state, rules) {
    const prevStatus = String((prev && prev.status) || 'beantragt').toLowerCase();
    const resubmit = initialStatusForSubmit(rules || (state && state.rules) || {});
    if (prevStatus === 'entwurf') return resubmit;
    if (prevStatus === 'nachforderung') return resubmit;
    if (prevStatus === 'genehmigt' && state && state.role === 'admin') return 'genehmigt';
    if (prevStatus === 'abgelehnt') return resubmit;
    if (prevStatus === 'abgesagt') return resubmit;
    if (prevStatus === 'genehmigt') return resubmit;
    if (prevStatus === 'kv_stellungnahme') return 'kv_stellungnahme';
    return prevStatus === 'beantragt' ? 'beantragt' : resubmit;
}

export function scopeFromState(state, opts) {
    const o = opts || {};
    const isStaff = state.role === 'admin' && !state.demoPersonaEmail;
    const email = effectiveAccountEmail(state);
    const teacher = o.teacherCode || effectiveTeacherCode(state) || '';
    return {
        role: state.role,
        accountEmail: email,
        teacherCode: String(teacher || '').trim(),
        onlyMine: !!o.onlyMine || (!isStaff && !o.scopeAll && !o.onlyApprovedPublic),
        onlyOpen: !!o.onlyOpen,
        onlyApprovedPublic: !!o.onlyApprovedPublic,
        kvKlassen: Array.isArray(o.kvKlassen)
            ? o.kvKlassen
            : state.role === 'kv'
              ? kvKlasseCodesForState(state)
              : []
    };
}
