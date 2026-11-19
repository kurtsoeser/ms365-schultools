/**
 * Entry / Wiring: Schulaktivitäten-Planer
 */
import {
    createInitialState,
    loadStammdaten,
    matchTeacherByEmail,
    resolveRole,
    persistRole,
    loadSavedSiteUrl,
    persistSiteUrl,
    emptyForm,
    filterItems,
    viewsForRole,
    scopeFromState,
    canEditItem,
    nextStatusAfterSave,
    kvKlasseCodesForState,
    effectiveAccountEmail,
    labelMaps
} from './schulaktivitaeten-planer-state.js';
import {
    resolveAktContext,
    loadAllAktData,
    createAktivitaetItem,
    updateAktivitaetItem,
    deleteAktivitaetItem,
    updateRegelwerkItem
} from './schulaktivitaeten-planer-graph.js';
import { newEntityId } from './schulaktivitaeten-planer-schema.js';
import {
    validateAktivitaet,
    DEFAULT_RULES,
    DEFAULT_KOSTEN_LIMITS,
    appendAenderungsLog,
    joinCodeList,
    parseCodeList,
    initialStatusForSubmit,
    suggestTeachersForKlassen,
    resolveAktOwnership
} from './schulaktivitaeten-planer-logic.js';
import { applyUeToDraft } from './schulaktivitaeten-planer-zeitraster.js';
import {
    applyAktRoleFromEntra,
    isAktDemoRoleUiEnabled,
    readDemoRoleOverrideFromUrl,
    shellPreferredAktRole
} from './schulaktivitaeten-planer-entra-role.js';
import { hydrateStammdatenFromLiveSite } from '../../shared/school-stammdaten-spo-lists.js';
import {
    toolSiteUrlFromLicenseResult,
    applyToolSiteUrlLocally,
    bootstrapSchoolSiteFromLicense
} from '../../shared/tenant-school-site-bootstrap.js';
import {
    renderApp,
    syncAktPageChrome,
    readFormFromDom,
    readFiltersFromDom,
    readRulesForm,
    readKostenLimitsForm
} from './schulaktivitaeten-planer-ui.js';
import { buildIcs, downloadIcs } from './schulaktivitaeten-planer-export.js';
import {
    downloadAktIcs,
    upsertAktToMyCalendar
} from './schulaktivitaeten-planer-calendar.js';
import {
    loadFeedMeta,
    publishAktivitaetenIcsFeed
} from './schulaktivitaeten-planer-feed.js';
import {
    getDemoSeedPackage,
    buildLocalDemoState,
    DEMO_SITE_DEFAULT,
    DEMO_TEACHER_PERSONA,
    DEMO_SEED_TAG
} from './schulaktivitaeten-planer-demo-data.js';
import {
    parseDemoImportJson,
    applyDemoStammdatenLocal,
    seedDemoSchulaktivitaeten,
    resetDemoSchulaktivitaeten
} from './schulaktivitaeten-planer-demo-seed.js';

const state = createInitialState();
state.feedMeta = loadFeedMeta();
let root = null;
/** @type {((ev: KeyboardEvent) => void) | null} */
let drawerEscHandler = null;

async function copyText(text) {
    const t = String(text || '');
    if (!t) throw new Error('Kein Link vorhanden.');
    if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(t);
        return;
    }
    const ta = document.createElement('textarea');
    ta.value = t;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    ta.remove();
}

async function publishFeed() {
    if (state.feedBusy) return;
    const siteUrl = String(state.siteUrl || '').trim();
    if (!siteUrl || state.localDemoOnly) {
        toast('Bitte zuerst eine SharePoint-Site laden (kein reiner Demo-Modus).');
        return;
    }
    state.feedBusy = true;
    paint();
    try {
        const meta = await publishAktivitaetenIcsFeed({
            siteUrl,
            items: state.items,
            stammdaten: state.stammdaten,
            calendarName: 'Schulaktivitäten'
        });
        state.feedMeta = meta;
        toast(
            'ICS-Feed veröffentlicht: ' +
                (meta.count || 0) +
                ' genehmigte Termine. Link kopieren und in Outlook/Apple abonnieren.'
        );
    } catch (e) {
        toast(String((e && e.message) || e));
    } finally {
        state.feedBusy = false;
        paint();
    }
}

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else window.alert(msg);
}

function paint() {
    if (!root) return;
    renderApp(state, root);
    syncAktPageChrome(state);
    bindStatic();
    if (state.focusUntisQueue) {
        state.focusUntisQueue = false;
        const el = document.getElementById('aktUntisQueue');
        if (el && typeof el.scrollIntoView === 'function') {
            try {
                el.scrollIntoView({ behavior: 'smooth', block: 'start' });
            } catch {
                el.scrollIntoView();
            }
        }
    }
    import('../../shared/frontend-planner-chrome-policy.js')
        .then((m) => {
            if (m && typeof m.refreshFrontendPlannerChromeAudience === 'function') {
                return m.refreshFrontendPlannerChromeAudience();
            }
        })
        .then(() => {
            // Lehrer-App-Rolle: Setup-Chrome auch bei IT-Shell ausblenden
            syncAktPageChrome(state);
        })
        .catch(() => {
            syncAktPageChrome(state);
        });
}

function currentAccount() {
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const a = window.ms365AuthGetAccountInfo();
            if (a) {
                return {
                    email: a.upn || a.username || a.email || a.mail || '',
                    username: a.username || a.upn || '',
                    name: a.name || ''
                };
            }
        }
    } catch {
        /* ignore */
    }
    try {
        if (typeof window.ms365GetAccountInfo === 'function') {
            const a = window.ms365GetAccountInfo();
            if (a) return a;
        }
    } catch {
        /* ignore */
    }
    try {
        const acc = window.msalAccount || null;
        if (acc) {
            const claims = acc.idTokenClaims || {};
            return {
                email: acc.username || claims.preferred_username || claims.upn || acc.mail || '',
                name: acc.name || ''
            };
        }
    } catch {
        /* ignore */
    }
    return { email: '', name: '' };
}

function syncAccount() {
    try {
        const info =
            typeof window.ms365AuthGetAccountInfo === 'function'
                ? window.ms365AuthGetAccountInfo()
                : null;
        const upn =
            typeof window.ms365AuthGetUserPrincipalName === 'function'
                ? window.ms365AuthGetUserPrincipalName()
                : '';
        const a = currentAccount();
        state.accountEmail = String(
            (info && (info.upn || info.username || info.email || info.mail)) ||
                a.email ||
                a.username ||
                upn ||
                ''
        )
            .trim()
            .toLowerCase();
        state.accountName = String((info && info.name) || a.name || '').trim();
    } catch {
        state.accountEmail = '';
        state.accountName = '';
    }
    const local = loadStammdaten();
    const cur = state.stammdaten || { teachers: [], classes: [], students: [] };
    // Live-Listen (nach hydrate) nicht mit localStorage überschreiben
    state.stammdaten = {
        teachers:
            Array.isArray(cur.teachers) && cur.teachers.length ? cur.teachers : local.teachers || [],
        classes:
            Array.isArray(cur.classes) && cur.classes.length ? cur.classes : local.classes || [],
        students:
            Array.isArray(cur.students) && cur.students.length
                ? cur.students
                : local.students || []
    };
    const match = matchTeacherByEmail(state.stammdaten.teachers, effectiveAccountEmail(state));
    if (match && state.view === 'antrag' && !state.form.lehrerCode) {
        state.form.lehrerCode = match.code || '';
        state.form.lehrerEmail = effectiveAccountEmail(state);
    }
}

/**
 * Angemeldete E-Mail zuverlässig ermitteln (MSAL + optional Graph /me).
 * @returns {Promise<string>}
 */
async function resolveSignedInEmail() {
    syncAccount();
    let email = String(effectiveAccountEmail(state) || state.accountEmail || '')
        .trim()
        .toLowerCase();
    if (email.indexOf('@') >= 0) return email;
    try {
        if (typeof window.ms365AuthGetUserPrincipalName === 'function') {
            email = String(window.ms365AuthGetUserPrincipalName() || '')
                .trim()
                .toLowerCase();
            if (email.indexOf('@') >= 0) {
                state.accountEmail = email;
                return email;
            }
        }
    } catch {
        /* ignore */
    }
    try {
        const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
        if (api && typeof api.getGraphToken === 'function' && typeof api.graphJson === 'function') {
            const tok = await api.getGraphToken(['https://graph.microsoft.com/User.Read']);
            const me = await api.graphJson(
                'GET',
                '/me?$select=mail,userPrincipalName',
                tok,
                undefined,
                'v1.0'
            );
            email = String((me && (me.mail || me.userPrincipalName)) || '')
                .trim()
                .toLowerCase();
            if (email.indexOf('@') >= 0) {
                state.accountEmail = email;
                return email;
            }
        }
    } catch {
        /* ignore */
    }
    return String(state.accountEmail || '')
        .trim()
        .toLowerCase();
}

/** Lokale Demo: IT bleibt angemeldet, Filter/KV laufen als Dave Grohl. */
function applyDemoTeacherPersona() {
    if (!state.demoPersonaEmail) return;
    if (readDemoRoleOverrideFromUrl(false)) return;
    const kv = kvKlasseCodesForState(state);
    state.role = 'lehrer';
    state.planerRoles = kv.length ? ['lehrer', 'kv'] : ['lehrer'];
    state.planerAccessDenied = false;
    persistRole('lehrer');
}

/** Stellt sicher, dass Dave + KV-Klassen (2AK/4AK) in den Stammdaten liegen. */
function ensureDemoPersonaInStammdaten() {
    const sd = state.stammdaten || { teachers: [], classes: [] };
    const teachers = Array.isArray(sd.teachers) ? sd.teachers.slice() : [];
    if (!matchTeacherByEmail(teachers, DEMO_TEACHER_PERSONA.email)) {
        teachers.push({
            code: DEMO_TEACHER_PERSONA.code,
            name: DEMO_TEACHER_PERSONA.name,
            email: DEMO_TEACHER_PERSONA.email
        });
    }
    const classes = (Array.isArray(sd.classes) ? sd.classes : []).map((c) => {
        const code = String(c.code || '').trim().toUpperCase();
        if (code !== '2AK' && code !== '4AK') return c;
        return {
            ...c,
            headEmail: c.headEmail || DEMO_TEACHER_PERSONA.email,
            kvEmail: c.kvEmail || DEMO_TEACHER_PERSONA.email
        };
    });
    state.stammdaten = { ...sd, teachers, classes };
}

/**
 * Wenn IT (Kurt) Demo-Seed auf SharePoint hat und selbst keine Lehrkraft ist:
 * Ansicht automatisch als Dave Grohl – ohne Admin/KV/Lehrer-Chips.
 */
function maybeAutoDemoPersonaFromSeed() {
    if (state.demoPersonaEmail) return;
    if (readDemoRoleOverrideFromUrl(false)) return;
    const email = String(state.accountEmail || '')
        .trim()
        .toLowerCase();
    if (!email) return;
    if (matchTeacherByEmail((state.stammdaten && state.stammdaten.teachers) || [], email)) return;
    if (!(state.planerRoles || []).includes('admin')) return;
    const hasSeed = (state.items || []).some((i) => {
        const note = String(i.notiz || '');
        const id = String(i.aktivitaetId || '');
        return note.includes(DEMO_SEED_TAG) || id.startsWith('akt-demo-gro');
    });
    if (!hasSeed) return;
    ensureDemoPersonaInStammdaten();
    state.demoPersonaEmail = DEMO_TEACHER_PERSONA.email;
    state.demoPersonaName = DEMO_TEACHER_PERSONA.name;
    applyDemoTeacherPersona();
}

/**
 * @param {{ preferShellDefault?: boolean, preferredActiveRole?: string|null }} [opts]
 */
async function syncAccountAndRole(opts) {
    syncAccount();
    const preferShellDefault = !!(opts && opts.preferShellDefault);
    const hasPreferred = opts && Object.prototype.hasOwnProperty.call(opts, 'preferredActiveRole');
    await applyAktRoleFromEntra(state, {
        demoRoleOverride: state.demoRoleOverride,
        preferShellDefault,
        ...(hasPreferred
            ? { preferredActiveRole: opts.preferredActiveRole }
            : preferShellDefault
              ? { preferredActiveRole: null }
              : { preferredActiveRole: state.role })
    });
    applyDemoTeacherPersona();
    const allowed = viewsForRole(state.role, state.planerRoles, {
        kvKlassen: kvKlasseCodesForState(state)
    }).map((v) => v.id);
    if (state.view === 'auslastung') state.view = 'kalender';
    if (!allowed.includes(state.view)) state.view = 'dashboard';
}

async function onItShellChanged() {
    const shellRole = shellPreferredAktRole(state.planerRoles || []);
    await syncAccountAndRole({
        preferShellDefault: true,
        preferredActiveRole: shellRole
    });
    paint();
}

const STAMM_READ_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
];
const STAMM_RW_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

/**
 * Token für Klassen/Lehrer lesen – still bevor Popup (sonst bleibt Dropdown leer).
 * @param {{ getGraphToken: Function }} api
 */
async function acquireStammdatenReadToken(api) {
    if (typeof window.ms365AuthAcquireTokenSilent === 'function') {
        try {
            return await window.ms365AuthAcquireTokenSilent(STAMM_READ_SCOPES);
        } catch {
            /* weiter */
        }
        try {
            return await window.ms365AuthAcquireTokenSilent(STAMM_RW_SCOPES);
        } catch {
            /* weiter */
        }
    }
    try {
        return await api.getGraphToken(STAMM_READ_SCOPES);
    } catch {
        return await api.getGraphToken(STAMM_RW_SCOPES);
    }
}

/**
 * Lehrer/Klassen aus SharePoint-Live-Listen (Fallback: lokaler Cache).
 * @param {string} siteId
 */
async function hydrateLiveStammdaten(siteId) {
    if (state.localDemoOnly || !siteId) {
        state.stammdatenLive = {
            ok: false,
            error: siteId ? '' : 'Keine Site-URL',
            classCount: (state.stammdaten.classes || []).length,
            teacherCount: (state.stammdaten.teachers || []).length,
            missing: [],
            source: 'local'
        };
        return;
    }
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api || typeof api.getGraphToken !== 'function') {
        state.stammdatenLive = {
            ok: false,
            error: 'SharePoint-Graph nicht geladen',
            classCount: (state.stammdaten.classes || []).length,
            teacherCount: (state.stammdaten.teachers || []).length,
            missing: [],
            source: 'local'
        };
        return;
    }
    try {
        const tok = await acquireStammdatenReadToken(api);
        const result = await hydrateStammdatenFromLiveSite({
            token: tok,
            siteId: String(siteId),
            siteWebUrl: String(state.siteUrl || '').trim(),
            graph: api,
            local: loadStammdaten()
        });
        state.stammdaten = result.stammdaten;
        const liveClasses = (result.live && result.live.classes) || [];
        const liveTeachers = (result.live && result.live.teachers) || [];
        const missing = (result.live && result.live.missing) || [];
        const usedLive = !!(liveClasses.length || liveTeachers.length);
        const liveSource = (result.live && result.live.source) || '';
        let err = '';
        if (!result.ok && !liveClasses.length) err = result.error || 'Live-Laden fehlgeschlagen';
        else if (missing.length && !liveClasses.length) err = 'Listen fehlen: ' + missing.join(', ');
        else if (!liveClasses.length) {
            err =
                'Klassenliste liefert 0 Zeilen – Site: ' +
                (state.siteUrl || '–') +
                '. Als Einreicher Leserecht prüfen (IT: „Nur Berechtigungen“ erneut) oder „Klassen neu laden“.';
        }
        state.stammdatenLive = {
            ok: liveClasses.length > 0,
            error: err,
            classCount: liveClasses.length,
            teacherCount: liveTeachers.length,
            missing,
            siteUrl: String(state.siteUrl || '').trim(),
            source: usedLive
                ? liveSource === 'spo-rest'
                    ? 'sharepoint-spo'
                    : 'sharepoint'
                : result.ok
                  ? 'local-fallback'
                  : 'error'
        };
    } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        state.stammdatenLive = {
            ok: false,
            error: /nicht angemeldet|anmelden/i.test(msg)
                ? 'Bitte anmelden und danach „Klassen neu laden“ – sonst bleiben die Listen leer.'
                : msg,
            classCount: 0,
            teacherCount: 0,
            missing: [],
            source: 'error'
        };
    }
}

async function refreshData() {
    state.loading = true;
    state.error = '';
    paint();
    try {
        const ctx = await resolveAktContext(state.siteUrl);
        state.ctx = ctx;
        state.localDemoOnly = false;
        await hydrateLiveStammdaten(ctx.siteId);
        await syncAccountAndRole();
        const data = await loadAllAktData(ctx);
        state.items = data.items || [];
        state.rules = data.rules || { ...DEFAULT_RULES };
        const pullSource = String((data && data.source) || 'graph');
        state.aktivitaetenLive = {
            ok: pullSource === 'graph' || pullSource === 'spo-rest',
            count: (state.items || []).length,
            source: pullSource,
            error: pullSource === 'graph-error' ? 'Graph-Lesen eingeschränkt – ggf. REST-Fallback' : '',
            pulledAt: new Date().toISOString()
        };
        maybeAutoDemoPersonaFromSeed();
        state.loading = false;
        paint();
    } catch (e) {
        state.loading = false;
        const msg = e && e.message ? e.message : String(e);
        // Lokale Demo nicht mit SharePoint-403 übermalen
        if (state.localDemoOnly && (state.items || []).length) {
            state.error = '';
            state.aktivitaetenLive = {
                ok: false,
                count: (state.items || []).length,
                source: 'demo',
                error: '',
                pulledAt: state.aktivitaetenLive && state.aktivitaetenLive.pulledAt
                    ? state.aktivitaetenLive.pulledAt
                    : ''
            };
            console.warn('[akt] refresh skipped – local demo kept', e);
        } else {
            state.error = msg;
            state.aktivitaetenLive = {
                ok: false,
                count: (state.items || []).length,
                source: state.aktivitaetenLive && state.aktivitaetenLive.source
                    ? state.aktivitaetenLive.source
                    : 'error',
                error: msg,
                pulledAt: state.aktivitaetenLive && state.aktivitaetenLive.pulledAt
                    ? state.aktivitaetenLive.pulledAt
                    : ''
            };
        }
        paint();
    }
}

function bindStatic() {
    root.querySelectorAll('[data-akt-view]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-akt-view') || 'dashboard';
            state.detailId = null;
            if (state.view === 'antrag' && !state.editingItemId) {
                state.form = emptyForm(prefillTeacher());
                state.formStep = 1;
                state.formAttempted = false;
                state.formTouched = {};
                state.begleitungPickerOpen = false;
                state.weitereAngabenOpen = false;
                state.multiDay = false;
                state.timeMode = 'ue';
                state.submitSuccess = null;
            }
            paint();
            // Antrag: Klassen erzwingen nachladen, wenn Dropdown noch leer
            if (
                state.view === 'antrag' &&
                !(state.stammdaten && state.stammdaten.classes && state.stammdaten.classes.length)
            ) {
                void (async () => {
                    const url = await ensureAktSiteUrl();
                    if (url) await refreshData().catch(() => {});
                })();
            }
        });
    });

    root.querySelectorAll('[data-akt-view-jump]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-akt-view-jump') || 'dashboard';
            state.focusUntisQueue = btn.hasAttribute('data-akt-untis-focus');
            const fs = btn.getAttribute('data-akt-filter-status');
            const fv = btn.getAttribute('data-akt-filter-verarbeitung');
            if (fs != null || fv != null) {
                // Explizite Jump-Filter: andere Filter leeren, sonst wirkt die Liste „leer“
                state.filters = {
                    klasse: '',
                    typ: '',
                    status: fs != null ? String(fs) : '',
                    lehrer: '',
                    beteiligung: '',
                    verarbeitung: fv != null ? String(fv) : ''
                };
                state.filterMultiOpen = { klasse: false, typ: false, status: false, lehrer: false };
            }
            paint();
        });
    });
    root.querySelectorAll('[data-akt-reload-stammdaten]').forEach((btn) => {
        btn.addEventListener('click', () => {
            void (async () => {
                // Ausdrücklicher Klick: hier darf bei Bedarf die Anmeldung starten.
                const url = await ensureAktSiteUrl({ interactive: true });
                if (!url) {
                    toast(
                        state.roleHint ||
                            'Schul-Site fehlt in diesem Browser. IT: unter Admin → Schule „Tool-Site-URL“ setzen (oder Listen-Assistent erneut prüfen) und IT-Sicherung ablegen – dann neu laden.'
                    );
                    paint();
                    return;
                }
                try {
                    await refreshData();
                    const n = (state.stammdaten.classes || []).length;
                    if (n) toast(n + ' Klassen geladen.');
                    else
                        toast(
                            state.stammdatenLive && state.stammdatenLive.error
                                ? state.stammdatenLive.error
                                : 'Immer noch keine Klassen – Leserecht als Einreicher prüfen.'
                        );
                } catch (e) {
                    toast(e && e.message ? e.message : String(e));
                }
            })();
        });
    });
    root.querySelectorAll('[data-akt-status-preset]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.filters = {
                ...state.filters,
                status: String(btn.getAttribute('data-akt-status-preset') || ''),
                verarbeitung: ''
            };
            paint();
        });
    });
    const calToday = document.getElementById('aktCalToday');
    if (calToday) {
        calToday.addEventListener('click', () => {
            const n = new Date();
            n.setHours(0, 0, 0, 0);
            const day = (n.getDay() + 6) % 7;
            const monday = new Date(n);
            monday.setDate(n.getDate() - day);
            state.calYear = n.getFullYear();
            state.calMonth = n.getMonth() + 1;
            state.calWeekAnchor = monday.toISOString();
            state.view = 'kalender';
            paint();
        });
    }
    root.querySelectorAll('[data-akt-dirty-form]').forEach((form) => {
        const kind = form.getAttribute('data-akt-dirty-form');
        const mark = () => {
            if (kind === 'limits') state.limitsDirty = true;
            if (kind === 'regeln') state.rulesDirty = true;
            state.settingsSavedFlash = '';
            const btn = form.querySelector('button[type="submit"]');
            if (btn) btn.disabled = false;
            const flash = form.querySelector('.akt-save-flash');
            if (flash) flash.remove();
        };
        form.addEventListener('input', mark);
        form.addEventListener('change', mark);
    });

    root.querySelectorAll('[data-akt-role]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const demoUi = isAktDemoRoleUiEnabled(
                state.entraGroupsConfigured,
                state.demoRoleOverride
            );
            const raw = String(btn.getAttribute('data-akt-role') || '').toLowerCase();
            if (!demoUi && !(state.planerRoles || []).includes(raw)) return;
            const role = raw === 'admin' ? 'admin' : raw === 'kv' ? 'kv' : 'lehrer';
            state.role = role;
            if (demoUi) state.demoRoleOverride = true;
            persistRole(role);
            const allowed = viewsForRole(role, state.planerRoles, {
                kvKlassen: kvKlasseCodesForState(state)
            }).map((v) => v.id);
            if (state.view === 'auslastung') state.view = 'kalender';
            if (!allowed.includes(state.view)) state.view = 'dashboard';
            paint();
        });
    });

    const loadBtn = document.getElementById('aktBtnLoad');
    if (loadBtn) {
        loadBtn.addEventListener('click', () => {
            const input = document.getElementById('aktSiteUrl');
            state.siteUrl = input ? String(input.value || '').trim() : state.siteUrl;
            persistSiteUrl(state.siteUrl);
            refreshData();
        });
    }

    ['aktFilterBeteiligung', 'aktFilterVerarbeitung'].forEach((id) => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('change', () => {
                state.filters = readFiltersFromDom();
                paint();
            });
        }
    });
    if (!state.filterMultiOpen || typeof state.filterMultiOpen !== 'object') {
        state.filterMultiOpen = { klasse: false, typ: false, status: false, lehrer: false };
    }
    if (!state.filterMultiQuery || typeof state.filterMultiQuery !== 'object') {
        state.filterMultiQuery = { klasse: '', typ: '', status: '', lehrer: '' };
    }
    root.querySelectorAll('[data-akt-filter-multi]').forEach((box) => {
        const key = String(box.getAttribute('data-akt-filter-multi') || '').trim();
        if (!key) return;
        box.addEventListener('toggle', () => {
            state.filterMultiOpen = { ...state.filterMultiOpen, [key]: !!box.open };
        });
        box.addEventListener('change', (ev) => {
            const t = ev.target;
            if (!t || t.type !== 'checkbox') return;
            state.filterMultiOpen = { ...state.filterMultiOpen, [key]: true };
            state.filters = readFiltersFromDom();
            paint();
        });
        const q = box.querySelector('[data-akt-filter-multi-q="' + key + '"]');
        if (q) {
            q.addEventListener('input', () => {
                const needle = String(q.value || '');
                state.filterMultiQuery = { ...state.filterMultiQuery, [key]: needle };
                const low = needle.trim().toLowerCase();
                box.querySelectorAll('.akt-filter-multi__item').forEach((row) => {
                    const hay = String(row.textContent || '')
                        .trim()
                        .toLowerCase();
                    row.classList.toggle('is-filtered-out', !!(low && hay.indexOf(low) === -1));
                });
            });
        }
        const clearBtn = box.querySelector('[data-akt-filter-multi-clear="' + key + '"]');
        if (clearBtn) {
            clearBtn.addEventListener('click', (ev) => {
                ev.preventDefault();
                state.filterMultiOpen = { ...state.filterMultiOpen, [key]: true };
                state.filters = { ...state.filters, [key]: '' };
                paint();
            });
        }
    });
    const meineKlassePick = document.getElementById('aktMeineKlassePick');
    if (meineKlassePick) {
        meineKlassePick.addEventListener('change', () => {
            state.meineKlasseCode = String(meineKlassePick.value || '').trim();
            paint();
        });
    }
    const reset = document.getElementById('aktFilterReset');
    if (reset) {
        reset.addEventListener('click', () => {
            state.filterMultiOpen = { klasse: false, typ: false, status: false, lehrer: false };
            state.filterMultiQuery = { klasse: '', typ: '', status: '', lehrer: '' };
            state.filters = {
                klasse: '',
                typ: '',
                status: '',
                lehrer: '',
                beteiligung: '',
                verarbeitung: ''
            };
            paint();
        });
    }

    const form = document.getElementById('aktForm');
    if (form) {
        form.addEventListener('submit', (ev) => {
            ev.preventDefault();
            state.formAttempted = true;
            saveForm({ submit: true }).catch((e) => toast(e && e.message ? e.message : String(e)));
        });
        form.addEventListener('input', (ev) => {
            const t = ev.target;
            if (t && t.getAttribute && t.getAttribute('data-akt-mcheck-filter')) {
                filterCheckListOptions(t.getAttribute('data-akt-mcheck-filter'), t.value);
                return;
            }
            state.form = { ...state.form, ...readFormFromDom() };
            if (!state.multiDay) state.form.enddatum = state.form.startdatum;
        });
        form.addEventListener('change', (ev) => {
            const t = ev.target;
            if (t && t.getAttribute && t.getAttribute('data-akt-mcheck-filter')) {
                return;
            }
            // Checkbox-Liste: State aktualisieren, aber kein paint() (sonst Fokus/Tab weg)
            if (t && t.closest && t.closest('[data-akt-mcheck]')) {
                const prevCodes = String(state.form.klasseCodes || '');
                state.form = { ...state.form, ...readFormFromDom() };
                if (!state.multiDay) state.form.enddatum = state.form.startdatum;
                const box = t.closest('[data-akt-mcheck]');
                const name = box && box.getAttribute('data-akt-mcheck');
                updateCheckListCount(name);
                if (name === 'klasseCodes' && String(state.form.klasseCodes || '') !== prevCodes) {
                    mergeSuggestedTeachers();
                }
                return;
            }
            state.form = { ...state.form, ...readFormFromDom() };
            if (!state.multiDay) state.form.enddatum = state.form.startdatum;
            // Nur bei Layout-Wechsel neu rendern – sonst bricht Tab-Navigation
            const fieldName = t && t.name ? String(t.name) : '';
            if (fieldName === 'teilnahmeModus') {
                paint();
            }
        });
    }
    root.querySelectorAll('[data-akt-step-next]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.form = { ...state.form, ...readFormFromDom() };
            state.formAttempted = true;
            state.formStep = Math.min(4, (Number(state.formStep) || 1) + 1);
            paint();
        });
    });
    root.querySelectorAll('[data-akt-step-prev]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.form = { ...state.form, ...readFormFromDom() };
            state.formStep = Math.max(1, (Number(state.formStep) || 1) - 1);
            paint();
        });
    });
    root.querySelectorAll('[data-akt-goto-step]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.form = { ...state.form, ...readFormFromDom() };
            const n = Number(btn.getAttribute('data-akt-goto-step') || '1');
            state.formStep = Math.min(4, Math.max(1, n || 1));
            paint();
        });
    });
    root.querySelectorAll('[data-akt-touch]').forEach((el) => {
        el.addEventListener('blur', () => {
            const key = el.getAttribute('data-akt-touch');
            if (!key) return;
            state.formTouched = { ...(state.formTouched || {}), [key]: true };
            state.form = { ...state.form, ...readFormFromDom() };
            // Kein paint() auf Blur – sonst springt der Tab-Fokus weg
        });
    });
    root.querySelectorAll('[data-akt-time-mode]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.form = { ...state.form, ...readFormFromDom() };
            state.timeMode = btn.getAttribute('data-akt-time-mode') === 'clock' ? 'clock' : 'ue';
            paint();
        });
    });
    root.querySelectorAll('[data-akt-toggle-begleitung]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.form = { ...state.form, ...readFormFromDom() };
            state.begleitungPickerOpen = !state.begleitungPickerOpen;
            paint();
        });
    });
    const begPick = document.getElementById('aktBegleitungPick');
    if (begPick) {
        begPick.addEventListener('change', () => {
            const code = String(begPick.value || '').trim();
            if (!code) return;
            state.form = { ...state.form, ...readFormFromDom() };
            const cur = parseCodeList(state.form.begleitungCodes || state.form.begleitung);
            if (!cur.includes(code)) cur.push(code);
            state.form.begleitungCodes = joinCodeList(cur);
            state.form.begleitung = state.form.begleitungCodes;
            state.begleitungPickerOpen = false;
            paint();
        });
    }
    root.querySelectorAll('[data-akt-remove-begleitung]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const code = String(btn.getAttribute('data-akt-remove-begleitung') || '').trim();
            state.form = { ...state.form, ...readFormFromDom() };
            const cur = parseCodeList(state.form.begleitungCodes || state.form.begleitung).filter(
                (c) => c !== code
            );
            state.form.begleitungCodes = joinCodeList(cur);
            state.form.begleitung = state.form.begleitungCodes;
            paint();
        });
    });
    const weitere = root.querySelector('[data-akt-weitere]');
    if (weitere) {
        weitere.addEventListener('toggle', () => {
            state.weitereAngabenOpen = !!weitere.open;
        });
    }
    root.querySelectorAll('[data-akt-save-draft]').forEach((btn) => {
        btn.addEventListener('click', () => {
            saveForm({ draft: true }).catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    });

    const rulesForm = document.getElementById('aktRulesForm');
    if (rulesForm) {
        rulesForm.addEventListener('submit', (ev) => {
            ev.preventDefault();
            saveRules().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }

    const limitsForm = document.getElementById('aktKostenLimitsForm');
    if (limitsForm) {
        limitsForm.addEventListener('submit', (ev) => {
            ev.preventDefault();
            saveKostenLimits().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }
    const limitsReset = document.getElementById('aktKostenLimitsReset');
    if (limitsReset) {
        limitsReset.addEventListener('click', () => {
            if (
                !window.confirm(
                    'Kostenschwellen auf den Standard zurücksetzen? Nicht gespeicherte Änderungen gehen verloren.'
                )
            ) {
                return;
            }
            state.rules = {
                ...(state.rules || { ...DEFAULT_RULES }),
                kostenLimits: DEFAULT_KOSTEN_LIMITS.map((r) => ({ ...r })),
                itemId: (state.rules && state.rules.itemId) || 'local-rw'
            };
            state.limitsDirty = true;
            state.settingsSavedFlash = '';
            paint();
            toast('Standard wiederhergestellt – bitte speichern.');
        });
    }

    const openDetail = (id) => {
        if (!id) return;
        state.detailId = id;
        state.submitSuccess = null;
        paint();
    };

    root.querySelectorAll('[data-akt-detail]').forEach((el) => {
        const handler = (ev) => {
            ev.preventDefault();
            ev.stopPropagation();
            openDetail(el.getAttribute('data-akt-detail'));
        };
        el.addEventListener('click', handler);
        el.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter' || ev.key === ' ') {
                ev.preventDefault();
                handler(ev);
            }
        });
    });
    const closeDetail = () => {
        state.detailId = null;
        paint();
    };
    const close = document.getElementById('aktDetailClose');
    if (close) close.addEventListener('click', closeDetail);
    const closeBtn = document.getElementById('aktDetailCloseBtn');
    if (closeBtn) closeBtn.addEventListener('click', closeDetail);
    if (drawerEscHandler) {
        document.removeEventListener('keydown', drawerEscHandler);
        drawerEscHandler = null;
    }
    if (state.detailId) {
        drawerEscHandler = (ev) => {
            if (ev.key !== 'Escape') return;
            document.removeEventListener('keydown', drawerEscHandler);
            drawerEscHandler = null;
            closeDetail();
        };
        document.addEventListener('keydown', drawerEscHandler);
    }

    root.querySelectorAll('[data-akt-success-open]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-akt-success-open');
            state.submitSuccess = null;
            state.view = 'liste';
            openDetail(id);
        });
    });
    root.querySelectorAll('[data-akt-success-home]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.submitSuccess = null;
            state.detailId = null;
            state.view = 'dashboard';
            paint();
        });
    });
    root.querySelectorAll('[data-akt-multiday]').forEach((el) => {
        el.addEventListener('change', () => {
            state.form = { ...state.form, ...readFormFromDom() };
            state.multiDay = !!el.checked;
            if (!state.multiDay) {
                state.form.enddatum = state.form.startdatum;
            } else if (!state.form.enddatum) {
                state.form.enddatum = state.form.startdatum;
            }
            paint();
        });
    });

    root.querySelectorAll('[data-akt-edit]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-akt-edit');
            const it = state.items.find((x) => x.itemId === id);
            if (!it || !canEditItem(it, state)) return;
            state.editingItemId = it.itemId;
            state.form = emptyForm(it);
            state.formStep = 1;
            state.formAttempted = false;
            state.formTouched = {};
            state.begleitungPickerOpen = false;
            state.weitereAngabenOpen = false;
            state.multiDay = !!(it.enddatum && it.startdatum && it.enddatum !== it.startdatum);
            state.timeMode = it.startUe || it.endUe ? 'ue' : 'clock';
            state.detailId = null;
            state.submitSuccess = null;
            state.view = 'antrag';
            paint();
        });
    });

    root.querySelectorAll('[data-akt-delete]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-akt-delete');
            deleteItem(id).catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    });
    root.querySelectorAll('[data-akt-cancel]').forEach((btn) => {
        btn.addEventListener('click', () => {
            cancelItem(btn.getAttribute('data-akt-cancel')).catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });

    root.querySelectorAll('[data-akt-cal-ics]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-akt-cal-ics');
            const it = (state.items || []).find((x) => x.aktivitaetId === id || x.itemId === id);
            if (!it) {
                toast('Aktivität nicht gefunden.');
                return;
            }
            try {
                downloadAktIcs(it, state.stammdaten);
                toast('ICS heruntergeladen – in Outlook/Apple/Google öffnen.');
            } catch (e) {
                toast(e && e.message ? e.message : String(e));
            }
        });
    });
    root.querySelectorAll('[data-akt-cal-mine]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-akt-cal-mine');
            putAktInMyCalendar(id, { invite: false }).catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });
    root.querySelectorAll('[data-akt-cal-invite]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-akt-cal-invite');
            const ok = window.confirm(
                'Termin in deinem Outlook-Kalender anlegen und Einladungen an Klassen-Teams / Begleitung senden?'
            );
            if (!ok) return;
            putAktInMyCalendar(id, { invite: true }).catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });

    root.querySelectorAll('[data-akt-approve]').forEach((btn) => {
        btn.addEventListener('click', () => {
            decide(btn.getAttribute('data-akt-approve'), 'genehmigt').catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });
    root.querySelectorAll('[data-akt-reject]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const grund =
                window.prompt('Warum wird der Antrag abgelehnt? (Pflichtfeld)', '') || '';
            if (!String(grund).trim()) {
                toast('Bitte eine Begründung für die Ablehnung eingeben.');
                return;
            }
            decide(btn.getAttribute('data-akt-reject'), 'abgelehnt', grund).catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });
    root.querySelectorAll('[data-akt-nachforderung]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const text =
                window.prompt('Welche Angaben werden noch benötigt? (Pflichtfeld)', '') || '';
            if (!String(text).trim()) {
                toast('Bitte angeben, welche Angaben noch benötigt werden.');
                return;
            }
            decide(btn.getAttribute('data-akt-nachforderung'), 'nachforderung', '', text).catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });
    root.querySelectorAll('[data-akt-untis]').forEach((btn) => {
        btn.addEventListener('click', () => {
            markUntisDone(btn.getAttribute('data-akt-untis')).catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });
    root.querySelectorAll('[data-akt-kv]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const text =
                window.prompt(
                    'Stellungnahme des Klassenvorstands (z. B. bisherige Abwesenheiten der Klasse):',
                    ''
                ) || '';
            if (!String(text).trim()) {
                toast('Bitte eine Stellungnahme eingeben.');
                return;
            }
            submitKvStellungnahme(btn.getAttribute('data-akt-kv'), text).catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });

    root.querySelectorAll('[data-akt-cal-mode]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const mode = String(btn.getAttribute('data-akt-cal-mode') || 'month');
            state.calMode = mode === 'agenda' || mode === 'week' || mode === 'month' ? mode : 'month';
            paint();
        });
    });
    root.querySelectorAll('[data-akt-cal-nav]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const delta = Number(btn.getAttribute('data-akt-cal-nav') || '0');
            const mode = String(state.calMode || 'month');
            if (mode === 'week') {
                let anchor = new Date(state.calWeekAnchor || Date.now());
                if (Number.isNaN(anchor.getTime())) anchor = new Date();
                if (delta === 0) {
                    const d = new Date();
                    d.setHours(0, 0, 0, 0);
                    const day = (d.getDay() + 6) % 7;
                    d.setDate(d.getDate() - day);
                    state.calWeekAnchor = d.toISOString();
                    state.calYear = d.getFullYear();
                    state.calMonth = d.getMonth() + 1;
                } else {
                    anchor.setDate(anchor.getDate() + delta * 7);
                    const day = (anchor.getDay() + 6) % 7;
                    anchor.setDate(anchor.getDate() - day);
                    state.calWeekAnchor = anchor.toISOString();
                    state.calYear = anchor.getFullYear();
                    state.calMonth = anchor.getMonth() + 1;
                }
            } else {
                if (delta === 0) {
                    const n = new Date();
                    state.calYear = n.getFullYear();
                    state.calMonth = n.getMonth() + 1;
                } else {
                    state.calMonth += delta;
                    if (state.calMonth < 1) {
                        state.calMonth = 12;
                        state.calYear -= 1;
                    } else if (state.calMonth > 12) {
                        state.calMonth = 1;
                        state.calYear += 1;
                    }
                }
            }
            paint();
        });
    });

    const reportPrev = document.getElementById('aktReportPrev');
    const reportNext = document.getElementById('aktReportNext');
    if (reportPrev) {
        reportPrev.addEventListener('click', () => {
            state.reportYear = Number(state.reportYear || new Date().getFullYear()) - 1;
            paint();
        });
    }
    if (reportNext) {
        reportNext.addEventListener('click', () => {
            state.reportYear = Number(state.reportYear || new Date().getFullYear()) + 1;
            paint();
        });
    }

    const icsBtn = document.getElementById('aktBtnIcs');
    if (icsBtn) {
        icsBtn.addEventListener('click', () => {
            const scope = scopeFromState(state, { scopeAll: state.role === 'admin' });
            const items = filterItems(state.items, state.filters, scope).filter(
                (i) => String(i.status).toLowerCase() !== 'abgelehnt'
            );
            const ics = buildIcs(items, state.stammdaten, 'Schulaktivitäten');
            downloadIcs(ics);
            toast('ICS exportiert (' + items.length + ' Termine).');
        });
    }

    const feedBtn = document.getElementById('aktBtnFeed');
    if (feedBtn) {
        feedBtn.disabled = !!state.feedBusy;
        feedBtn.addEventListener('click', () => publishFeed().catch((e) => toast(String((e && e.message) || e))));
    }
    const feedRefresh = document.getElementById('aktBtnFeedRefresh');
    if (feedRefresh) {
        feedRefresh.disabled = !!state.feedBusy;
        feedRefresh.addEventListener('click', () => publishFeed().catch((e) => toast(String((e && e.message) || e))));
    }
    const copyWebcal = document.getElementById('aktBtnFeedCopyWebcal');
    if (copyWebcal) {
        copyWebcal.addEventListener('click', () => {
            const url = (state.feedMeta && (state.feedMeta.webcalUrl || state.feedMeta.httpsUrl)) || '';
            copyText(url)
                .then(() => toast('webcal-Link kopiert – in Outlook unter „Kalender hinzufügen → Aus dem Internet“ einfügen.'))
                .catch((e) => toast(String((e && e.message) || e)));
        });
    }
    const copyHttps = document.getElementById('aktBtnFeedCopyHttps');
    if (copyHttps) {
        copyHttps.addEventListener('click', () => {
            const url = (state.feedMeta && state.feedMeta.httpsUrl) || '';
            copyText(url)
                .then(() => toast('HTTPS-Link kopiert.'))
                .catch((e) => toast(String((e && e.message) || e)));
        });
    }

    const wireDemo = (id) => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', () => runDemoImport(getDemoSeedPackage()).catch((e) => toast(String((e && e.message) || e))));
    };
    wireDemo('aktBtnDemo');

    const wireReset = (id) => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', () => runDemoReset().catch((e) => toast(String((e && e.message) || e))));
    };
    wireReset('aktBtnDemoReset');

    const importInput = document.getElementById('aktImportJson');
    if (importInput) {
        importInput.addEventListener('change', () => {
            const file = importInput.files && importInput.files[0];
            importInput.value = '';
            if (!file) return;
            const reader = new FileReader();
            reader.onload = () => {
                try {
                    const pack = parseDemoImportJson(String(reader.result || ''));
                    runDemoImport(pack).catch((e) => toast(String((e && e.message) || e)));
                } catch (e) {
                    toast(e && e.message ? e.message : String(e));
                }
            };
            reader.onerror = () => toast('Datei konnte nicht gelesen werden.');
            reader.readAsText(file, 'UTF-8');
        });
    }
}

function prefillTeacher() {
    syncAccount();
    const email = effectiveAccountEmail(state) || state.accountEmail || '';
    const match = matchTeacherByEmail(state.stammdaten.teachers, email);
    return {
        lehrerCode: match ? match.code || '' : '',
        lehrerEmail: email
    };
}

function filterCheckListOptions(name, query) {
    const q = String(query || '')
        .trim()
        .toLowerCase();
    const box = root && root.querySelector('[data-akt-mcheck="' + name + '"]');
    if (!box) return;
    box.querySelectorAll('.akt-mcheck__item').forEach((row) => {
        const hay = String(row.textContent || '')
            .trim()
            .toLowerCase();
        row.classList.toggle('is-filtered-out', !!(q && hay.indexOf(q) === -1));
    });
}

function updateCheckListCount(name) {
    if (!name || !root) return;
    const box = root.querySelector('[data-akt-mcheck="' + name + '"]');
    if (!box) return;
    const n = box.querySelectorAll('input[type="checkbox"]:checked').length;
    const el = box.querySelector('[data-akt-mcheck-count="' + name + '"]');
    if (el) el.textContent = n ? n + ' ausgewählt' : 'Keine Auswahl';
}

function mergeSuggestedTeachers() {
    const suggested = suggestTeachersForKlassen(
        state.form.klasseCodes || state.form.klasseCode,
        (state.stammdaten && state.stammdaten.teachers) || []
    );
    if (!suggested.length) return;
    const current = parseCodeList(state.form.begleitungCodes || state.form.begleitung);
    let added = 0;
    suggested.forEach((code) => {
        if (!current.includes(code)) {
            current.push(code);
            added += 1;
        }
    });
    if (!added) return;
    state.form.begleitungCodes = joinCodeList(current);
    state.form.begleitung = state.form.begleitungCodes;
    toast(added + ' betroffene Lehrkraft/Lehrkräfte vorgeschlagen.');
}

function enrichDraftForSave(draft, prev) {
    const next = applyUeToDraft({ ...draft }).draft;
    next.begleitungCodes = joinCodeList(next.begleitungCodes || next.begleitung);
    const teachers = (state.stammdaten && state.stammdaten.teachers) || [];
    const teacherByCode = Object.fromEntries(
        teachers.filter((t) => t && t.code).map((t) => [String(t.code), t.name || t.code])
    );
    // SharePoint: BegleitungCodes = Kürzel (maschinenlesbar), Begleitung = Anzeigenamen
    next.begleitung = parseCodeList(next.begleitungCodes)
        .map((c) => teacherByCode[c] || c)
        .join('; ');
    next.klasseCodes = joinCodeList(next.klasseCodes || next.klasseCode);
    const own = resolveAktOwnership(
        {
            ...next,
            beantragtVon: next.beantragtVon || (prev && prev.beantragtVon) || '',
            lehrerEmail: next.lehrerEmail || '',
            lehrerCode: next.lehrerCode || ''
        },
        { accountEmail: state.accountEmail, teachers }
    );
    next.beantragtVon = own.beantragtVon;
    next.lehrerEmail = own.lehrerEmail;
    next.lehrerCode = own.lehrerCode;

    const prevStatus = String((prev && prev.status) || next.status || 'beantragt').toLowerCase();
    next.status = nextStatusAfterSave({ status: prevStatus }, state, state.rules);
    if (
        prevStatus === 'nachforderung' &&
        (next.status === 'beantragt' || next.status === 'kv_stellungnahme')
    ) {
        next.aenderungsLog = appendAenderungsLog(
            (prev && prev.aenderungsLog) || next.aenderungsLog,
            'Erneut eingereicht nach Nachforderung'
        );
        next.rueckfrageText = (prev && prev.rueckfrageText) || next.rueckfrageText || '';
    } else if (prevStatus === 'genehmigt' && state.role === 'admin') {
        next.aenderungsLog = appendAenderungsLog(
            (prev && prev.aenderungsLog) || next.aenderungsLog,
            'Admin-Änderung an genehmigtem Antrag'
        );
        next.untisErledigt = !!(prev && prev.untisErledigt);
        next.untisErledigtVon = (prev && prev.untisErledigtVon) || '';
        next.untisErledigtAm = (prev && prev.untisErledigtAm) || '';
        next.genehmigtVon = (prev && prev.genehmigtVon) || next.genehmigtVon;
        next.genehmigtAm = (prev && prev.genehmigtAm) || next.genehmigtAm;
    } else if (prev) {
        next.aenderungsLog = (prev && prev.aenderungsLog) || next.aenderungsLog || '';
        next.rueckfrageText = (prev && prev.rueckfrageText) || next.rueckfrageText || '';
        next.untisErledigt = !!(prev && prev.untisErledigt);
        next.untisErledigtVon = (prev && prev.untisErledigtVon) || '';
        next.untisErledigtAm = (prev && prev.untisErledigtAm) || '';
    }
    return next;
}

/**
 * @param {{ draft?: boolean, submit?: boolean }} [opts]
 */
async function saveForm(opts) {
    const asDraft = !!(opts && opts.draft);
    const prev = state.editingItemId
        ? state.items.find((x) => x.itemId === state.editingItemId)
        : null;
    const signedInEmail = await resolveSignedInEmail();
    if (!signedInEmail || signedInEmail.indexOf('@') < 0) {
        throw new Error(
            'Angemeldete E-Mail konnte nicht gelesen werden. Bitte oben rechts ab- und wieder anmelden, dann erneut einreichen.'
        );
    }
    state.accountEmail = signedInEmail;
    if (state.form) {
        if (!state.form.lehrerEmail) state.form.lehrerEmail = signedInEmail;
        if (!state.form.beantragtVon) state.form.beantragtVon = signedInEmail;
    }
    let draft = enrichDraftForSave(
        {
            ...state.form,
            ...readFormFromDom(),
            lehrerEmail: signedInEmail,
            beantragtVon: signedInEmail
        },
        prev || null
    );
    draft.lehrerEmail = draft.lehrerEmail || signedInEmail;
    draft.beantragtVon = draft.beantragtVon || signedInEmail;

    const rules = state.rules || DEFAULT_RULES;
    if (asDraft) {
        if (!String(draft.titel || '').trim()) {
            state.form = { ...draft };
            state.formStep = 1;
            paint();
            throw new Error('Bitte gib der Aktivität einen Titel, damit der Entwurf gespeichert werden kann.');
        }
        draft.status = 'entwurf';
    } else {
        const check = validateAktivitaet({
            draft: {
                ...draft,
                aktivitaetId: (prev && prev.aktivitaetId) || state.form.aktivitaetId || ''
            },
            existing: state.items,
            rules: {
                minVorlaufTage: rules.minVorlaufTage,
                maxGleichzeitigProKlasse: rules.maxGleichzeitigProKlasse,
                vorlaufModus: rules.vorlaufModus,
                genehmigungPfad: rules.genehmigungPfad
            }
        });
        if (!check.ok) {
            state.form = { ...draft, status: (prev && prev.status) || draft.status };
            state.formStep = 4;
            paint();
            throw new Error(check.errors[0] || 'Validierung fehlgeschlagen.');
        }
        if (!prev || String(prev.status).toLowerCase() === 'entwurf') {
            draft.status = initialStatusForSubmit(rules);
        }
    }

    const submitStatus = draft.status;

    if (state.localDemoOnly || !state.ctx) {
        if (state.editingItemId) {
            state.items = state.items.map((it) =>
                it.itemId === state.editingItemId
                    ? {
                          ...it,
                          ...draft,
                          aktivitaetId: it.aktivitaetId || draft.aktivitaetId || newEntityId('akt')
                      }
                    : it
            );
            toast(
                asDraft
                    ? 'Entwurf gespeichert (Demo).'
                    : (draft.status === 'beantragt' || draft.status === 'kv_stellungnahme') &&
                        String((prev && prev.status) || '') === 'nachforderung'
                      ? 'Erneut eingereicht (Demo).'
                      : 'Aktivität aktualisiert (Demo).'
            );
        } else {
            const id = newEntityId('akt');
            state.items = [
                {
                    itemId: 'local-' + id,
                    aktivitaetId: id,
                    ...draft,
                    status: submitStatus,
                    beantragtVon: state.accountEmail,
                    _localOnly: true
                },
                ...state.items
            ];
            toast(
                asDraft
                    ? 'Entwurf gespeichert (Demo).'
                    : submitStatus === 'kv_stellungnahme'
                      ? 'Eingereicht – wartet auf Prüfung (Demo).'
                      : 'Zur Genehmigung eingereicht (Demo).'
            );
        }
        const savedRef = state.editingItemId
            ? state.items.find((x) => x.itemId === state.editingItemId)
            : state.items[0];
        const savedId =
            (savedRef && (savedRef.aktivitaetId || savedRef.itemId)) || state.editingItemId || '';
        const savedTitle = draft.titel || 'Aktivität';
        state.editingItemId = null;
        state.form = emptyForm(prefillTeacher());
        state.formStep = 1;
        state.formAttempted = false;
        state.formTouched = {};
        state.multiDay = false;
        state.weitereAngabenOpen = false;
        state.localDemoOnly = true;
        if (asDraft) {
            state.submitSuccess = null;
            state.view = 'liste';
        } else {
            state.view = 'dashboard';
            state.submitSuccess = { id: savedId, titel: savedTitle };
            state.detailId = null;
        }
        paint();
        return;
    }

    state.loading = true;
    paint();
    try {
        /** @type {object|null} */
        let createdLocal = null;
        const labels = labelMaps(state.stammdaten);
        if (state.editingItemId) {
            await updateAktivitaetItem(
                state.ctx,
                state.editingItemId,
                {
                    ...prev,
                    ...draft,
                    aktivitaetId: (prev && prev.aktivitaetId) || draft.aktivitaetId || newEntityId('akt')
                },
                {
                    teacherLabels: labels.lehrer,
                    accountEmail: state.accountEmail,
                    teachers: (state.stammdaten && state.stammdaten.teachers) || []
                }
            );
            toast(asDraft ? 'Entwurf gespeichert.' : 'Aktivität aktualisiert.');
        } else {
            const newId = newEntityId('akt');
            draft.aktivitaetId = newId;
            const account = String(state.accountEmail || '')
                .trim()
                .toLowerCase();
            createdLocal = await createAktivitaetItem(
                state.ctx,
                {
                    ...draft,
                    aktivitaetId: newId,
                    status: submitStatus,
                    beantragtVon: account || draft.beantragtVon || signedInEmail,
                    lehrerEmail: draft.lehrerEmail || account || signedInEmail,
                    authorEmail: account || signedInEmail
                },
                {
                    teacherLabels: labels.lehrer,
                    accountEmail: account || signedInEmail || state.accountEmail,
                    teachers: (state.stammdaten && state.stammdaten.teachers) || []
                }
            );
            toast(
                asDraft
                    ? 'Entwurf gespeichert.'
                    : submitStatus === 'kv_stellungnahme'
                      ? 'Eingereicht – wartet auf Prüfung.'
                      : 'Zur Genehmigung eingereicht.'
            );
        }
        const savedTitle = draft.titel || 'Aktivität';
        const savedId =
            (createdLocal && (createdLocal.aktivitaetId || createdLocal.itemId)) ||
            (prev && (prev.aktivitaetId || prev.itemId)) ||
            draft.aktivitaetId ||
            '';
        state.editingItemId = null;
        state.form = emptyForm(prefillTeacher());
        state.formStep = 1;
        state.formAttempted = false;
        state.formTouched = {};
        state.multiDay = false;
        state.weitereAngabenOpen = false;
        if (asDraft) {
            state.submitSuccess = null;
            state.view = 'liste';
        } else {
            state.view = 'dashboard';
            state.submitSuccess = { id: savedId, titel: savedTitle };
            state.detailId = null;
        }
        await refreshData();
        // Reload leer / ohne Ownership → eben erzeugten Eintrag lokal behalten
        if (createdLocal && createdLocal.aktivitaetId) {
            const id = String(createdLocal.aktivitaetId);
            const has = (state.items || []).some((it) => String(it.aktivitaetId || '') === id);
            if (!has) {
                state.items = [createdLocal, ...(state.items || [])];
            }
        }
        paint();
    } catch (e) {
        state.loading = false;
        throw e;
    }
}

async function decide(itemId, status, ablehnungsGrund, rueckfrageText) {
    const it = state.items.find((x) => x.itemId === itemId);
    if (!it) throw new Error('Eintrag nicht gefunden.');
    const now = new Date().toISOString();
    const patch = {
        ...it,
        status,
        ablehnungsGrund: status === 'abgelehnt' ? ablehnungsGrund || '' : it.ablehnungsGrund || '',
        rueckfrageText:
            status === 'nachforderung' ? String(rueckfrageText || '').trim() : it.rueckfrageText || ''
    };
    if (status === 'genehmigt' || status === 'abgelehnt') {
        patch.genehmigtVon = state.accountEmail;
        patch.genehmigtAm = now;
    }
    if (status === 'nachforderung') {
        patch.aenderungsLog = appendAenderungsLog(
            it.aenderungsLog,
            'Zur Ergänzung / Revision: ' + String(rueckfrageText || '').trim()
        );
        patch.untisErledigt = false;
        patch.untisErledigtVon = '';
        patch.untisErledigtAm = '';
    }
    if (status === 'abgesagt') {
        patch.aenderungsLog = appendAenderungsLog(it.aenderungsLog, 'Als abgesagt markiert');
    }
    const rules = state.rules || DEFAULT_RULES;
    if (status === 'genehmigt') {
        const check = validateAktivitaet({
            draft: patch,
            existing: state.items,
            rules: {
                minVorlaufTage: 0,
                maxGleichzeitigProKlasse: rules.maxGleichzeitigProKlasse
            }
        });
        if (!check.ok) throw new Error(check.errors[0] || 'Genehmigung nicht möglich.');
    }

    const toastMsg =
        status === 'genehmigt'
            ? 'Genehmigt'
            : status === 'nachforderung'
              ? 'Zur Ergänzung zurückgegeben'
              : status === 'abgesagt'
                ? 'Als abgesagt markiert'
                : 'Abgelehnt';

    if (state.localDemoOnly || !state.ctx) {
        state.items = state.items.map((row) => (row.itemId === itemId ? { ...row, ...patch } : row));
        toast(toastMsg + ' (Demo).');
        state.detailId = null;
        state.localDemoOnly = true;
        paint();
        return;
    }

    await updateAktivitaetItem(state.ctx, itemId, patch);
    toast(toastMsg + '.');
    state.detailId = null;
    await refreshData();
}

async function markUntisDone(itemId) {
    const it = state.items.find((x) => x.itemId === itemId);
    if (!it) throw new Error('Eintrag nicht gefunden.');
    if (String(it.status || '').toLowerCase() !== 'genehmigt') {
        throw new Error('Nur genehmigte Anträge können in Untis abgehakt werden.');
    }
    const now = new Date().toISOString();
    const patch = {
        ...it,
        untisErledigt: true,
        untisErledigtVon: state.accountEmail,
        untisErledigtAm: now,
        aenderungsLog: appendAenderungsLog(it.aenderungsLog, 'In Untis verarbeitet markiert')
    };

    if (state.localDemoOnly || !state.ctx) {
        state.items = state.items.map((row) => (row.itemId === itemId ? { ...row, ...patch } : row));
        toast('Als in Untis verarbeitet markiert (Demo).');
        state.detailId = null;
        state.localDemoOnly = true;
        paint();
        return;
    }

    await updateAktivitaetItem(state.ctx, itemId, patch);
    toast('Als in Untis verarbeitet markiert.');
    state.detailId = null;
    await refreshData();
}

/** Klassenvorstand: Stellungnahme → weiter an Direktion (Status beantragt). */
async function submitKvStellungnahme(itemId, text) {
    const it = state.items.find((x) => x.itemId === itemId);
    if (!it) throw new Error('Eintrag nicht gefunden.');
    if (String(it.status || '').toLowerCase() !== 'kv_stellungnahme') {
        throw new Error('Keine KV-Stellungnahme für diesen Status.');
    }
    const now = new Date().toISOString();
    const note = String(text || '').trim();
    const patch = {
        ...it,
        status: 'beantragt',
        kvStellungnahme: note,
        kvStellungnahmeVon: state.accountEmail,
        kvStellungnahmeAm: now,
        aenderungsLog: appendAenderungsLog(it.aenderungsLog, 'KV-Stellungnahme abgegeben')
    };

    if (state.localDemoOnly || !state.ctx) {
        state.items = state.items.map((row) => (row.itemId === itemId ? { ...row, ...patch } : row));
        toast('KV-Stellungnahme gespeichert – weiter an Direktion (Demo).');
        state.detailId = null;
        state.localDemoOnly = true;
        paint();
        return;
    }

    await updateAktivitaetItem(state.ctx, itemId, patch);
    toast('KV-Stellungnahme gespeichert – weiter an Direktion.');
    state.detailId = null;
    await refreshData();
}

async function cancelItem(itemId) {
    const it = state.items.find((x) => x.itemId === itemId);
    if (!it || !canEditItem(it, state)) throw new Error('Absagen nicht erlaubt.');
    if (!window.confirm('Aktivität als abgesagt markieren? (Empfohlen statt Löschen.)')) return;
    await decide(itemId, 'abgesagt');
}

/**
 * Genehmigte Aktivität → persönlicher Outlook-Kalender (optional mit Einladungen).
 * @param {string} id
 * @param {{ invite?: boolean }} [opts]
 */
async function putAktInMyCalendar(id, opts) {
    const it = (state.items || []).find((x) => x.aktivitaetId === id || x.itemId === id);
    if (!it) throw new Error('Aktivität nicht gefunden.');
    const accountEmail = await resolveSignedInEmail();
    if (!accountEmail || accountEmail.indexOf('@') < 0) {
        throw new Error('Bitte anmelden – Outlook-Kalender braucht Ihre Microsoft-Konto-E-Mail.');
    }
    const invite = !!(opts && opts.invite);
    const result = await upsertAktToMyCalendar(it, {
        accountEmail,
        stammdaten: state.stammdaten,
        invite
    });
    const warn =
        result.warnings && result.warnings.length
            ? ' Hinweis: ' + result.warnings.slice(0, 2).join(' ')
            : '';
    if (invite) {
        toast(
            (result.created ? 'Termin angelegt' : 'Termin aktualisiert') +
                ' und ' +
                result.invited +
                ' Einladung(en) gesendet.' +
                warn
        );
    } else {
        toast(
            (result.created ? 'Im Outlook-Kalender eingetragen.' : 'Outlook-Termin aktualisiert.') +
                warn
        );
    }
    paint();
}

async function deleteItem(itemId) {
    const it = state.items.find((x) => x.itemId === itemId);
    if (!it || !canEditItem(it, state)) throw new Error('Löschen nicht erlaubt.');
    const st = String(it.status || '').toLowerCase();
    if (st === 'genehmigt' || st === 'beantragt' || st === 'kv_stellungnahme' || st === 'nachforderung') {
        toast('Bitte „Als abgesagt markieren“ verwenden – physisches Löschen nur bei Entwurf/Abgelehnt.');
        return;
    }
    if (!window.confirm('Antrag wirklich unwiderruflich löschen?')) return;

    if (state.localDemoOnly || !state.ctx) {
        state.items = state.items.filter((x) => x.itemId !== itemId);
        toast('Gelöscht (Demo).');
        state.detailId = null;
        paint();
        return;
    }

    await deleteAktivitaetItem(state.ctx, itemId);
    toast('Gelöscht.');
    state.detailId = null;
    await refreshData();
}

async function saveRules() {
    const patch = readRulesForm();
    if (!patch) return;

    if (state.localDemoOnly || !state.ctx || !(state.rules && state.rules.itemId)) {
        state.rules = {
            ...(state.rules || { ...DEFAULT_RULES }),
            ...patch,
            kostenLimits: (state.rules && state.rules.kostenLimits) || [],
            itemId: (state.rules && state.rules.itemId) || 'local-rw'
        };
        state.localDemoOnly = true;
        state.rulesDirty = false;
        state.settingsSavedFlash = 'regeln';
        paint();
        return;
    }

    await updateRegelwerkItem(state.ctx, state.rules.itemId, {
        ...patch,
        kostenLimits: (state.rules && state.rules.kostenLimits) || []
    });
    state.rulesDirty = false;
    state.settingsSavedFlash = 'regeln';
    await refreshData();
}

async function saveKostenLimits() {
    const limits = readKostenLimitsForm();
    if (!limits) return;
    const invalid = limits.find((r) => r.warnBis < r.okBis);
    if (invalid) {
        toast('„Rücksprache erforderlich bis“ muss ≥ „Ohne Rücksprache bis“ sein (' + invalid.typ + ').');
        return;
    }

    if (state.localDemoOnly || !state.ctx || !(state.rules && state.rules.itemId)) {
        state.rules = {
            ...(state.rules || { ...DEFAULT_RULES }),
            kostenLimits: limits,
            itemId: (state.rules && state.rules.itemId) || 'local-rw'
        };
        state.localDemoOnly = true;
        state.limitsDirty = false;
        state.settingsSavedFlash = 'limits';
        paint();
        return;
    }

    const patch = {
        minVorlaufTage: state.rules.minVorlaufTage,
        maxGleichzeitigProKlasse: state.rules.maxGleichzeitigProKlasse,
        vorlaufModus: state.rules.vorlaufModus,
        genehmigungPfad: state.rules.genehmigungPfad,
        aktiv: state.rules.aktiv !== false,
        kostenLimits: limits
    };
    await updateRegelwerkItem(state.ctx, state.rules.itemId, patch);
    state.limitsDirty = false;
    state.settingsSavedFlash = 'limits';
    await refreshData();
}

function applyLocalDemo(pack) {
    const local = buildLocalDemoState(pack);
    state.items = local.items;
    state.rules = local.rules;
    if (local.stammdaten) {
        state.stammdaten = {
            classes: local.stammdaten.classes || state.stammdaten.classes || [],
            teachers: local.stammdaten.teachers || state.stammdaten.teachers || []
        };
    }
    state.error = '';
    state.localDemoOnly = true;
    state.aktivitaetenLive = {
        ok: false,
        count: (state.items || []).length,
        source: 'demo',
        error: '',
        pulledAt: ''
    };
    state.view = 'dashboard';
    // IT (Kurt) bleibt Account – Ansicht läuft als Dave Grohl (Lehrer + KV).
    // Admin/KV/Lehrer-Chips nur mit ?demoRole=1.
    ensureDemoPersonaInStammdaten();
    state.demoPersonaEmail = DEMO_TEACHER_PERSONA.email;
    state.demoPersonaName = DEMO_TEACHER_PERSONA.name;
    state.demoRoleOverride = readDemoRoleOverrideFromUrl(false);
    applyDemoTeacherPersona();
    if (!state.siteUrl && local.siteDefault) {
        state.siteUrl = local.siteDefault;
        persistSiteUrl(state.siteUrl);
    }
    // Kalender auf „heute“, damit Showcase-Termine (isoAdd) sichtbar sind
    const now = new Date();
    state.calYear = now.getFullYear();
    state.calMonth = now.getMonth() + 1;
    const day = (now.getDay() + 6) % 7;
    const monday = new Date(now);
    monday.setHours(0, 0, 0, 0);
    monday.setDate(now.getDate() - day);
    state.calWeekAnchor = monday.toISOString();
}

/**
 * @param {object} pack
 */
async function runDemoImport(pack) {
    const p = parseDemoImportJson(pack);
    const stamOk = applyDemoStammdatenLocal(p.stammdaten);
    applyLocalDemo(p);
    state.stammdaten = loadStammdaten();
    if (!(state.stammdaten.classes || []).length && p.stammdaten) {
        state.stammdaten = {
            classes: p.stammdaten.classes || [],
            teachers: p.stammdaten.teachers || []
        };
    }
    paint();

    const n = (p.counts && p.counts.aktivitaeten) || (p.aktivitaeten || []).length;
    toast('Demo geladen (' + n + ' Aktivitäten)' + (stamOk ? ', Stammdaten übernommen' : '') + '.');

    const writeSp = window.confirm(
        'Demo lokal geladen (' +
            n +
            ' Einträge, SJ ' +
            (p.schoolYear || '2026/27') +
            ').\n\nJetzt auch auf SharePoint schreiben?\n\nSite: ' +
            (state.siteUrl || DEMO_SITE_DEFAULT) +
            '\n\n(Listen müssen existieren – sonst zuerst „Listen“ anlegen.)'
    );
    if (!writeSp) return;

    const webUrl = state.siteUrl || DEMO_SITE_DEFAULT;
    state.siteUrl = webUrl;
    persistSiteUrl(webUrl);
    state.loading = true;
    state.error = '';
    paint();
    try {
        await seedDemoSchulaktivitaeten(webUrl, (msg) => console.log('[akt-demo]', msg), { pack: p });
        state.localDemoOnly = false;
        await refreshData();
        toast('Demo auf SharePoint geschrieben und neu geladen.');
    } catch (e) {
        state.localDemoOnly = true;
        // Roh-403 nicht in die Banner-Zeile – lokal bleibt die Demo voll nutzbar.
        state.error = '';
        toast(
            'Demo bleibt lokal (SharePoint-Schreiben nicht möglich). Für die Vorführung reicht das.'
        );
        console.warn('[akt-demo] SharePoint seed failed', e);
        paint();
    } finally {
        state.loading = false;
        paint();
    }
}

async function runDemoReset() {
    const localOnly = state.localDemoOnly && !state.ctx;
    const msg = localOnly
        ? 'Lokale Demo-Daten leeren?'
        : 'Alle Demo-Einträge (Seed-Tag / akt-demo-*) auf SharePoint löschen?\n\nSite: ' +
          (state.siteUrl || DEMO_SITE_DEFAULT) +
          '\n\nEchte (nicht-Demo) Einträge bleiben erhalten.';
    if (!window.confirm(msg)) return;

    if (localOnly || !state.siteUrl) {
        state.items = [];
        state.rules = { ...DEFAULT_RULES, itemId: 'local-rw', title: 'Standard', regelwerkId: 'akt-rw-1', aktiv: true };
        state.localDemoOnly = false;
        state.aktivitaetenLive = { ok: false, count: 0, source: '', error: '', pulledAt: '' };
        state.demoPersonaEmail = '';
        state.demoPersonaName = '';
        state.error = '';
        paint();
        toast('Lokale Demo geleert.');
        return;
    }

    state.loading = true;
    state.error = '';
    paint();
    try {
        const result = await resetDemoSchulaktivitaeten(
            state.siteUrl || DEMO_SITE_DEFAULT,
            (msgLine) => console.log('[akt-demo-reset]', msgLine)
        );
        state.localDemoOnly = false;
        await refreshData();
        toast('Demo zurückgesetzt (' + (result.deleted || 0) + ' gelöscht).');
    } catch (e) {
        state.error = e && e.message ? e.message : String(e);
        toast(state.error);
        paint();
    } finally {
        state.loading = false;
        paint();
    }
}

/**
 * Lehrer-Browser oft ohne lokale IT-Einrichtung: Site-URL aus Lizenz übernehmen.
 *
 * Standard ist still (nur Token-Cache): Beim Laden der Seite, bei Konto-Wechsel und beim
 * Blättern darf der Planer nicht von selbst zur Microsoft-Anmeldung springen. Nur ein
 * ausdrücklicher Klick (`interactive: true`) darf eine Anmeldung auslösen.
 * @param {{ interactive?: boolean }} [opts]
 * @returns {Promise<string>}
 */
async function ensureAktSiteFromLicense(opts) {
    const interactive = !!(opts && opts.interactive);
    const local = String(state.siteUrl || '').trim() || loadSavedSiteUrl();
    if (local) {
        state.siteUrl = local;
        return local;
    }
    try {
        const api = window.ms365LicenseApi;
        if (!api || typeof api.fetchLicenseMe !== 'function') return '';
        let token = '';
        if (typeof api.acquireLicenseToken === 'function') {
            token = await api.acquireLicenseToken(interactive ? undefined : { silentOnly: true });
        } else if (interactive && typeof window.ms365AuthAcquireIdToken === 'function') {
            token = await window.ms365AuthAcquireIdToken();
        }
        if (!token) return '';
        const result = await api.fetchLicenseMe(token);
        if (!result || result.allowed !== true) return '';
        const url = toolSiteUrlFromLicenseResult(result);
        if (!url) {
            state.roleHint =
                'Keine Schul-Site in der Lizenz (Tool-Site-URL). IT: unter Admin → Schule die SharePoint-URL der MS365-Schultools-Site setzen – sonst weiß der Lehrer-Browser nicht, welche Site gemeint ist.';
            return '';
        }
        applyToolSiteUrlLocally(url);
        try {
            await bootstrapSchoolSiteFromLicense(result, { force: false });
        } catch {
            /* optional */
        }
        state.siteUrl = url;
        persistSiteUrl(url);
        return url;
    } catch {
        return '';
    }
}

/**
 * Effektive Site-URL: lokal → Lizenz. Immer bevor „keine URL“ gemeldet wird.
 * @param {{ interactive?: boolean }} [opts] siehe ensureAktSiteFromLicense
 * @returns {Promise<string>}
 */
async function ensureAktSiteUrl(opts) {
    let url = String(state.siteUrl || '').trim() || loadSavedSiteUrl();
    if (url) {
        state.siteUrl = url;
        return url;
    }
    url = await ensureAktSiteFromLicense(opts);
    if (url) return url;
    url = loadSavedSiteUrl();
    if (url) {
        state.siteUrl = url;
        return url;
    }
    return '';
}

function boot() {
    root = document.getElementById('aktApp');
    if (!root) return;
    state.demoRoleOverride = readDemoRoleOverrideFromUrl(false);
    state.role = resolveRole();
    state.siteUrl = loadSavedSiteUrl() || state.siteUrl;
    state.stammdaten = loadStammdaten();
    state.form = emptyForm(prefillTeacher());
    window.addEventListener('ms365-auth-state-changed', () => {
        void (async () => {
            await syncAccountAndRole({ preferShellDefault: true });
            paint();
            // Nach Konto-Wechsel Site aus Lizenz holen und Klassen neu lesen
            const url = await ensureAktSiteUrl();
            if (url) {
                try {
                    await refreshData();
                } catch {
                    paint();
                }
            } else {
                paint();
            }
        })();
    });
    window.addEventListener('ms365-it-shell-changed', () => {
        void onItShellChanged();
    });
    window.addEventListener('ms365-license-changed', (ev) => {
        const detail = (ev && ev.detail) || {};
        if (!detail.allowed || !detail.result) return;
        const url = toolSiteUrlFromLicenseResult(detail.result);
        if (!url) return;
        applyToolSiteUrlLocally(url);
        state.siteUrl = url;
        persistSiteUrl(url);
        void refreshData().catch(() => {});
    });
    paint();
    void (async () => {
        await syncAccountAndRole({ preferShellDefault: true });
        const url = await ensureAktSiteUrl();
        paint();
        if (url) {
            refreshData().catch(() => {
                /* error already in state */
            });
        } else {
            state.roleHint =
                state.roleHint ||
                'SharePoint-Site fehlt in diesem Browser. IT: Tool-Site-URL in der Lizenz setzen (Admin → Schule) und IT-Sicherung ablegen – dann als Lehrkraft neu laden.';
            paint();
        }
    })();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
