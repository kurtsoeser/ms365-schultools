/**
 * Entry / Wiring: sRDP-Anmeldung App (Schulaktivitäten-Niveau).
 */
import {
    createInitialState,
    loadStammdaten,
    loadWizardConfig,
    persistWizardConfig,
    matchTeacherByEmail,
    teacherDisplayLabel,
    matchStudentByEmail,
    classGroupLinksFromAppData,
    kvKlasseCodesForEmail,
    persistRole,
    persistSiteUrl,
    loadSavedSiteUrl,
    resolveFormOptions,
    looksLikeDemoKlassen,
    looksLikeDemoSeminare,
    terminJahrFromListTitle,
    emptyForm,
    viewsForRole,
    applyFilterJump,
    joinCodeList,
    effectiveAccountEmail
} from './srdp-anmeldung-state.js';
import {
    resolveSrdpContext,
    loadSrdpListChoiceOptions,
    loadSrdpKlasseChoicesViaSpo,
    loadSrdpLehrerChoicesViaSpo,
    loadAllAnmeldungen,
    createAnmeldungItem,
    updateAnmeldungNacharbeit,
    fetchMyMemberGroups
} from './srdp-anmeldung-graph.js';
import {
    validateAnmeldungDraft,
    buildPruefplanKurz,
    buildPruefplanDetail,
    isSeminarWahlfach,
    resolveSrdpOwnership,
    assertSrdpOwnership,
    isSrdpAnmeldeFensterOffen,
    findAnmeldungBySchuelerEmail,
    splitPersonName,
    filterAbschlussklassen,
    resolveStudentAssignedKlasse,
    looksLikeDemoLfs,
    wizardConfiguredChoices
} from './srdp-anmeldung-logic.js';
import { refreshPruefungstermineFromWeb } from './srdp-pruefungstermine.js';
import {
    readDemoRoleOverrideFromUrl,
    inferSrdpRolesFromStammdaten,
    shellPreferredSrdpRole,
    resolveSrdpRolesForAccount
} from './srdp-anmeldung-entra-role.js';
import { renderApp, readFormFromDom, syncSrdpPageChrome, firstCalendarExamDate } from './srdp-anmeldung-ui.js';
import {
    buildLocalDemoState,
    buildDemoWizard,
    buildDemoStammdaten,
    DEMO_TEACHER_PERSONA,
    DEMO_STUDENT_PERSONA,
    DEMO_SEED_TAG,
    stripDemoSeedItems,
    isDemoSeedItem
} from './srdp-anmeldung-demo.js';
import { hydrateStammdatenFromLiveSite } from '../../shared/school-stammdaten-spo-lists.js';
import { announcePlannerState } from '../../shared/app-role-chrome.js';
import { toIsoDateOnly, startOfWeekMonday } from '../../shared/ms365-calendar-shell.js';

const state = createInitialState();
let root = null;

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function paint() {
    announcePlannerState('srdp', state);
    renderApp(state, root);
}

function refreshChrome() {
    import('../../shared/frontend-planner-chrome-policy.js')
        .then((m) => {
            if (m && typeof m.refreshFrontendPlannerChromeAudience === 'function') {
                return m.refreshFrontendPlannerChromeAudience();
            }
        })
        .then(() => syncSrdpPageChrome(state))
        .catch(() => syncSrdpPageChrome(state));
}

function refreshDerived() {
    if (!state.localDemoOnly) {
        const sd = loadStammdaten();
        const cur = state.stammdaten || { teachers: [], classes: [], students: [] };
        state.stammdaten = {
            teachers:
                Array.isArray(cur.teachers) && cur.teachers.length
                    ? cur.teachers
                    : sd.teachers || [],
            classes:
                Array.isArray(cur.classes) && cur.classes.length ? cur.classes : sd.classes || [],
            students:
                Array.isArray(cur.students) && cur.students.length
                    ? cur.students
                    : sd.students || [],
            subjects:
                Array.isArray(cur.subjects) && cur.subjects.length
                    ? cur.subjects
                    : sd.subjects || []
        };
        const wiz = loadWizardConfig();
        if (wiz && Object.keys(wiz).length) state.wizard = Object.assign({}, state.wizard, wiz);
    }
    state.options = resolveFormOptions(state.wizard, state.stammdaten, state.wizard.profileId);
    const email = effectiveAccountEmail(state);
    const teacher = matchTeacherByEmail(state.stammdaten.teachers, email);
    if (teacher) state.teacherLabel = teacherDisplayLabel(teacher);
    const kv = kvKlasseCodesForEmail(state.stammdaten, email);
    if (kv.length) state.kvKlassen = kv;
}

/**
 * Konto robust lesen (MSAL kann beim ersten Paint noch nicht fertig sein).
 */
async function readAccount() {
    let email = '';
    let name = '';
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            if (info) {
                email = String(
                    info.upn || info.username || info.mail || info.email || ''
                ).trim();
                name = String(info.name || '').trim();
            }
        }
    } catch {
        /* ignore */
    }
    if (!email) {
        try {
            if (typeof window.ms365AuthGetUserPrincipalName === 'function') {
                email = String(window.ms365AuthGetUserPrincipalName() || '').trim();
            }
        } catch {
            /* ignore */
        }
    }
    if (!email) {
        try {
            if (typeof window.ms365MsalGetAccount === 'function') {
                const acc = await window.ms365MsalGetAccount();
                if (acc) {
                    email = String(acc.username || acc.mail || '').trim();
                    if (!name) name = String(acc.name || '').trim();
                }
            }
        } catch {
            /* ignore */
        }
    }
    state.accountEmail = email.toLowerCase();
    if (name) state.accountName = name;
}

/**
 * Demo-Schüler/Lehrer in Stammdaten + Rollen neu ableiten (nach Auth-Ready).
 */
async function syncAccountAndRole() {
    await readAccount();
    try {
        const { ensureSignedInMs365DemoPortalAccountInLocalStammdaten } = await import(
            '../../shared/demo-portal-teacher.js'
        );
        ensureSignedInMs365DemoPortalAccountInLocalStammdaten();
    } catch {
        /* ignore */
    }
    refreshDerived();
    ensureKlassenOptions();
    if (state.demoRoleOverride) {
        state.planerRoles = ['schueler', 'lehrer', 'admin'];
        state.planerRoleSources = { schueler: 'demo', lehrer: 'demo', admin: 'demo' };
        state.role = shellPreferredSrdpRole(state.planerRoles, state.role, { allowSaved: true });
    } else {
        let resolved = null;
        try {
            const liveSd = loadStammdaten();
            const probe = state.localDemoOnly
                ? Object.assign({}, state, {
                      stammdaten: {
                          teachers: liveSd.teachers || [],
                          classes: liveSd.classes || [],
                          students: liveSd.students || []
                      }
                  })
                : state;
            resolved = await resolveSrdpRolesForAccount(probe);
        } catch {
            resolved = null;
        }
        const fallback = inferSrdpRolesFromStammdaten(state);
        const active = (resolved && resolved.active) || fallback[0] || 'schueler';
        state.planerRoles = [active];
        state.planerRoleSources = (resolved && resolved.sources) || {};
        state.role = active;
    }
    if (state.accountEmail || state.demoRoleOverride) persistRole(state.role);
    if (state.role === 'schueler') prefillStudentForm();
    ensureViewForRole();
}

/**
 * Leeres Klassen-Dropdown verhindern.
 * Demo-Klassen nur im lokalen Demo-Modus – nie in Live-Wizard-Storage schreiben.
 */
/** Demo-Klassen und Demo-Fristen aus dem Live-Wizard entfernen (alter Browser-Stand). */
function purgeStickyDemoWizard() {
    if (state.localDemoOnly) return;
    const wiz = state.wizard || {};
    const patch = {};
    if (looksLikeDemoKlassen(wiz.klassen)) {
        patch.klassen = [];
        patch.anmeldeFrist = '';
        patch.haupttermin = '';
    }
    if (looksLikeDemoSeminare(wiz.seminarText) || looksLikeDemoSeminare(wiz.seminar)) {
        patch.seminarText = '';
        patch.seminar = '';
    }
    if (looksLikeDemoLfs(wiz.lfsText) || looksLikeDemoLfs(wiz.lfs)) {
        patch.lfsText = '';
        patch.lfs = '';
    }
    if (!Object.keys(patch).length) return;
    state.wizard = Object.assign({}, wiz, patch);
    persistWizardConfig(patch);
    state.options = resolveFormOptions(state.wizard, state.stammdaten, state.wizard.profileId);
}

function ensureKlassenOptions() {
    const cur = state.options || resolveFormOptions(state.wizard, state.stammdaten, state.wizard.profileId);
    if (Array.isArray(cur.klassen) && cur.klassen.length) {
        state.options = cur;
        return;
    }
    let wiz = state.wizard || {};
    if (state.localDemoOnly && (!Array.isArray(wiz.klassen) || !wiz.klassen.length)) {
        try {
            const demo = buildDemoWizard();
            wiz = Object.assign({}, wiz, {
                klassen: demo.klassen.slice(),
                jahr: wiz.jahr || wiz.terminJahr || demo.jahr,
                terminJahr: wiz.terminJahr || wiz.jahr || demo.terminJahr,
                lfsText: wiz.lfsText || demo.lfsText,
                wahlfachText: wiz.wahlfachText || demo.wahlfachText,
                seminarText: wiz.seminarText || demo.seminarText,
                anmeldeFrist: wiz.anmeldeFrist || demo.anmeldeFrist,
                haupttermin: wiz.haupttermin || demo.haupttermin,
                profileId: wiz.profileId || demo.profileId
            });
            state.wizard = wiz;
            // absichtlich nicht persistieren – Demo bleibt session-lokal
        } catch {
            /* ignore */
        }
    }
    if (
        state.localDemoOnly &&
        (!state.stammdaten || !(state.stammdaten.classes || []).length)
    ) {
        try {
            const sd = buildDemoStammdaten();
            state.stammdaten = Object.assign({}, state.stammdaten || {}, sd);
        } catch {
            /* ignore */
        }
    }
    state.options = resolveFormOptions(state.wizard, state.stammdaten, state.wizard.profileId, {
        allowDemoKlassen: !!state.localDemoOnly
    });
}

/**
 * Nach erfolgreichem SharePoint-Load: Demo-Pollution aus Wizard entfernen,
 * Klassen aus Stammdaten bzw. Listen-Choices übernehmen.
 */
function scrubDemoWizardPollution(listChoices) {
    const choices = listChoices || {};
    const jahr = String(
        (state.options && state.options.jahr) ||
            state.wizard.jahr ||
            state.wizard.terminJahr ||
            ''
    ).trim();
    const fromSd = filterAbschlussklassen((state.stammdaten && state.stammdaten.classes) || [], jahr);
    const fromList = Array.isArray(choices.klassen) ? choices.klassen : [];
    const patch = {};
    if (fromSd.length) patch.klassen = fromSd.slice();
    else if (fromList.length) patch.klassen = fromList.slice();
    else if (looksLikeDemoKlassen(state.wizard && state.wizard.klassen)) patch.klassen = [];

    if (Array.isArray(choices.lfs) && choices.lfs.length) {
        patch.listLfs = choices.lfs.slice();
        const configured = wizardConfiguredChoices(state.wizard, 'lfsText', 'lfs');
        if (!configured.length || looksLikeDemoLfs(configured)) {
            patch.lfsText = choices.lfs.join('\n');
            patch.lfs = patch.lfsText;
        }
    }
    if (Array.isArray(choices.wahlfaecher) && choices.wahlfaecher.length) {
        patch.listWahlfaecher = choices.wahlfaecher.slice();
        const configured = wizardConfiguredChoices(state.wizard, 'wahlfachText', 'wahlfach');
        if (!configured.length) {
            patch.wahlfachText = choices.wahlfaecher.join('\n');
            patch.wahlfach = patch.wahlfachText;
        }
    }
    if (Array.isArray(choices.seminare) && choices.seminare.length) {
        patch.listSeminare = choices.seminare.slice();
        const configured = wizardConfiguredChoices(state.wizard, 'seminarText', 'seminar');
        if (!configured.length || looksLikeDemoSeminare(configured)) {
            patch.seminarText = choices.seminare.join('\n');
            patch.seminar = patch.seminarText;
        }
    }
    if (fromList.length) patch.listKlassen = fromList.slice();
    if (Array.isArray(choices.lehrer) && choices.lehrer.length) {
        patch.listLehrer = choices.lehrer.slice();
    }
    if (choices.bfkKnown) patch.bfkDatum = choices.bfkDatum || '';

    // Demo-Fristen (relativ zu „heute“) nicht als Live-Frist behalten
    if (looksLikeDemoKlassen((state.wizard && state.wizard.klassen) || []) || patch.klassen) {
        const demo = (() => {
            try {
                return buildDemoWizard();
            } catch {
                return null;
            }
        })();
        if (demo && state.wizard.anmeldeFrist === demo.anmeldeFrist) patch.anmeldeFrist = '';
        if (demo && state.wizard.haupttermin === demo.haupttermin) patch.haupttermin = '';
    }

    if (!Object.keys(patch).length) return;
    state.wizard = Object.assign({}, state.wizard, patch);
    persistWizardConfig(patch);
    state.options = resolveFormOptions(state.wizard, state.stammdaten, state.wizard.profileId);
}

function ensureViewForRole() {
    const allowed = viewsForRole(state.role, { kvKlassen: state.kvKlassen }).map((v) => v.id);
    if (state.view === 'einrichtung') state.view = 'system';
    if (!allowed.includes(state.view)) {
        state.view = allowed[0] || (state.role === 'schueler' ? 'anmelden' : 'dashboard');
    }
}

function prefillStudentForm() {
    if (state.role !== 'schueler') return;
    const email = effectiveAccountEmail(state);
    const students = (state.stammdaten && state.stammdaten.students) || [];
    const classes = (state.stammdaten && state.stammdaten.classes) || [];
    const st = matchStudentByEmail(students, email);
    const seed = { wizardStep: state.form && state.form.wizardStep };
    if (st) {
        const fullName = String(st.name || st.displayName || '').trim();
        const splitFromName = fullName ? splitPersonName(fullName) : { vorname: '', nachname: '' };
        seed.nachname = st.nachname || st.lastName || splitFromName.nachname || '';
        seed.vorname = st.vorname || st.firstName || splitFromName.vorname || '';
    } else if (state.demoPersonaName || state.accountName) {
        const split = splitPersonName(state.demoPersonaName || state.accountName);
        seed.vorname = split.vorname;
        seed.nachname = split.nachname;
    }
    seed.klasse = resolveStudentAssignedKlasse({
        email,
        students,
        classes,
        classLinks: classGroupLinksFromAppData(),
        memberGroups: state.memberGroups || [],
        preferredCodes: (state.options && state.options.klassen) || []
    });
    const cur = state.form || emptyForm();
    state.form = emptyForm({
        ...cur,
        // Konto/Stammdaten haben Vorrang; sonst bisherige Eingabe behalten.
        nachname: seed.nachname || cur.nachname || '',
        vorname: seed.vorname || cur.vorname || '',
        klasse: seed.klasse || cur.klasse || '',
        wizardStep: cur.wizardStep || seed.wizardStep || 1
    });
}

let memberGroupFlight = null;

/** Gruppenmitgliedschaft neu lesen und die Klasse daraus setzen. Kein Abmelden. */
function refreshStudentClassFromGroups() {
    if (state.role !== 'schueler' || state.localDemoOnly) return Promise.resolve();
    if (state.memberGroupsLoaded) {
        prefillStudentForm();
        return Promise.resolve();
    }
    if (memberGroupFlight) return memberGroupFlight;
    memberGroupFlight = fetchMyMemberGroups()
        .then((groups) => {
            state.memberGroupTries = (state.memberGroupTries || 0) + 1;
            if (groups.length) state.memberGroups = groups;
            if (groups.length || state.memberGroupTries >= 2) state.memberGroupsLoaded = true;
            const before = String((state.form && state.form.klasse) || '');
            prefillStudentForm();
            const after = String((state.form && state.form.klasse) || '');
            if (after && after !== before) paint();
        })
        .catch(() => {
            state.memberGroupTries = (state.memberGroupTries || 0) + 1;
            if (state.memberGroupTries >= 2) state.memberGroupsLoaded = true;
        })
        .finally(() => {
            memberGroupFlight = null;
        });
    return memberGroupFlight;
}

function focusCalendarOnExams() {
    if (state.calAnchored) return;
    const iso = firstCalendarExamDate(state);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return;
    const parts = iso.split('-').map((n) => Number(n));
    state.calYear = parts[0];
    state.calMonth = parts[1];
    state.calWeekAnchor = iso;
    state.calMode = 'month';
}

function alignCalendarToToday() {
    const today = toIsoDateOnly(new Date());
    state.calWeekAnchor = today;
    const d = new Date();
    state.calYear = d.getFullYear();
    state.calMonth = d.getMonth() + 1;
    state.calMode = 'month';
}

/**
 * @param {string} siteId
 */
async function hydrateLiveStammdaten(siteId) {
    if (!siteId || state.localDemoOnly) {
        state.stammdatenLive = {
            ok: !!state.localDemoOnly,
            error: '',
            classCount: ((state.stammdaten && state.stammdaten.classes) || []).length,
            teacherCount: ((state.stammdaten && state.stammdaten.teachers) || []).length,
            missing: [],
            source: state.localDemoOnly ? 'demo' : 'cache'
        };
        return;
    }
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api || typeof api.getGraphToken !== 'function') {
        state.stammdatenLive = {
            ok: false,
            error: 'Graph nicht geladen',
            classCount: ((state.stammdaten && state.stammdaten.classes) || []).length,
            teacherCount: ((state.stammdaten && state.stammdaten.teachers) || []).length,
            missing: [],
            source: 'cache'
        };
        return;
    }
    try {
        const tok = await api.getGraphToken([
            'https://graph.microsoft.com/User.Read',
            'https://graph.microsoft.com/Sites.Read.All'
        ]);
        const result = await hydrateStammdatenFromLiveSite({
            token: tok,
            siteId: String(siteId),
            siteWebUrl: String(state.siteUrl || '').trim(),
            graph: api,
            local: loadStammdaten()
        });
        state.stammdaten = result.stammdaten;
        state.stammdatenLive = {
            ok: true,
            error: '',
            classCount: (result.stammdaten.classes || []).length,
            teacherCount: (result.stammdaten.teachers || []).length,
            missing: result.missing || [],
            source: 'live'
        };
        refreshDerived();
    } catch (e) {
        state.stammdatenLive = {
            ok: false,
            error: (e && e.message) || String(e),
            classCount: ((state.stammdaten && state.stammdaten.classes) || []).length,
            teacherCount: ((state.stammdaten && state.stammdaten.teachers) || []).length,
            missing: [],
            source: 'cache'
        };
        /* lokaler Stand bleibt – kein Wipe */
    }
}

function isAccessDeniedError(err) {
    const msg = String((err && err.message) || err || '');
    return /403|Forbidden|Zugriff verweigert|access denied/i.test(msg);
}

async function loadFromSharePoint() {
    const siteUrl = String(state.siteUrl || loadSavedSiteUrl() || '').trim();
    state.siteUrl = siteUrl;
    if (!siteUrl) {
        state.error = 'Keine SharePoint-Site – Demo starten oder unter Systemadministration eine URL setzen.';
        state.syncMeta = Object.assign({}, state.syncMeta, {
            source: state.localDemoOnly ? 'demo' : 'none',
            connected: false
        });
        return;
    }
    state.loading = true;
    state.error = '';
    state.softError = '';
    paint();
    try {
        const wiz = loadWizardConfig();
        const ctx = await resolveSrdpContext(siteUrl, {
            profileId: state.options.profileId || wiz.profileId || 'hak',
            jahr: state.options.jahr || wiz.jahr || wiz.terminJahr,
            listId: state.listId || wiz.listId || ''
        });
        state.siteId = ctx.siteId;
        state.listId = ctx.list.id;
        state.listWebUrl = ctx.list.webUrl || '';
        state.localDemoOnly = false;
        state.ctx = ctx;
        const titleYear = terminJahrFromListTitle(ctx.list && ctx.list.title);
        if (titleYear) {
            state.wizard = Object.assign({}, state.wizard, { jahr: titleYear, terminJahr: titleYear });
            state.options = resolveFormOptions(state.wizard, state.stammdaten, state.wizard.profileId);
        }
        persistWizardConfig({
            listId: ctx.list.id,
            siteUrl,
            jahr: state.options.jahr || wiz.jahr || wiz.terminJahr,
            terminJahr: state.options.jahr || wiz.terminJahr || wiz.jahr,
            profileId: state.options.profileId || wiz.profileId || 'hak'
        });

        await hydrateLiveStammdaten(ctx.siteId);

        let listChoices = {};
        try {
            const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
            const tok =
                api && typeof api.getGraphToken === 'function'
                    ? await api.getGraphToken([
                          'https://graph.microsoft.com/User.Read',
                          'https://graph.microsoft.com/Sites.ReadWrite.All'
                      ])
                    : null;
            if (tok) {
                listChoices = await loadSrdpListChoiceOptions(tok, ctx.siteId, ctx.list.id);
            }
        } catch {
            listChoices = {};
        }
        if (!listChoices.klassen || !listChoices.klassen.length) {
            try {
                const viaSpo = await loadSrdpKlasseChoicesViaSpo(siteUrl, ctx.list.id);
                if (viaSpo.length) listChoices = Object.assign({}, listChoices, { klassen: viaSpo });
            } catch {
                /* Stammdaten bleiben die Quelle */
            }
        }
        if (!listChoices.lehrer || !listChoices.lehrer.length) {
            try {
                const lehrerSpo = await loadSrdpLehrerChoicesViaSpo(siteUrl, ctx.list.id);
                if (lehrerSpo.length) listChoices = Object.assign({}, listChoices, { lehrer: lehrerSpo });
            } catch {
                /* unten Stammdaten oder leere Auswahl */
            }
        }
        scrubDemoWizardPollution(listChoices);

        state.items = await loadAllAnmeldungen(ctx);
        state.syncMeta = {
            source: 'sharepoint',
            lastPullAt: new Date().toISOString(),
            listTitle: (ctx.list && ctx.list.title) || '',
            connected: true
        };
        if (state.role === 'schueler') prefillStudentForm();
    } catch (e) {
        if (state.localDemoOnly || (state.items || []).some(isDemoSeedItem)) {
            state.localDemoOnly = true;
            state.error = '';
            state.softError = isAccessDeniedError(e)
                ? 'SharePoint-Zugriff eingeschränkt – lokale Demo bleibt erhalten.'
                : 'SharePoint nicht erreichbar – lokale Demo bleibt erhalten.';
            state.syncMeta = Object.assign({}, state.syncMeta, {
                source: 'demo',
                connected: false
            });
        } else if (isAccessDeniedError(e)) {
            state.error = '';
            state.softError =
                'Kein SharePoint-Zugriff – Site in der Systemadministration prüfen oder Demo lokal laden.';
            state.syncMeta = Object.assign({}, state.syncMeta, {
                source: state.syncMeta && state.syncMeta.source === 'sharepoint' ? 'cache' : 'none',
                connected: false
            });
            /* kein Wipe von Cache/Items */
        } else {
            state.error = (e && e.message) || String(e);
            state.syncMeta = Object.assign({}, state.syncMeta, { connected: false });
        }
    } finally {
        state.loading = false;
        paint();
    }
}

function enableDemo(opts) {
    const o = opts || {};
    const keepDemoRole = !!state.demoRoleOverride;
    const liveRole = state.role === 'admin' || state.role === 'lehrer' ? state.role : 'schueler';
    const demo = buildLocalDemoState(state);
    Object.assign(state, demo);
    if (keepDemoRole) {
        state.demoRoleOverride = true;
        state.planerRoles = ['schueler', 'lehrer', 'admin'];
        state.role = shellPreferredSrdpRole(state.planerRoles, liveRole, { allowSaved: true });
    } else {
        state.demoRoleOverride = false;
        state.planerRoles = [liveRole];
        state.role = liveRole;
    }
    state.syncMeta = {
        source: 'demo',
        lastPullAt: '',
        listTitle: '',
        connected: false
    };
    if (o.withPersona) {
        state.demoPersonaEmail = DEMO_TEACHER_PERSONA.email;
        state.demoPersonaName = DEMO_TEACHER_PERSONA.name;
        state.role = 'lehrer';
        if (!keepDemoRole) state.planerRoles = ['lehrer'];
        state.teacherLabel = DEMO_TEACHER_PERSONA.label;
        state.kvKlassen = kvKlasseCodesForEmail(state.stammdaten, DEMO_TEACHER_PERSONA.email);
    } else {
        state.role = keepDemoRole
            ? shellPreferredSrdpRole(state.planerRoles, state.role, { allowSaved: true })
            : liveRole;
        if (state.role === 'lehrer') {
            state.demoPersonaEmail = DEMO_TEACHER_PERSONA.email;
            state.demoPersonaName = DEMO_TEACHER_PERSONA.name;
            state.teacherLabel = DEMO_TEACHER_PERSONA.label;
            state.kvKlassen = kvKlasseCodesForEmail(state.stammdaten, DEMO_TEACHER_PERSONA.email);
        } else if (state.role === 'schueler') {
            state.demoPersonaEmail = DEMO_STUDENT_PERSONA.email;
            state.demoPersonaName = DEMO_STUDENT_PERSONA.name;
        } else {
            state.demoPersonaEmail = '';
            state.demoPersonaName = '';
        }
    }
    persistRole(state.role);
    refreshDerived();
    Object.assign(state, {
        wizard: demo.wizard,
        stammdaten: demo.stammdaten,
        items: demo.items,
        options: resolveFormOptions(demo.wizard, demo.stammdaten, 'hak', { allowDemoKlassen: true })
    });
    if (state.role === 'schueler') prefillStudentForm();
    alignCalendarToToday();
    ensureViewForRole();
    toast('Demo-Daten geladen (' + DEMO_SEED_TAG + ').');
    refreshChrome();
    paint();
}

function resetDemoSeed() {
    state.items = stripDemoSeedItems(state.items);
    const demo = buildLocalDemoState(state);
    state.items = (state.items || []).concat(demo.items);
    state.localDemoOnly = true;
    state.wizard = demo.wizard;
    state.stammdaten = demo.stammdaten;
    state.options = resolveFormOptions(demo.wizard, demo.stammdaten, 'hak', { allowDemoKlassen: true });
    state.stammdatenLive = demo.stammdatenLive;
    alignCalendarToToday();
    toast('Demo-Seed zurückgesetzt.');
    paint();
}

/**
 * Nacharbeit lokal umschalten und bei Live-Items nach SharePoint pushen.
 * Demo-Seeds bleiben local-first.
 * @param {string} id
 */
async function toggleNacharbeit(id) {
    const itemId = String(id || '').trim();
    if (!itemId) return;
    const prev = (state.items || []).find((it) => String(it.itemId) === itemId);
    if (!prev) return;
    const nextVal = !prev.nacharbeitErledigt;
    state.items = (state.items || []).map((it) => {
        if (String(it.itemId) !== itemId) return it;
        return Object.assign({}, it, { nacharbeitErledigt: nextVal });
    });
    paint();

    if (state.localDemoOnly || isDemoSeedItem(prev) || /^demo-/i.test(itemId)) {
        return;
    }
    try {
        if (!state.ctx) await loadFromSharePoint();
        if (!state.ctx) throw new Error(state.error || state.softError || 'Liste nicht geladen.');
        await updateAnmeldungNacharbeit(state.ctx, itemId, nextVal);
        toast(nextVal ? 'Prüfungsliste als erledigt gespeichert.' : 'Nacharbeit wieder geöffnet.');
    } catch (e) {
        state.items = (state.items || []).map((it) => {
            if (String(it.itemId) !== itemId) return it;
            return Object.assign({}, it, { nacharbeitErledigt: !!prev.nacharbeitErledigt });
        });
        if (isAccessDeniedError(e)) {
            state.softError =
                'Nacharbeit konnte in SharePoint nicht gespeichert werden (kein Schreibzugriff).';
        } else {
            state.softError = (e && e.message) || String(e);
        }
        paint();
    }
}

async function submitAnmeldung() {
    await syncAccountAndRole();
    const draft = readFormFromDom(state.form);
    // Name aus Konto/Stammdaten nachziehen, falls Formular nur Klasse zeigt.
    const seeded = resolveSrdpOwnership(draft, {
        accountEmail: effectiveAccountEmail(state),
        accountName: state.demoPersonaName || state.accountName,
        students: (state.stammdaten && state.stammdaten.students) || []
    });
    draft.nachname = draft.nachname || seeded.nachname;
    draft.vorname = draft.vorname || seeded.vorname;
    draft.klasse = draft.klasse || seeded.klasse;
    state.form = draft;
    const ownership = resolveSrdpOwnership(draft, {
        accountEmail: effectiveAccountEmail(state),
        accountName: state.demoPersonaName || state.accountName,
        students: (state.stammdaten && state.stammdaten.students) || []
    });
    try {
        assertSrdpOwnership(ownership);
    } catch (e) {
        state.formErrors = [(e && e.message) || String(e)];
        paint();
        return;
    }
    const validation = validateAnmeldungDraft({
        profileId: state.options.profileId,
        variante: draft.variante,
        nachname: draft.nachname,
        vorname: draft.vorname,
        wahlfach: draft.wahlfachMuendlich,
        seminar: draft.seminar,
        lfs: draft.lfs,
        lehrerLfs: draft.lehrerLFS,
        bestaetigung: draft.bestaetigung
    });
    if (!validation.ok) {
        state.formErrors = validation.errors;
        paint();
        return;
    }
    if (!draft.klasse) {
        state.formErrors = ['Klasse fehlt.'];
        paint();
        return;
    }
    if (!isSrdpAnmeldeFensterOffen(state.options)) {
        state.formErrors = ['Das Anmeldefenster ist geschlossen.'];
        paint();
        return;
    }
    const already = findAnmeldungBySchuelerEmail(state.items, ownership.schuelerEmail);
    if (already) {
        state.submitSuccess = already;
        state.view = 'meine';
        state.formErrors = [];
        toast('Du bist bereits angemeldet.');
        paint();
        return;
    }
    state.formErrors = [];
    state.saving = true;
    paint();
    try {
        let saved;
        if (state.localDemoOnly) {
            saved = Object.assign({}, draft, {
                itemId: 'demo-new-' + Date.now(),
                title: validation.title,
                pruefplanKurz: validation.pruefplanKurz,
                schuelerEmail: ownership.schuelerEmail,
                beantragtVon: ownership.beantragtVon,
                createdBy: ownership.schuelerEmail,
                bestaetigung: true,
                nacharbeitErledigt: false,
                seedTag: DEMO_SEED_TAG,
                created: toIsoDateOnly(new Date()),
                wahlfachMuendlich: draft.wahlfachMuendlich
            });
            state.items = [saved].concat(state.items || []);
        } else {
            if (!state.ctx) {
                await loadFromSharePoint();
            }
            if (!state.ctx) throw new Error(state.error || 'Liste nicht geladen.');
            saved = await createAnmeldungItem(state.ctx, draft, {
                profileId: state.options.profileId,
                jahr: state.options.jahr,
                accountEmail: ownership.schuelerEmail,
                accountName: state.demoPersonaName || state.accountName,
                students: (state.stammdaten && state.stammdaten.students) || []
            });
            state.items = [saved].concat(state.items || []);
        }
        state.submitSuccess = saved;
        state.view = 'meine';
        toast('Anmeldung gespeichert.');
    } catch (e) {
        if (isAccessDeniedError(e) && state.localDemoOnly) {
            state.formErrors = [];
            state.softError = 'Speichern in SharePoint nicht möglich – lokal gespeichert.';
        } else {
            state.formErrors = [(e && e.message) || String(e)];
        }
    } finally {
        state.saving = false;
        paint();
    }
}

function syncFilterMultiFromDom(key) {
    const boxes = root.querySelectorAll(
        'input[data-srdp-filter-multi-item="' + key + '"]'
    );
    const vals = [];
    boxes.forEach((el) => {
        if (el.checked) vals.push(el.value);
    });
    const joined = joinCodeList(vals);
    if (key === 'klasse') state.filterKlasse = joined;
    else if (key === 'variante') state.filterVariante = joined;
    else if (key === 'lehrer') state.filterLehrer = joined;
}

function onClick(ev) {
    const t = ev.target;
    if (!(t instanceof Element)) return;

    const roleBtn = t.closest('[data-srdp-role]');
    if (roleBtn) {
        const role = roleBtn.getAttribute('data-srdp-role') || 'schueler';
        state.role = role === 'admin' ? 'admin' : role === 'lehrer' ? 'lehrer' : 'schueler';
        persistRole(state.role);
        state.submitSuccess = null;
        state.detailItemId = '';
        if (state.localDemoOnly) {
            if (state.role === 'lehrer') {
                state.demoPersonaEmail = DEMO_TEACHER_PERSONA.email;
                state.demoPersonaName = DEMO_TEACHER_PERSONA.name;
                state.teacherLabel = DEMO_TEACHER_PERSONA.label;
                state.kvKlassen = kvKlasseCodesForEmail(
                    state.stammdaten,
                    DEMO_TEACHER_PERSONA.email
                );
            } else if (state.role === 'schueler') {
                state.demoPersonaEmail = DEMO_STUDENT_PERSONA.email;
                state.demoPersonaName = DEMO_STUDENT_PERSONA.name;
            } else {
                state.demoPersonaEmail = '';
                state.demoPersonaName = '';
            }
        } else if (state.role !== 'schueler') {
            state.demoPersonaEmail = '';
            state.demoPersonaName = '';
        }
        if (state.role === 'schueler') prefillStudentForm();
        ensureViewForRole();
        refreshChrome();
        paint();
        return;
    }

    const groupBtn = t.closest('[data-srdp-group]');
    if (groupBtn) {
        const g = String(groupBtn.getAttribute('data-srdp-group') || '');
        state.groupBy = g === 'variante' || g === 'wahlfach' || g === 'klasse' ? g : '';
        paint();
        return;
    }

    const sysTab = t.closest('[data-srdp-sys-tab]');
    if (sysTab) {
        const raw = String(sysTab.getAttribute('data-srdp-sys-tab') || 'site').toLowerCase();
        state.systemTab =
            raw === 'perms' || raw === 'kalender' || raw === 'wartung' || raw === 'danger'
                ? raw
                : 'site';
        paint();
        return;
    }

    const viewBtn = t.closest('[data-srdp-view]');
    if (viewBtn) {
        let view = viewBtn.getAttribute('data-srdp-view') || state.view;
        if (view === 'einrichtung') view = 'system';
        state.view = view;
        if (state.view === 'system' && !state.systemTab) state.systemTab = 'site';
        if (state.view === 'kalender') focusCalendarOnExams();
        state.submitSuccess = null;
        paint();
        return;
    }

    const jumpListe = t.closest('[data-srdp-jump-liste]');
    if (jumpListe) {
        const klasse = jumpListe.getAttribute('data-srdp-jump-klasse') || '';
        applyFilterJump(state, { klasse });
        state.view = 'liste';
        state.submitSuccess = null;
        paint();
        return;
    }

    const jump = t.closest('[data-srdp-view-jump]');
    if (jump) {
        let view = jump.getAttribute('data-srdp-view-jump') || state.view;
        if (view === 'einrichtung') view = 'system';
        // Filter-Jumps leeren andere Filter, außer explizite Attribute
        if (jump.hasAttribute('data-srdp-jump-klasse') || jump.hasAttribute('data-srdp-jump-variante')) {
            applyFilterJump(state, {
                klasse: jump.getAttribute('data-srdp-jump-klasse') || '',
                variante: jump.getAttribute('data-srdp-jump-variante') || ''
            });
        }
        state.view = view;
        if (state.view === 'kalender') focusCalendarOnExams();
        state.submitSuccess = null;
        paint();
        return;
    }

    if (t.closest('[data-srdp-drawer-close]')) {
        state.detailItemId = '';
        paint();
        return;
    }

    const openDetail = t.closest('[data-srdp-open-detail]');
    if (openDetail) {
        state.detailItemId = openDetail.getAttribute('data-srdp-open-detail') || '';
        paint();
        return;
    }

    const toggleNach = t.closest('[data-srdp-toggle-nacharbeit]');
    if (toggleNach) {
        const id = toggleNach.getAttribute('data-srdp-toggle-nacharbeit');
        void toggleNacharbeit(id);
        return;
    }

    if (t.closest('[data-srdp-dismiss-success]')) {
        state.submitSuccess = null;
        paint();
        return;
    }

    if (t.closest('[data-srdp-wizard-next]')) {
        void (async () => {
            state.form = readFormFromDom(state.form);
            if (!effectiveAccountEmail(state)) await syncAccountAndRole();
            const step = Number(state.form.wizardStep) || 1;
            if (!effectiveAccountEmail(state)) {
                state.formErrors = [
                    'Bitte oben rechts anmelden – ohne Schul-Konto kann die Anmeldung nicht gespeichert werden.'
                ];
                paint();
                return;
            }
            if (step === 1) {
                const owned = resolveSrdpOwnership(state.form, {
                    accountEmail: effectiveAccountEmail(state),
                    accountName: state.demoPersonaName || state.accountName,
                    students: (state.stammdaten && state.stammdaten.students) || []
                });
                state.form = emptyForm({
                    ...state.form,
                    nachname: state.form.nachname || owned.nachname,
                    vorname: state.form.vorname || owned.vorname,
                    klasse: state.form.klasse || owned.klasse || ''
                });
                if (!String(state.form.nachname || '').trim() || !String(state.form.vorname || '').trim()) {
                    state.formErrors = [
                        'Name aus dem Konto fehlt – bitte anmelden oder Name ergänzen.'
                    ];
                    paint();
                    return;
                }
                if (!String(state.form.klasse || '').trim()) {
                    state.formErrors = ['Bitte die Klasse wählen.'];
                    paint();
                    return;
                }
            }
            if (step === 2 && !state.form.variante) {
                state.formErrors = ['Bitte eine Variante wählen.'];
                paint();
                return;
            }
            state.formErrors = [];
            state.form.wizardStep = Math.min(4, step + 1);
            paint();
        })();
        return;
    }

    if (t.closest('[data-srdp-wizard-prev]')) {
        state.form = readFormFromDom(state.form);
        state.form.wizardStep = Math.max(1, (Number(state.form.wizardStep) || 1) - 1);
        state.formErrors = [];
        paint();
        return;
    }

    if (t.closest('[data-srdp-submit]')) {
        submitAnmeldung();
        return;
    }

    if (t.closest('[data-srdp-demo]')) {
        enableDemo();
        return;
    }

    if (t.closest('[data-srdp-demo-reset]')) {
        resetDemoSeed();
        return;
    }

    if (t.closest('[data-srdp-reload]')) {
        const saved = String(loadSavedSiteUrl() || '').trim();
        const current = String(state.siteUrl || '').trim();
        const preferred = saved || current;
        const isDemoSite = /demo\.sharepoint\.com/i.test(preferred);
        if (preferred && !isDemoSite) {
            state.siteUrl = preferred;
            state.localDemoOnly = false;
            void loadFromSharePoint();
        } else if (state.localDemoOnly) {
            enableDemo();
        } else {
            void loadFromSharePoint();
        }
        return;
    }

    if (t.closest('[data-srdp-save-site]')) {
        const input = document.getElementById('srdpSiteUrl');
        const url = input ? String(input.value || '').trim() : '';
        state.siteUrl = url;
        persistSiteUrl(url);
        state.localDemoOnly = false;
        loadFromSharePoint();
        return;
    }

    if (t.closest('[data-srdp-save-termine]')) {
        const fristEl = document.getElementById('srdpAnmeldeFrist');
        const hauptEl = document.getElementById('srdpHaupttermin');
        const anmeldeFrist = fristEl ? String(fristEl.value || '').trim() : '';
        const haupttermin = hauptEl ? String(hauptEl.value || '').trim() : '';
        state.wizard = persistWizardConfig({ anmeldeFrist, haupttermin });
        state.options = resolveFormOptions(state.wizard, state.stammdaten, state.options.profileId);
        toast('Anmeldefrist und Haupttermin gespeichert.');
        paint();
        return;
    }

    if (t.closest('[data-srdp-filter-reset]')) {
        applyFilterJump(state, {});
        paint();
        return;
    }

    const clearMulti = t.closest('[data-srdp-filter-multi-clear]');
    if (clearMulti) {
        const key = clearMulti.getAttribute('data-srdp-filter-multi-clear');
        if (key === 'klasse') state.filterKlasse = '';
        if (key === 'variante') state.filterVariante = '';
        if (key === 'lehrer') state.filterLehrer = '';
        paint();
        return;
    }

    const calMode = t.closest('[data-srdp-cal-mode]');
    if (calMode) {
        const mode = String(calMode.getAttribute('data-srdp-cal-mode') || 'month');
        state.calMode = mode === 'agenda' || mode === 'week' || mode === 'month' ? mode : 'month';
        paint();
        return;
    }

    const calNav = t.closest('[data-srdp-cal-nav]');
    if (calNav) {
        state.calAnchored = true;
        const delta = Number(calNav.getAttribute('data-srdp-cal-nav') || '0');
        const mode = String(state.calMode || 'month');
        if (mode === 'week') {
            let anchor = new Date(state.calWeekAnchor || Date.now());
            if (Number.isNaN(anchor.getTime())) anchor = new Date();
            if (delta === 0) {
                anchor = startOfWeekMonday(new Date());
            } else {
                anchor.setDate(anchor.getDate() + delta * 7);
            }
            state.calWeekAnchor = toIsoDateOnly(anchor);
        } else {
            if (delta === 0) {
                const d = new Date();
                state.calYear = d.getFullYear();
                state.calMonth = d.getMonth() + 1;
            } else {
                let m = Number(state.calMonth) || 1;
                let y = Number(state.calYear) || new Date().getFullYear();
                m += delta;
                while (m < 1) {
                    m += 12;
                    y -= 1;
                }
                while (m > 12) {
                    m -= 12;
                    y += 1;
                }
                state.calMonth = m;
                state.calYear = y;
            }
        }
        paint();
    }
}

function onChange(ev) {
    const t = ev.target;
    if (!(t instanceof Element)) return;

    if (t.matches('input[data-srdp-filter-multi-item]')) {
        const key = t.getAttribute('data-srdp-filter-multi-item');
        syncFilterMultiFromDom(key);
        paint();
        return;
    }

    const filter = t.getAttribute('data-srdp-filter');
    if (filter === 'query') {
        state.filterQuery = t.value;
        paint();
        return;
    }

    if (
        t.id === 'srdpWahlfach' ||
        t.id === 'srdpVariante' ||
        t.id === 'srdpLfs' ||
        t.id === 'srdpSeminar' ||
        t.id === 'srdpSeminarText'
    ) {
        state.form = readFormFromDom(state.form);
        const wrap = document.getElementById('srdpSeminarWrap');
        if (wrap && t.id === 'srdpWahlfach') {
            wrap.hidden = !isSeminarWahlfach(state.form.wahlfachMuendlich);
        }
        if (t.id === 'srdpSeminar') {
            const text = document.getElementById('srdpSeminarText');
            if (text) text.value = String(t.value || '');
            state.form.seminar = String(t.value || '').trim();
        }
        const live = document.getElementById('srdpLivePlan');
        if (live && (t.id === 'srdpVariante' || t.id === 'srdpLfs')) {
            const plan = buildPruefplanDetail(state.form, state.options.profileId, {
                terminJahr: state.options.jahr,
                bfkDatum: state.options.bfkDatum || ''
            });
            const bits = plan.schriftlich
                .map((row) => row.part + (row.datum ? ' · ' + row.datum : row.datumHinweis ? ' · ' + row.datumHinweis : ''))
                .filter(Boolean);
            if (bits.length) {
                live.hidden = false;
                live.textContent = 'Schriftlich: ' + bits.join(', ');
            } else {
                live.hidden = true;
            }
        }
        if (t.id === 'srdpVariante' || t.id === 'srdpLfs') paint();
    }
}

function onToggle(ev) {
    const t = ev.target;
    if (!(t instanceof Element)) return;
    const multi = t.closest('[data-srdp-filter-multi]');
    if (multi && t.tagName === 'DETAILS') {
        const key = multi.getAttribute('data-srdp-filter-multi');
        if (!state.filterMultiOpen) state.filterMultiOpen = {};
        state.filterMultiOpen[key] = !!multi.open;
    }
}

async function resolveSiteFromEvents() {
    const url = loadSavedSiteUrl();
    if (url && url !== state.siteUrl) {
        state.siteUrl = url;
        if (!state.localDemoOnly) await loadFromSharePoint();
        else paint();
    }
}

async function onAuthReadyOrChanged() {
    await syncAccountAndRole();
    refreshChrome();
    await resolveSiteFromEvents();
    paint();
    void refreshStudentClassFromGroups();
}

async function boot() {
    root = document.getElementById('srdpApp');
    if (!root) return;

    // Auth-Listener früh – sonst verpassen wir den ersten ms365-auth-state-changed
    window.addEventListener('ms365-auth-state-changed', () => {
        void onAuthReadyOrChanged();
    });
    window.addEventListener('ms365-auth-widget-ready', () => {
        void onAuthReadyOrChanged();
    });
    window.addEventListener('ms365-license-changed', () => {
        void syncAccountAndRole().then(() => {
            paint();
            return resolveSiteFromEvents();
        });
    });
    window.addEventListener('ms365-it-shell-changed', () => {
        void syncAccountAndRole().then(() => {
            refreshChrome();
            paint();
        });
    });

    root.addEventListener('click', onClick);
    root.addEventListener('keydown', (ev) => {
        const t = ev.target;
        if (!(t instanceof Element)) return;
        if (ev.key !== 'Enter' && ev.key !== ' ') return;
        const row = t.closest('tr[data-srdp-open-detail]');
        if (!row || t.closest('button, a, input, select, textarea')) return;
        ev.preventDefault();
        state.detailItemId = row.getAttribute('data-srdp-open-detail') || '';
        paint();
    });
    document.body.addEventListener('click', (ev) => {
        const t = ev.target;
        if (!(t instanceof Element) || !t.closest('#srdpDrawerHost')) return;
        onClick(ev);
    });
    root.addEventListener('change', onChange);
    root.addEventListener('toggle', onToggle, true);
    root.addEventListener('input', (ev) => {
        const el = ev.target;
        if (!(el instanceof Element)) return;
        if (el.getAttribute('data-srdp-filter') === 'query') {
            state.filterQuery = el.value;
            const start = el instanceof HTMLInputElement ? el.selectionStart : null;
            paint();
            const next = document.querySelector('[data-srdp-filter="query"]');
            if (next instanceof HTMLInputElement) {
                next.focus();
                if (typeof start === 'number') next.setSelectionRange(start, start);
            }
            return;
        }
        const qKey = el.getAttribute('data-srdp-filter-multi-q');
        if (qKey) {
            if (!state.filterMultiQuery) state.filterMultiQuery = {};
            state.filterMultiQuery[qKey] = el.value;
            if (!state.filterMultiOpen) state.filterMultiOpen = {};
            state.filterMultiOpen[qKey] = true;
            paint();
        }
    });

    state.demoRoleOverride = readDemoRoleOverrideFromUrl();
    state.siteUrl = loadSavedSiteUrl() || state.siteUrl;
    await syncAccountAndRole();

    const params = new URLSearchParams(window.location.search || '');
    if (params.get('demo') === '1' || state.demoRoleOverride) {
        enableDemo({ withPersona: state.demoRoleOverride && params.get('demo') !== '1' });
    } else if (state.siteUrl) {
        purgeStickyDemoWizard();
        ensureKlassenOptions();
        ensureViewForRole();
        paint();
        await loadFromSharePoint();
        ensureKlassenOptions();
        if (state.role === 'schueler') prefillStudentForm();
        void refreshStudentClassFromGroups();
        if (state.error && !(state.items || []).length) {
            const hidden = /nicht sichtbar/i.test(state.error);
            state.softError = hidden
                ? state.error
                : state.error + ' – Tipp: „Demo lokal“ unter Systemadministration oder Liste anlegen.';
            state.error = '';
        }
        paint();
    } else {
        enableDemo();
    }

    // Nochmal nachziehen – MSAL oft erst jetzt fertig
    await syncAccountAndRole();
    refreshChrome();
    paint();
    void refreshPruefungstermineFromWeb({ terminJahr: state.options && state.options.jahr }).then((result) => {
        state.termineSource = (result && result.source) || 'snapshot';
        state.termineFetchedAt = (result && result.fetchedAt) || '';
        paint();
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}
