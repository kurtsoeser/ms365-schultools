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
    kvKlasseCodesForEmail,
    persistRole,
    persistSiteUrl,
    loadSavedSiteUrl,
    resolveFormOptions,
    emptyForm,
    viewsForRole,
    applyFilterJump,
    joinCodeList,
    effectiveAccountEmail
} from './srdp-anmeldung-state.js';
import {
    resolveSrdpContext,
    loadAllAnmeldungen,
    createAnmeldungItem,
    updateAnmeldungNacharbeit
} from './srdp-anmeldung-graph.js';
import {
    validateAnmeldungDraft,
    buildPruefplanKurz,
    isSeminarWahlfach,
    resolveSrdpOwnership,
    assertSrdpOwnership,
    isSrdpAnmeldeFensterOffen,
    findAnmeldungBySchuelerEmail
} from './srdp-anmeldung-logic.js';
import {
    readDemoRoleOverrideFromUrl,
    inferSrdpRolesFromStammdaten,
    shellPreferredSrdpRole
} from './srdp-anmeldung-entra-role.js';
import { renderApp, readFormFromDom, syncSrdpPageChrome } from './srdp-anmeldung-ui.js';
import {
    buildLocalDemoState,
    DEMO_TEACHER_PERSONA,
    DEMO_STUDENT_PERSONA,
    DEMO_SEED_TAG,
    stripDemoSeedItems,
    isDemoSeedItem
} from './srdp-anmeldung-demo.js';
import { hydrateStammdatenFromLiveSite } from '../../shared/school-stammdaten-spo-lists.js';
import { toIsoDateOnly, startOfWeekMonday } from '../../shared/ms365-calendar-shell.js';

const state = createInitialState();
let root = null;

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function paint() {
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

async function readAccount() {
    try {
        if (typeof window.ms365MsalGetAccount === 'function') {
            const acc = await window.ms365MsalGetAccount();
            if (acc) {
                state.accountEmail = String(acc.username || acc.mail || '').trim();
                state.accountName = String(acc.name || '').trim();
            }
        }
    } catch {
        /* ignore */
    }
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
    const st = matchStudentByEmail(state.stammdaten.students, effectiveAccountEmail(state));
    const seed = { wizardStep: state.form && state.form.wizardStep };
    if (st) {
        seed.nachname = st.nachname || st.lastName || '';
        seed.vorname = st.vorname || st.firstName || '';
        seed.klasse = st.klasse || st.klasseCode || st.classCode || '';
    } else if (state.accountName || state.demoPersonaName) {
        const parts = String(state.demoPersonaName || state.accountName)
            .split(/\s+/);
        if (parts.length >= 2) {
            seed.vorname = parts[0];
            seed.nachname = parts.slice(1).join(' ');
        }
    }
    state.form = emptyForm(Object.assign({}, state.form, seed));
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
        const ctx = await resolveSrdpContext(siteUrl, {
            profileId: state.options.profileId,
            jahr: state.options.jahr
        });
        state.siteId = ctx.siteId;
        state.listId = ctx.list.id;
        state.listWebUrl = ctx.list.webUrl || '';
        state.localDemoOnly = false;
        state.ctx = ctx;
        await hydrateLiveStammdaten(ctx.siteId);
        state.items = await loadAllAnmeldungen(ctx);
        state.syncMeta = {
            source: 'sharepoint',
            lastPullAt: new Date().toISOString(),
            listTitle: (ctx.list && ctx.list.title) || '',
            connected: true
        };
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
    const demo = buildLocalDemoState(state);
    Object.assign(state, demo);
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
        state.teacherLabel = DEMO_TEACHER_PERSONA.label;
        state.kvKlassen = kvKlasseCodesForEmail(state.stammdaten, DEMO_TEACHER_PERSONA.email);
    } else {
        state.role = shellPreferredSrdpRole(state.planerRoles, state.role);
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
        options: resolveFormOptions(demo.wizard, demo.stammdaten, 'hak')
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
    state.options = resolveFormOptions(demo.wizard, demo.stammdaten, 'hak');
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
    const draft = readFormFromDom();
    state.form = draft;
    const ownership = resolveSrdpOwnership(draft, {
        accountEmail: effectiveAccountEmail(state),
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
                prefillStudentForm();
            } else {
                state.demoPersonaEmail = '';
                state.demoPersonaName = '';
            }
        }
        ensureViewForRole();
        refreshChrome();
        paint();
        return;
    }

    const viewBtn = t.closest('[data-srdp-view]');
    if (viewBtn) {
        let view = viewBtn.getAttribute('data-srdp-view') || state.view;
        if (view === 'einrichtung') view = 'system';
        state.view = view;
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
        state.form = readFormFromDom();
        const step = Number(state.form.wizardStep) || 1;
        if (step === 1 && (!state.form.klasse || !state.form.nachname || !state.form.vorname)) {
            state.formErrors = ['Bitte Klasse, Nachname und Vorname ausfüllen.'];
            paint();
            return;
        }
        if (step === 2 && !state.form.variante) {
            state.formErrors = ['Bitte eine Variante wählen.'];
            paint();
            return;
        }
        state.formErrors = [];
        state.form.wizardStep = Math.min(4, step + 1);
        paint();
        return;
    }

    if (t.closest('[data-srdp-wizard-prev]')) {
        state.form = readFormFromDom();
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

    if (t.id === 'srdpWahlfach' || t.id === 'srdpVariante') {
        state.form = readFormFromDom();
        const wrap = document.getElementById('srdpSeminarWrap');
        if (wrap && t.id === 'srdpWahlfach') {
            wrap.hidden = !isSeminarWahlfach(state.form.wahlfachMuendlich);
        }
        const live = document.getElementById('srdpLivePlan');
        if (live && t.id === 'srdpVariante') {
            const plan = buildPruefplanKurz(state.form.variante, state.options.profileId);
            if (plan) {
                live.hidden = false;
                live.innerHTML = '<strong>Prüfplan:</strong> ' + plan;
            } else {
                live.hidden = true;
            }
        }
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

async function boot() {
    root = document.getElementById('srdpApp');
    if (!root) return;

    state.demoRoleOverride = readDemoRoleOverrideFromUrl();
    state.siteUrl = loadSavedSiteUrl() || state.siteUrl;
    await readAccount();
    refreshDerived();

    const inferred = inferSrdpRolesFromStammdaten(state);
    state.planerRoles = state.demoRoleOverride
        ? ['schueler', 'lehrer', 'admin']
        : inferred;
    state.role = shellPreferredSrdpRole(state.planerRoles, state.role);
    persistRole(state.role);

    const params = new URLSearchParams(window.location.search || '');
    if (params.get('demo') === '1' || state.demoRoleOverride) {
        enableDemo({ withPersona: state.demoRoleOverride && params.get('demo') !== '1' });
    } else if (state.siteUrl) {
        ensureViewForRole();
        paint();
        await loadFromSharePoint();
        if (state.error && /fehlt|Liste/i.test(state.error) && !(state.items || []).length) {
            state.softError =
                state.error + ' – Tipp: „Demo lokal“ unter Systemadministration oder Liste anlegen.';
            state.error = '';
            paint();
        }
    } else {
        enableDemo();
    }

    if (state.role === 'schueler') prefillStudentForm();
    ensureViewForRole();
    refreshChrome();
    paint();

    root.addEventListener('click', onClick);
    root.addEventListener('change', onChange);
    root.addEventListener('toggle', onToggle, true);
    root.addEventListener('input', (ev) => {
        const el = ev.target;
        if (!(el instanceof Element)) return;
        if (el.getAttribute('data-srdp-filter') === 'query') {
            state.filterQuery = el.value;
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

    window.addEventListener('ms365-auth-state-changed', () => {
        void (async () => {
            await readAccount();
            refreshDerived();
            ensureViewForRole();
            refreshChrome();
            await resolveSiteFromEvents();
            paint();
        })();
    });
    window.addEventListener('ms365-license-changed', () => {
        void resolveSiteFromEvents();
    });
    window.addEventListener('ms365-it-shell-changed', () => refreshChrome());
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}
