/**
 * Entry / Wiring: Schularbeiten-Planer
 */
import {
    createInitialState,
    loadStammdaten,
    matchTeacherByEmail,
    matchStudentByEmail,
    resolveRole,
    persistRole,
    persistDemoKlasseCode,
    persistSchuljahr,
    loadSchuljahr,
    schoolYearOptions,
    loadSavedSiteUrl,
    persistSiteUrl,
    emptyForm,
    filterSchularbeiten,
    labelMaps,
    persistPlanerSettings,
    viewsForRole,
    VIEWS,
    scopeFromState,
    canEditSchularbeit,
    canDeleteSchularbeit,
    persistCalShowSettings,
    persistCalModeSettings
} from './schularbeiten-planer-state.js';
import {
    resolvePlanerContext,
    loadAllPlanerData,
    createSchularbeitItem,
    updateSchularbeitItem,
    deleteSchularbeitItem,
    createFensterItem,
    deleteFensterItem,
    updateRegelwerkItem,
    createFachMetaItem,
    updateFachMetaItem,
    deleteFachMetaItem,
    emptyPlanerLists
} from './schularbeiten-planer-graph.js';
import { upsertSchulterminFromSchularbeit } from './schularbeiten-planer-sync.js';
import {
    upsertGroupCalendarEvent,
    removeGroupCalendarEventForSa,
    syncSchularbeitenToGroupCalendars,
    syncSchularbeitenToUserCalendar
} from './schularbeiten-planer-calendar-sync.js';
import {
    seedDemoSchularbeiten,
    parseDemoImportJson,
    applyDemoStammdatenLocal,
    packToLocalPlanerState
} from './schularbeiten-planer-demo-seed.js';
import {
    newEntityId,
    LIST_TITLES,
    LIST_KEYS,
    DEFAULT_FACH_META_STANDARD_DAUER,
    DEFAULT_FACH_META_PRO_SEMESTER,
    FACH_META_COLOR_PALETTE
} from './schularbeiten-planer-schema.js';
import { addDays, mondayOfWeekContaining, toIsoDateOnly } from './schularbeiten-planer-logic.js';
import { DEMO_SITE_DEFAULT } from './schularbeiten-planer-demo-data.js';
import {
    applySchularbeitenPackagePermissions,
    savePermissionsConfig,
    normalizePermissionsConfig
} from './schularbeiten-planer-permissions.js';
import {
    PLANER_GROUP_FIELDS,
    readPermissionsFromPickers,
    wirePermissionGroupPickersDelegated,
    persistPickersToStorage,
    refreshSchularbeitenPlanerExtraUsers
} from './schularbeiten-permissions-ui.js';
import {
    renderApp,
    readFormFromDom,
    readFiltersFromDom,
    readFensterForm,
    readRulesForm
} from './schularbeiten-planer-ui.js';
import { buildIcs, downloadIcs } from './schularbeiten-planer-export.js';
import { validateSchularbeit } from './schularbeiten-planer-logic.js';
import {
    applyPlanerRoleFromEntra,
    isPlanerDemoRoleUiEnabled,
    canShowPlanerItToolbar,
    entraGroupsConfigured
} from './schularbeiten-planer-entra-role.js';
import { loadEffectivePermissionsConfig } from './schularbeiten-planer-permissions.js';
import {
    syncPlannerPermissionsFromSite,
    publishPlannerPermissionsToSite
} from './schularbeiten-planer-remote-config.js';

const state = createInitialState();
let root = null;
/** @type {Map<string, ReturnType<typeof setTimeout>>} */
const fachMetaSaveTimers = new Map();

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else window.alert(msg);
}

function graphMapOpts() {
    return { labels: labelMaps(state.stammdaten) };
}

function placePlanerAuthWidget() {
    const slot = document.getElementById('saNavAuthSlot');
    const wrap = document.getElementById('ms365AuthWidget');
    if (!slot || !wrap) return;
    wrap.style.position = '';
    wrap.style.top = '';
    wrap.style.right = '';
    wrap.style.zIndex = '';
    wrap.style.marginLeft = '';
    wrap.style.flexWrap = 'wrap';
    wrap.style.width = '100%';
    if (wrap.parentElement !== slot) slot.appendChild(wrap);
}

function syncPlanerPageChrome() {
    const showIt = canShowPlanerItToolbar(state.planerRoles, state.role);
    document.querySelectorAll('[data-sa-it-only]').forEach((el) => {
        el.hidden = !showIt;
    });
}

function paint() {
    if (!root) return;
    renderApp(state, root);
    bindStatic();
    placePlanerAuthWidget();
    syncPlanerPageChrome();
}

function bindStatic() {
    root.querySelectorAll('[data-sa-view]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-sa-view') || 'dashboard';
            state.detailId = null;
            paint();
        });
    });

    root.querySelectorAll('[data-sa-role]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const raw = btn.getAttribute('data-sa-role') || 'lehrer';
            const role = raw === 'admin' ? 'admin' : raw === 'schueler' ? 'schueler' : 'lehrer';
            const demoUi = isPlanerDemoRoleUiEnabled(state.entraGroupsConfigured, state.demoRoleOverride);
            const entitled = demoUi
                ? ['admin', 'lehrer', 'schueler']
                : (state.planerRoles || []);
            if (!entitled.includes(role)) return;
            state.role = role;
            state.roleSource =
                (state.planerRoleSources && state.planerRoleSources[role]) || (demoUi ? 'demo' : 'entra');
            persistRole(role);
            state.roleHint = '';
            state.planerAccessDenied = false;
            const allowed = viewsForRole(role).map((v) => v.id);
            if (!allowed.includes(state.view)) state.view = 'dashboard';
            paint();
        });
    });

    const demoKlasse = document.getElementById('saDemoKlasse');
    if (demoKlasse) {
        demoKlasse.addEventListener('change', () => {
            state.demoKlasseCode = String(demoKlasse.value || '').trim();
            persistDemoKlasseCode(state.demoKlasseCode);
            paint();
        });
    }

    const loadBtn = document.getElementById('saBtnLoad');
    if (loadBtn) {
        loadBtn.addEventListener('click', () => {
            const input = document.getElementById('saSiteUrl');
            state.siteUrl = input ? String(input.value || '').trim() : state.siteUrl;
            persistSiteUrl(state.siteUrl);
            refreshData().catch((e) => {
                state.error = e && e.message ? e.message : String(e);
                state.loading = false;
                paint();
            });
        });
    }

    const schuljahrEl = document.getElementById('saSchuljahr');
    if (schuljahrEl) {
        schuljahrEl.addEventListener('change', () => {
            state.schuljahr = persistSchuljahr(schuljahrEl.value);
            refreshData({ silent: true }).catch((e) => {
                state.error = e && e.message ? e.message : String(e);
                paint();
            });
        });
    }

    ['saFilterKlasse', 'saFilterLehrer'].forEach((id) => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('change', () => {
                state.filters = readFiltersFromDom(state.filters);
                paint();
            });
        }
    });
    root.querySelectorAll('[data-sa-filter-pill]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const key = btn.getAttribute('data-sa-filter-key');
            const val = btn.getAttribute('data-sa-filter-value') || '';
            if (key === 'fach' || key === 'status') {
                const cur = String((state.filters && state.filters[key]) || '');
                state.filters = { ...state.filters, [key]: cur === val ? '' : val };
                paint();
            }
        });
    });
    const reset = document.getElementById('saFilterReset');
    if (reset) {
        reset.addEventListener('click', () => {
            state.filters = { klasse: '', fach: '', lehrer: '', status: '' };
            paint();
        });
    }

    bindFormLive();
    bindMeineActions();
    bindAdminActions();
    bindKalender();
    bindExport();
    bindDetail();
    bindPhase5();
    bindJsonImport();
    bindWeekNav();
}

function bindWeekNav() {
    const monday =
        state.dashboardWeekMonday || mondayOfWeekContaining(toIsoDateOnly(new Date()) || '');
    const prev = root.querySelector('[data-sa-week-prev]');
    if (prev) {
        prev.addEventListener('click', () => {
            state.dashboardWeekMonday = addDays(monday, -7) || monday;
            paint();
        });
    }
    const next = root.querySelector('[data-sa-week-next]');
    if (next) {
        next.addEventListener('click', () => {
            state.dashboardWeekMonday = addDays(monday, 7) || monday;
            paint();
        });
    }
    const todayBtn = root.querySelector('[data-sa-week-today]');
    if (todayBtn) {
        todayBtn.addEventListener('click', () => {
            state.dashboardWeekMonday = '';
            paint();
        });
    }
}

function bindFormLive() {
    const ids = [
        'saTitel',
        'saThema',
        'saFach',
        'saKlasse',
        'saDatum',
        'saBeginn',
        'saDauer',
        'saSemester',
        'saLehrer',
        'saNotiz'
    ];
    const sync = () => {
        if (state.view !== 'neu') return;
        const form = readFormFromDom();
        const teacher = state.stammdaten.teachers.find((t) => t.code === form.lehrerCode);
        form.lehrerEmail = teacher ? teacher.email : '';
        form.schularbeitId = state.form.schularbeitId || '';
        state.form = form;
        paint();
        // restore focus roughly
        const activeId = document.activeElement && document.activeElement.id;
        if (activeId && ids.includes(activeId)) {
            const el = document.getElementById(activeId);
            if (el) el.focus();
        }
    };
    ids.forEach((id) => {
        const el = document.getElementById(id);
        if (!el) return;
        el.addEventListener('change', sync);
        if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
            el.addEventListener('input', () => {
                // debounce light: only store, re-validate on change for selects; for text update state without full paint spam
                const form = readFormFromDom();
                const teacher = state.stammdaten.teachers.find((t) => t.code === form.lehrerCode);
                form.lehrerEmail = teacher ? teacher.email : '';
                form.schularbeitId = state.form.schularbeitId || '';
                state.form = form;
            });
        }
    });

    // Re-validate paint on date/select change already via sync.
    // For thema/notiz: button enable via submit click validation.

    const submit = document.getElementById('saBtnSubmit');
    if (submit) {
        submit.addEventListener('click', () => {
            submitForm().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }
    const cancel = document.getElementById('saBtnCancelForm');
    if (cancel) {
        cancel.addEventListener('click', () => {
            state.form = emptyForm(defaultTeacherForm());
            state.editingItemId = null;
            state.view = 'meine';
            paint();
        });
    }
}

function defaultTeacherForm() {
    const t = state.teacherMatch;
    if (!t) return {};
    return { lehrerCode: t.code, lehrerEmail: t.email };
}

async function submitForm() {
    if (!state.ctx) throw new Error('Bitte zuerst Daten laden.');
    const form = readFormFromDom();
    const teacher = state.stammdaten.teachers.find((t) => t.code === form.lehrerCode);
    form.lehrerEmail = (teacher && teacher.email) || state.accountEmail || '';
    const prev =
        state.editingItemId ? state.items.find((x) => x.itemId === state.editingItemId) : null;
    const prevStatus = prev ? String(prev.status || '').toLowerCase() : '';
    const keepStatus =
        prev &&
        state.role === 'admin' &&
        (prevStatus === 'fixiert' || prevStatus === 'abgelehnt') &&
        prevStatus;

    const draft = {
        schularbeitId: (prev && prev.schularbeitId) || state.form.schularbeitId || newEntityId('sa'),
        titel: form.titel,
        thema: form.thema,
        fachCode: form.fachCode,
        klasseCode: form.klasseCode,
        lehrerCode: form.lehrerCode,
        lehrerEmail: form.lehrerEmail,
        datum: form.datum,
        beginnUhrzeit: form.beginnUhrzeit,
        dauerMinuten: form.dauerMinuten,
        semester: form.semester,
        notiz: form.notiz,
        status: keepStatus || 'beantragt',
        beantragtVon: (prev && prev.beantragtVon) || state.accountEmail || '',
        schuljahr: state.schuljahr || loadSchuljahr(),
        fixiertVon: keepStatus === 'fixiert' && prev ? prev.fixiertVon : '',
        fixiertAm: keepStatus === 'fixiert' && prev ? prev.fixiertAm : '',
        ablehnungsGrund: keepStatus === 'abgelehnt' && prev ? prev.ablehnungsGrund : '',
        teamsCalendarEventId: prev ? prev.teamsCalendarEventId : ''
    };
    const check = validateSchularbeit({
        draft,
        existing: filterSchularbeiten(state.items, {}, scopeFromState(state, { scopeAll: true })),
        rules: state.rules,
        windows: state.windows
    });
    if (!check.canSubmit) {
        toast(check.errors[0] || 'Regelverstoß');
        state.form = { ...form, schularbeitId: draft.schularbeitId };
        paint();
        return;
    }

    state.loading = true;
    paint();
    const mapOpts = graphMapOpts();
    try {
        if (state.editingItemId) {
            const payload = prev ? { ...prev, ...draft, itemId: prev.itemId } : draft;
            await updateSchularbeitItem(state.ctx, state.editingItemId, payload, mapOpts);
            if (String(payload.status || '').toLowerCase() === 'fixiert') {
                await maybeSyncSchultermin(payload);
                await maybeSyncClassCalendar(payload);
            }
            toast(keepStatus === 'fixiert' ? 'Termin aktualisiert.' : 'Antrag aktualisiert.');
        } else {
            await createSchularbeitItem(state.ctx, draft, mapOpts);
            toast('Antrag eingereicht.');
        }
        const returnView = state.editReturnView || (state.role === 'admin' ? 'liste' : 'meine');
        state.editingItemId = null;
        state.editReturnView = '';
        state.form = emptyForm(defaultTeacherForm());
        state.view = returnView;
        await refreshData({ silent: true });
    } finally {
        state.loading = false;
        paint();
    }
}

function bindMeineActions() {
    root.querySelectorAll('[data-sa-edit]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-sa-edit');
            const sa = state.items.find((x) => x.itemId === id);
            const scope = scopeFromState(state);
            if (!sa || !canEditSchularbeit(sa, scope)) return;
            state.detailId = null;
            state.editReturnView = state.view;
            state.editingItemId = id;
            state.form = emptyForm({
                titel: sa.titel,
                thema: sa.thema,
                fachCode: sa.fachCode,
                klasseCode: sa.klasseCode,
                lehrerCode: sa.lehrerCode,
                lehrerEmail: sa.lehrerEmail,
                datum: sa.datum,
                beginnUhrzeit: sa.beginnUhrzeit || '08:00',
                dauerMinuten: sa.dauerMinuten,
                semester: sa.semester,
                notiz: sa.notiz,
                schularbeitId: sa.schularbeitId
            });
            state.view = 'neu';
            paint();
        });
    });
    root.querySelectorAll('[data-sa-del]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-sa-del');
            const sa = state.items.find((x) => x.itemId === id);
            const scope = scopeFromState(state);
            if (!id || !state.ctx || !sa || !canDeleteSchularbeit(sa, scope)) return;
            const label =
                String(sa.status || '').toLowerCase() === 'fixiert' ? 'Fixierten Termin' : 'Antrag';
            if (!window.confirm(label + ' wirklich löschen?')) return;
            state.detailId = null;
            maybeRemoveClassCalendar(sa)
                .then(() => deleteSchularbeitItem(state.ctx, id))
                .then(() => refreshData({ silent: true }))
                .then(() => toast('Gelöscht.'))
                .catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    });
}

function bindAdminActions() {
    refreshSchularbeitenPlanerExtraUsers(() => persistPickersToStorage(PLANER_GROUP_FIELDS));
    root.querySelectorAll('[data-sa-fix]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-sa-fix');
            const sa = state.items.find((x) => x.itemId === id);
            if (!sa || !state.ctx) return;
            const patch = {
                ...sa,
                status: 'fixiert',
                ablehnungsGrund: '',
                fixiertVon: state.accountEmail || '',
                fixiertAm: new Date().toISOString()
            };
            state.detailId = null;
            updateSchularbeitItem(state.ctx, id, patch, graphMapOpts())
                .then(() => maybeSyncSchultermin(patch))
                .then((syncResult) =>
                    maybeSyncClassCalendar(patch).then((calResult) => ({ syncResult, calResult }))
                )
                .then(({ syncResult, calResult }) =>
                    refreshData({ silent: true }).then(() => {
                        const parts = ['Fixiert'];
                        if (syncResult && syncResult.created) parts.push('Schultermine angelegt');
                        else if (syncResult) parts.push('Schultermine aktualisiert');
                        if (calResult && calResult.created) parts.push('Klassenkalender angelegt');
                        else if (calResult) parts.push('Klassenkalender aktualisiert');
                        toast(parts.join(' · ') + '.');
                    })
                )
                .catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    });
    root.querySelectorAll('[data-sa-reject]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-sa-reject');
            const sa = state.items.find((x) => x.itemId === id);
            if (!sa || !state.ctx) return;
            const reason = window.prompt('Ablehnungsgrund (optional):', '') || '';
            const patch = {
                ...sa,
                status: 'abgelehnt',
                ablehnungsGrund: reason,
                fixiertVon: state.accountEmail || '',
                fixiertAm: new Date().toISOString()
            };
            state.detailId = null;
            maybeRemoveClassCalendar(sa)
                .then(() => {
                    patch.teamsCalendarEventId = '';
                    return updateSchularbeitItem(state.ctx, id, patch, graphMapOpts());
                })
                .then(() => refreshData({ silent: true }))
                .then(() => toast('Abgelehnt.'))
                .catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    });

    const addTf = document.getElementById('saBtnAddFenster');
    if (addTf) {
        addTf.addEventListener('click', () => {
            const win = readFensterForm();
            if (!win.titel || !win.startdatum || !win.enddatum) {
                toast('Titel, Von und Bis sind Pflicht.');
                return;
            }
            if (!state.ctx) return;
            createFensterItem(state.ctx, win)
                .then(() => refreshData({ silent: true }))
                .then(() => toast('Sperrzeit angelegt.'))
                .catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }
    root.querySelectorAll('[data-sa-del-fenster]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-sa-del-fenster');
            if (!id || !state.ctx) return;
            if (!window.confirm('Terminfenster löschen?')) return;
            deleteFensterItem(state.ctx, id)
                .then(() => refreshData({ silent: true }))
                .then(() => toast('Gelöscht.'))
                .catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    });

    const saveRw = document.getElementById('saBtnSaveRules');
    if (saveRw) {
        saveRw.addEventListener('click', () => {
            if (!state.ctx || !state.rules.itemId) {
                toast('Kein Regelwerk-Eintrag geladen.');
                return;
            }
            const rw = { ...state.rules, ...readRulesForm(), regelwerkId: state.rules.regelwerkId };
            updateRegelwerkItem(state.ctx, state.rules.itemId, rw)
                .then(() => refreshData({ silent: true }))
                .then(() => toast('Regelwerk gespeichert.'))
                .catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }

    const savePerms = document.getElementById('saBtnSavePerms');
    if (savePerms) {
        savePerms.addEventListener('click', () => {
            persistPickersToStorage(PLANER_GROUP_FIELDS);
            state.entraGroupsConfigured = entraGroupsConfigured(loadEffectivePermissionsConfig());
            resolvePlanerRole()
                .then(() => paint())
                .catch(() => paint());
            toast('Gruppen für Berechtigungen gespeichert.');
        });
    }
    const applyPerms = document.getElementById('saBtnApplyPerms');
    if (applyPerms) {
        applyPerms.addEventListener('click', () => {
            const webUrl = state.siteUrl || (state.ctx && state.ctx.webUrl) || '';
            if (!webUrl) {
                toast('Bitte zuerst SharePoint-Site laden.');
                return;
            }
            if (
                !window.confirm(
                    'Berechtigungen auf allen Planer-Listen anwenden?\n\n' +
                        webUrl +
                        '\n\nVererbung wird gebrochen; breite Site-Besucher/Mitglieder werden von den Listen entfernt.'
                )
            ) {
                return;
            }
            const cfg = normalizePermissionsConfig(readPermissionsFromDom());
            savePermissionsConfig(cfg);
            state.loading = true;
            paint();
            applySchularbeitenPackagePermissions(webUrl, cfg, (msg) => console.log('[sa-perms]', msg))
                .then(() => toast('SharePoint-Berechtigungen angewendet.'))
                .catch((e) => toast(e && e.message ? e.message : String(e)))
                .finally(() => {
                    state.loading = false;
                    paint();
                });
        });
    }
}

function readPermissionsFromDom() {
    return readPermissionsFromPickers(PLANER_GROUP_FIELDS);
}

async function moveSchularbeitOnCalendar(itemId, newDateIso) {
    if (state.role !== 'admin' || !state.ctx) return;
    const sa = state.items.find((x) => x.itemId === itemId);
    if (!sa) return;
    const scope = scopeFromState(state);
    if (!canEditSchularbeit(sa, scope)) {
        toast('Dieser Termin kann nicht verschoben werden.');
        return;
    }
    const nextDate = String(newDateIso || '').trim().slice(0, 10);
    if (!nextDate || nextDate === String(sa.datum || '').slice(0, 10)) return;

    const draft = { ...sa, datum: nextDate };
    const check = validateSchularbeit({
        draft,
        existing: filterSchularbeiten(state.items, {}, scopeFromState(state, { scopeAll: true })),
        rules: state.rules,
        windows: state.windows
    });
    if (!check.canSubmit) {
        const ok = window.confirm(
            (check.errors[0] || 'Regelverstoß') + '\n\nTrotzdem auf ' + nextDate + ' verschieben?'
        );
        if (!ok) return;
    }

    state.loading = true;
    paint();
    try {
        const payload = { ...sa, datum: nextDate };
        await updateSchularbeitItem(state.ctx, sa.itemId, payload, graphMapOpts());
        if (String(payload.status || '').toLowerCase() === 'fixiert') {
            await maybeSyncSchultermin(payload);
            await maybeSyncClassCalendar(payload);
        }
        await refreshData({ silent: true });
        toast('Termin verschoben auf ' + nextDate + '.');
    } catch (e) {
        state.error = e && e.message ? e.message : String(e);
        toast(state.error);
        state.loading = false;
        paint();
    }
}

function bindCalShowToggles() {
    const map = [
        ['saCalShowBeantragt', 'beantragt'],
        ['saCalShowFixiert', 'fixiert'],
        ['saCalShowAbgelehnt', 'abgelehnt']
    ];
    map.forEach(([id, key]) => {
        const el = document.getElementById(id);
        if (!el) return;
        el.addEventListener('change', () => {
            state.calShow = persistCalShowSettings({ [key]: !!el.checked });
            paint();
        });
    });
}

function bindKalenderDragDrop() {
    if (state.role !== 'admin' || !root) return;
    root.querySelectorAll('[data-sa-cal-drag]').forEach((btn) => {
        btn.addEventListener('dragstart', (e) => {
            const id = btn.getAttribute('data-sa-cal-drag');
            if (!id || !e.dataTransfer) return;
            e.dataTransfer.setData('text/plain', id);
            e.dataTransfer.effectAllowed = 'move';
            btn.classList.add('is-dragging');
        });
        btn.addEventListener('dragend', () => {
            btn.classList.remove('is-dragging');
            root.querySelectorAll('.sa-cal__cell.is-drop-target').forEach((c) => c.classList.remove('is-drop-target'));
        });
    });
    root.querySelectorAll('[data-sa-cal-droppable]').forEach((cell) => {
        cell.addEventListener('dragover', (e) => {
            e.preventDefault();
            if (e.dataTransfer) e.dataTransfer.dropEffect = 'move';
            cell.classList.add('is-drop-target');
        });
        cell.addEventListener('dragleave', () => {
            cell.classList.remove('is-drop-target');
        });
        cell.addEventListener('drop', (e) => {
            e.preventDefault();
            cell.classList.remove('is-drop-target');
            const id = e.dataTransfer ? e.dataTransfer.getData('text/plain') : '';
            const iso = cell.getAttribute('data-sa-cal-drop');
            if (!id || !iso) return;
            moveSchularbeitOnCalendar(id, iso).catch((err) =>
                toast(err && err.message ? err.message : String(err))
            );
        });
    });
}

function syncCalMonthFromWeekMonday() {
    const mon = String(state.calWeekMonday || '').trim().slice(0, 10);
    if (!mon) return;
    const parts = mon.split('-');
    if (parts.length < 2) return;
    state.calYear = parseInt(parts[0], 10);
    state.calMonth = parseInt(parts[1], 10) - 1;
}

function bindKalender() {
    bindCalShowToggles();
    bindKalenderDragDrop();
    const modeMonth = document.getElementById('saCalModeMonth');
    const modeWeek = document.getElementById('saCalModeWeek');
    if (modeMonth) {
        modeMonth.addEventListener('click', () => {
            if (state.calMode === 'month') return;
            state.calMode = persistCalModeSettings('month');
            syncCalMonthFromWeekMonday();
            paint();
        });
    }
    if (modeWeek) {
        modeWeek.addEventListener('click', () => {
            if (state.calMode === 'week') return;
            state.calMode = persistCalModeSettings('week');
            const anchor = `${state.calYear}-${String(state.calMonth + 1).padStart(2, '0')}-15`;
            state.calWeekMonday = mondayOfWeekContaining(anchor) || mondayOfWeekContaining(toIsoDateOnly(new Date()) || '') || '';
            paint();
        });
    }
    const prev = document.getElementById('saCalPrev');
    const next = document.getElementById('saCalNext');
    const today = document.getElementById('saCalToday');
    if (prev) {
        prev.addEventListener('click', () => {
            if (state.calMode === 'week') {
                const mon =
                    state.calWeekMonday ||
                    mondayOfWeekContaining(
                        `${state.calYear}-${String(state.calMonth + 1).padStart(2, '0')}-15`
                    ) ||
                    '';
                state.calWeekMonday = addDays(mon, -7) || mon;
                syncCalMonthFromWeekMonday();
            } else {
                state.calMonth -= 1;
                if (state.calMonth < 0) {
                    state.calMonth = 11;
                    state.calYear -= 1;
                }
            }
            paint();
        });
    }
    if (next) {
        next.addEventListener('click', () => {
            if (state.calMode === 'week') {
                const mon =
                    state.calWeekMonday ||
                    mondayOfWeekContaining(
                        `${state.calYear}-${String(state.calMonth + 1).padStart(2, '0')}-15`
                    ) ||
                    '';
                state.calWeekMonday = addDays(mon, 7) || mon;
                syncCalMonthFromWeekMonday();
            } else {
                state.calMonth += 1;
                if (state.calMonth > 11) {
                    state.calMonth = 0;
                    state.calYear += 1;
                }
            }
            paint();
        });
    }
    if (today) {
        today.addEventListener('click', () => {
            const n = new Date();
            state.calYear = n.getFullYear();
            state.calMonth = n.getMonth();
            state.calWeekMonday = mondayOfWeekContaining(toIsoDateOnly(n) || '') || '';
            paint();
        });
    }
}

function bindExport() {
    const icsBtn = document.getElementById('saBtnIcs');
    if (icsBtn) {
        icsBtn.addEventListener('click', () => {
            const items = filterSchularbeiten(
                state.items,
                state.filters,
                scopeFromState(state, {
                    scopeAll: state.role === 'admin',
                    onlyMine: state.role === 'lehrer'
                })
            ).filter((s) =>
                state.role === 'schueler' ? true : s.status === 'fixiert' || s.status === 'beantragt'
            );
            const ics = buildIcs(items, state.stammdaten, 'Schularbeiten');
            downloadIcs(ics);
            toast('ICS heruntergeladen (' + items.length + ').');
        });
    }
    const printBtn = document.getElementById('saBtnPrint');
    if (printBtn) {
        printBtn.addEventListener('click', () => window.print());
    }
    const groupCalBtn = document.getElementById('saBtnExportGroupCal');
    if (groupCalBtn) {
        groupCalBtn.addEventListener('click', () => {
            exportSyncGroupCalendars().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }
    wirePersonalCalendarSync();
}

function personalCalendarItemsFromState() {
    const scope = scopeFromState(state, { onlyMine: true });
    return filterSchularbeiten(state.items, state.filters, scope).filter((s) => {
        const st = String(s.status || '').toLowerCase();
        return st === 'fixiert' || st === 'beantragt';
    });
}

async function syncMyCalendarFromUi() {
    const email = String(state.accountEmail || '').trim().toLowerCase();
    if (!email) {
        toast('Bitte mit Microsoft anmelden.');
        return;
    }
    const syncable = personalCalendarItemsFromState();
    if (!syncable.length) {
        toast('Keine eigenen Termine (fixiert/beantragt) im aktuellen Filter.');
        return;
    }
    if (
        !window.confirm(
            syncable.length +
                ' eigene Termine in Ihren Outlook-Kalender schreiben?\n\n' +
                'Bereits verknüpfte Termine auf diesem Gerät werden aktualisiert.'
        )
    ) {
        return;
    }
    state.loading = true;
    paint();
    try {
        const labels = labelMaps(state.stammdaten);
        const scope = scopeFromState(state, { onlyMine: true });
        const result = await syncSchularbeitenToUserCalendar(syncable, {
            accountEmail: email,
            scope,
            fach: labels.fach,
            klasse: labels.klasse
        });
        let text = result.ok + ' in Ihrem Kalender';
        if (result.fail) text += ', ' + result.fail + ' fehlgeschlagen';
        if (result.errors.length) text += ' – ' + result.errors.slice(0, 2).join('; ');
        toast(text + '.');
    } catch (e) {
        toast(e && e.message ? e.message : String(e));
    } finally {
        state.loading = false;
        paint();
    }
}

function wirePersonalCalendarSync() {
    root.querySelectorAll('[data-sa-sync-my-cal]').forEach((btn) => {
        if (btn.disabled) return;
        btn.addEventListener('click', () => {
            syncMyCalendarFromUi().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    });
}

async function maybeSyncSchultermin(sa) {
    if (!state.settings || !state.settings.syncSchultermine || !state.ctx) return null;
    const labels = labelMaps(state.stammdaten);
    const result = await upsertSchulterminFromSchularbeit(state.ctx, sa, {
        listTitle: state.settings.schultermineList || 'Schultermine',
        fachLabel: labels.fach[sa.fachCode] || sa.fachCode,
        klasseLabel: labels.klasse[sa.klasseCode] || sa.klasseCode
    });
    if (sa.itemId && result.itemId) {
        await updateSchularbeitItem(
            state.ctx,
            sa.itemId,
            {
                ...sa,
                schulterminKey: result.itemId
            },
            graphMapOpts()
        );
        sa.schulterminKey = result.itemId;
    }
    return result;
}

async function maybeSyncClassCalendar(sa) {
    if (!state.settings || !state.settings.syncClassCalendar || !state.ctx) return null;
    const labels = labelMaps(state.stammdaten);
    const result = await upsertGroupCalendarEvent(sa, {
        fachLabel: labels.fach[sa.fachCode] || sa.fachCode,
        klasseLabel: labels.klasse[sa.klasseCode] || sa.klasseCode
    });
    if (sa.itemId && result.eventId) {
        await updateSchularbeitItem(
            state.ctx,
            sa.itemId,
            {
                ...sa,
                teamsCalendarEventId: result.eventId
            },
            graphMapOpts()
        );
        sa.teamsCalendarEventId = result.eventId;
    }
    return result;
}

async function maybeRemoveClassCalendar(sa) {
    if (!sa || !sa.teamsCalendarEventId) return null;
    try {
        return await removeGroupCalendarEventForSa(sa);
    } catch (e) {
        console.warn('Klassenkalender-Löschen:', e);
        return null;
    }
}

async function persistGroupCalendarEventId(sa, eventId) {
    if (!state.ctx || !sa.itemId || !eventId) return;
    await updateSchularbeitItem(
        state.ctx,
        sa.itemId,
        {
            ...sa,
            teamsCalendarEventId: eventId
        },
        graphMapOpts()
    );
    sa.teamsCalendarEventId = eventId;
}

async function runGroupCalendarSync(items, confirmText) {
    if (!state.ctx) {
        toast('Bitte zuerst mit SharePoint verbinden.');
        return;
    }
    const fixed = (items || []).filter((x) => String(x.status || '').toLowerCase() === 'fixiert');
    if (!fixed.length) {
        toast('Keine fixierten Schularbeiten in der Auswahl.');
        return;
    }
    if (!window.confirm(confirmText || fixed.length + ' fixierte Termine in Gruppenkalender schreiben?')) {
        return;
    }
    state.loading = true;
    paint();
    try {
        const labels = labelMaps(state.stammdaten);
        const result = await syncSchularbeitenToGroupCalendars(fixed, {
            fach: labels.fach,
            klasse: labels.klasse,
            persistEventId: (sa, eventId) => persistGroupCalendarEventId(sa, eventId)
        });
        await refreshData({ silent: true });
        let text = result.ok + ' in Gruppenkalender geschrieben';
        if (result.fail) text += ', ' + result.fail + ' fehlgeschlagen';
        if (result.errors.length) text += ' – ' + result.errors.slice(0, 2).join('; ');
        toast(text + '.');
    } catch (e) {
        toast(e && e.message ? e.message : String(e));
    } finally {
        state.loading = false;
        paint();
    }
}

async function bulkSyncClassCalendars() {
    const fixed = (state.items || []).filter((x) => x.status === 'fixiert');
    await runGroupCalendarSync(
        fixed,
        fixed.length + ' fixierte Schularbeit(en) in die jeweiligen Klassenkalender schreiben?'
    );
}

function exportItemsForSync() {
    return filterSchularbeiten(
        state.items,
        state.filters,
        scopeFromState(state, {
            scopeAll: state.role === 'admin',
            onlyMine: state.role === 'lehrer'
        })
    );
}

async function exportSyncGroupCalendars() {
    const items = exportItemsForSync();
    const fixed = items.filter((s) => String(s.status || '').toLowerCase() === 'fixiert');
    const klasse = String((state.filters && state.filters.klasse) || '').trim();
    const labels = labelMaps(state.stammdaten);
    const classCodes = Array.from(new Set(fixed.map((s) => s.klasseCode).filter(Boolean)));
    let confirmText =
        fixed.length +
        ' fixierte Termine in ' +
        (klasse
            ? 'den Gruppenkalender der Klasse ' + (labels.klasse[klasse] || klasse)
            : classCodes.length + ' Klassen-Gruppenkalender') +
        ' schreiben?\n\nBereits verknüpfte Termine (TeamsCalendarEventId) werden aktualisiert.';
    await runGroupCalendarSync(fixed, confirmText);
}

function bindPhase5() {
    const copyUrl = document.getElementById('saBtnCopyUrl');
    if (copyUrl) {
        copyUrl.addEventListener('click', () => {
            const el = document.getElementById('saPlanerUrl');
            const text = el ? el.value : '';
            copyText(text).then(() => toast('URL kopiert.')).catch(() => toast('Kopieren fehlgeschlagen.'));
        });
    }
    const copyEmbed = document.getElementById('saBtnCopyEmbed');
    if (copyEmbed) {
        copyEmbed.addEventListener('click', () => {
            const el = document.getElementById('saEmbedSnippet');
            const text = el ? el.textContent : '';
            copyText(text).then(() => toast('Snippet kopiert.')).catch(() => toast('Kopieren fehlgeschlagen.'));
        });
    }
    const saveSettings = document.getElementById('saBtnSaveSettings');
    if (saveSettings) {
        saveSettings.addEventListener('click', () => {
            const syncEl = document.getElementById('saSyncTermine');
            const listEl = document.getElementById('saSyncList');
            const calEl = document.getElementById('saSyncClassCalendar');
            state.settings = persistPlanerSettings({
                syncSchultermine: !!(syncEl && syncEl.checked),
                schultermineList: listEl ? String(listEl.value || '').trim() || 'Schultermine' : 'Schultermine',
                syncClassCalendar: !!(calEl && calEl.checked)
            });
            toast('Einstellungen gespeichert.');
        });
    }
    const bulkCal = document.getElementById('saBtnSyncClassCal');
    if (bulkCal) {
        bulkCal.addEventListener('click', () => {
            bulkSyncClassCalendars().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }
    wireFachMetaTable();

    const btnAdmin = document.getElementById('saBtnRoleAdmin');
    if (btnAdmin) {
        btnAdmin.addEventListener('click', () => {
            if (state.entraGroupsConfigured && !state.demoRoleOverride) return;
            state.role = 'admin';
            state.roleSource = 'demo';
            persistRole('admin');
            state.roleHint = '';
            state.planerAccessDenied = false;
            paint();
        });
    }

    wireListResetActions();
}

function normalizeFachMetaHex(input, fallback) {
    const raw = String(input || '').trim();
    const m = raw.match(/^#?([0-9a-fA-F]{6})$/);
    if (m) return '#' + m[1].toLowerCase();
    return fallback || '#6366f1';
}

function readFachMetaRowPayload(tr) {
    const picker = tr.querySelector('[data-sa-fm-field="farbePicker"]');
    const text = tr.querySelector('[data-sa-fm-field="farbe"]');
    const farbe = normalizeFachMetaHex(text && text.value, picker && picker.value);
    const pro = Number((tr.querySelector('[data-sa-fm-field="pro"]') || {}).value);
    const dauer = Number((tr.querySelector('[data-sa-fm-field="dauer"]') || {}).value);
    const codeEl = tr.querySelector('[data-sa-fm-field="code"]');
    const code = codeEl
        ? String(codeEl.value || '').trim()
        : String(tr.getAttribute('data-sa-fm-code') || '').trim();
    const subj = (state.stammdaten.subjects || []).find((s) => s.code === code);
    return {
        fachCode: code,
        name: (subj && subj.name) || code,
        farbe,
        hatSchularbeiten: true,
        proSemester: Number.isFinite(pro) ? pro : DEFAULT_FACH_META_PRO_SEMESTER,
        standardDauer: Number.isFinite(dauer) ? dauer : DEFAULT_FACH_META_STANDARD_DAUER,
        schuljahr: state.schuljahr || ''
    };
}

function scheduleFachMetaRowSave(tr) {
    if (!tr) return;
    const key = tr.getAttribute('data-sa-fm-id') || tr.getAttribute('data-sa-fm-draft') || '';
    if (!key) return;
    if (fachMetaSaveTimers.has(key)) clearTimeout(fachMetaSaveTimers.get(key));
    fachMetaSaveTimers.set(
        key,
        setTimeout(() => {
            fachMetaSaveTimers.delete(key);
            saveFachMetaRow(tr).catch((e) => toast(e && e.message ? e.message : String(e)));
        }, 450)
    );
}

async function saveFachMetaRow(tr) {
    const payload = readFachMetaRowPayload(tr);
    if (!payload.fachCode) return;

    const itemId = tr.getAttribute('data-sa-fm-id');
    const draftId = tr.getAttribute('data-sa-fm-draft');
    const isNew = Boolean(draftId) || !itemId;

    const applyLocal = (row) => {
        const idx = (state.fachMeta || []).findIndex((m) => m.fachCode === payload.fachCode);
        if (idx >= 0) state.fachMeta[idx] = { ...state.fachMeta[idx], ...row };
        else state.fachMeta.push(row);
    };

    if (!state.ctx) {
        if (!state.localDemoOnly) {
            toast('SharePoint-Site laden, um Fach-Meta zu speichern.');
            return;
        }
        const row = {
            itemId: isNew ? 'local-fm-' + payload.fachCode : itemId,
            ...payload
        };
        applyLocal(row);
        if (draftId) {
            state.fachMetaDrafts = (state.fachMetaDrafts || []).filter((d) => d.draftId !== draftId);
            paint();
        }
        return;
    }

    const duplicate = (state.fachMeta || []).find(
        (m) => m.fachCode === payload.fachCode && m.itemId !== itemId
    );
    if (duplicate) {
        toast('Fach „' + payload.fachCode + '“ ist bereits eingetragen.');
        return;
    }

    if (isNew) {
        const created = await createFachMetaItem(state.ctx, payload);
        if (draftId) {
            state.fachMetaDrafts = (state.fachMetaDrafts || []).filter((d) => d.draftId !== draftId);
        }
        applyLocal(created);
        paint();
        return;
    }

    await updateFachMetaItem(state.ctx, itemId, payload);
    const idx = (state.fachMeta || []).findIndex((m) => m.itemId === itemId);
    if (idx >= 0) {
        state.fachMeta[idx] = { ...state.fachMeta[idx], ...payload };
    }
}

function wireFachMetaTable() {
    const addBtn = document.getElementById('saBtnFachMetaAdd');
    if (addBtn) {
        addBtn.addEventListener('click', () => {
            const used = new Set((state.fachMeta || []).map((m) => m.fachCode));
            (state.fachMetaDrafts || []).forEach((d) => {
                if (d.fachCode) used.add(d.fachCode);
            });
            const openDraft = (state.fachMetaDrafts || []).some((d) => !d.fachCode);
            if (openDraft) {
                toast('Bitte zuerst das neue Fach in der offenen Zeile wählen.');
                return;
            }
            const available = (state.stammdaten.subjects || []).filter((s) => s.code && !used.has(s.code));
            if (!available.length) {
                toast('Alle Fächer aus den Stammdaten sind bereits eingetragen.');
                return;
            }
            const n = (state.fachMeta || []).length + (state.fachMetaDrafts || []).length;
            state.fachMetaDrafts = state.fachMetaDrafts || [];
            state.fachMetaDrafts.push({
                draftId: 'fm-draft-' + Date.now(),
                fachCode: '',
                farbe: FACH_META_COLOR_PALETTE[n % FACH_META_COLOR_PALETTE.length],
                proSemester: DEFAULT_FACH_META_PRO_SEMESTER,
                standardDauer: DEFAULT_FACH_META_STANDARD_DAUER
            });
            paint();
        });
    }

    root.querySelectorAll('[data-sa-fm-cancel-draft]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-sa-fm-cancel-draft');
            state.fachMetaDrafts = (state.fachMetaDrafts || []).filter((d) => d.draftId !== id);
            paint();
        });
    });

    const table = root.querySelector('.sa-fm-table');
    if (table) {
        table.addEventListener('change', (ev) => {
            const tr = ev.target.closest('tr[data-sa-fm-row]');
            if (tr) scheduleFachMetaRowSave(tr);
        });
        table.addEventListener('input', (ev) => {
            const t = ev.target;
            if (
                !t.matches(
                    '[data-sa-fm-field="farbePicker"], [data-sa-fm-field="farbe"], [data-sa-fm-field="pro"], [data-sa-fm-field="dauer"]'
                )
            ) {
                return;
            }
            const tr = t.closest('tr[data-sa-fm-row]');
            if (!tr) return;
            if (t.matches('[data-sa-fm-field="farbePicker"]')) {
                const hex = tr.querySelector('[data-sa-fm-field="farbe"]');
                if (hex) hex.value = t.value;
            }
            if (t.matches('[data-sa-fm-field="farbe"]')) {
                const v = String(t.value || '').trim();
                const picker = tr.querySelector('[data-sa-fm-field="farbePicker"]');
                if (picker && /^#?[0-9a-fA-F]{6}$/.test(v)) {
                    picker.value = v.charAt(0) === '#' ? v : '#' + v;
                }
            }
            scheduleFachMetaRowSave(tr);
        });
    }

    root.querySelectorAll('[data-sa-del-fachmeta]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-sa-del-fachmeta');
            if (!id) return;
            if (!window.confirm('Fach-Meta für diese Zeile löschen?')) return;
            const removeLocal = () => {
                state.fachMeta = (state.fachMeta || []).filter((m) => m.itemId !== id);
                paint();
            };
            if (!state.ctx || String(id).startsWith('local-')) {
                removeLocal();
                toast('Entfernt.');
                return;
            }
            deleteFachMetaItem(state.ctx, id)
                .then(() => refreshData({ silent: true }))
                .then(() => toast('Gelöscht.'))
                .catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    });
}

function readSelectedClearListKeys() {
    const boxes = root.querySelectorAll('input.sa-reset-check[name="saClearList"]:checked');
    const keys = [];
    boxes.forEach((el) => {
        const v = String(el.value || '').trim();
        if (v && LIST_KEYS.includes(v)) keys.push(v);
    });
    return keys;
}

function listTitlesForKeys(keys) {
    return keys.map((k) => LIST_TITLES[k] || k).join(', ');
}

function appendResetLog(line) {
    const el = document.getElementById('saResetLog');
    if (!el) return;
    el.hidden = false;
    el.textContent = (el.textContent ? el.textContent + '\n' : '') + line;
}

function clearResetLog() {
    const el = document.getElementById('saResetLog');
    if (!el) return;
    el.textContent = '';
    el.hidden = true;
}

async function runPlanerListClear(listKeys, { strictConfirm } = {}) {
    if (state.role !== 'admin') {
        toast('Nur für Administrator:innen.');
        return;
    }
    if (!state.ctx) {
        toast('Bitte zuerst SharePoint-Site laden.');
        return;
    }
    const keys = (listKeys || []).filter((k) => LIST_KEYS.includes(k));
    if (!keys.length) {
        toast('Keine Listen ausgewählt.');
        return;
    }
    const labels = listTitlesForKeys(keys);
    if (strictConfirm) {
        const typed = window.prompt(
            'Alle Einträge in diesen Listen werden unwiderruflich gelöscht:\n\n' +
                labels +
                '\n\nZum Bestätigen LEEREN eingeben:'
        );
        if (typed !== 'LEEREN') {
            toast('Abgebrochen.');
            return;
        }
    } else if (
        !window.confirm(
            'Alle Einträge in folgenden Listen löschen?\n\n' + labels + '\n\nDies kann nicht rückgängig gemacht werden.'
        )
    ) {
        return;
    }

    clearResetLog();
    state.loading = true;
    paint();
    try {
        const deleted = await emptyPlanerLists(state.ctx, keys, (p) => {
            if (p.phase === 'load') {
                appendResetLog('Lade ' + p.listTitle + ' …');
            } else if (p.total) {
                appendResetLog(p.listTitle + ': ' + p.done + ' / ' + p.total);
            }
        });
        const summary = keys
            .map((k) => (LIST_TITLES[k] || k) + ': ' + (deleted[k] || 0) + ' gelöscht')
            .join(' · ');
        appendResetLog('Fertig. ' + summary);
        state.localDemoOnly = false;
        await refreshData({ silent: true });
        toast('Listen geleert. ' + summary);
    } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        appendResetLog('Fehler: ' + msg);
        toast(msg);
    } finally {
        state.loading = false;
        paint();
    }
}

function clearLocalDemoView() {
    if (!window.confirm('Lokale Demo-Daten in der Anzeige verwerfen? SharePoint bleibt unverändert.')) return;
    state.items = [];
    state.windows = [];
    state.fachMeta = [];
    state.rules = {
        ...createInitialState().rules,
        name: 'Standard',
        itemId: '',
        regelwerkId: 'rw-1',
        aktiv: true
    };
    state.localDemoOnly = false;
    state.detailId = null;
    if (state.ctx) {
        refreshData({ silent: true }).then(() => toast('Lokale Demo entfernt, Daten von SharePoint geladen.'));
    } else {
        state.bootstrapped = false;
        paint();
        toast('Lokale Demo entfernt.');
    }
}

function wireListResetActions() {
    root.querySelectorAll('[data-sa-clear-one]').forEach((btn) => {
        btn.addEventListener('click', (ev) => {
            ev.preventDefault();
            ev.stopPropagation();
            const key = btn.getAttribute('data-sa-clear-one');
            if (!key) return;
            runPlanerListClear([key], { strictConfirm: false });
        });
    });
    const sel = document.getElementById('saBtnClearSelected');
    if (sel) {
        sel.addEventListener('click', () => runPlanerListClear(readSelectedClearListKeys(), { strictConfirm: false }));
    }
    const all = document.getElementById('saBtnClearAllPlaner');
    if (all) {
        all.addEventListener('click', () => runPlanerListClear([...LIST_KEYS], { strictConfirm: true }));
    }
    const local = document.getElementById('saBtnClearLocalDemo');
    if (local) {
        local.addEventListener('click', () => clearLocalDemoView());
    }
}

function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
        return navigator.clipboard.writeText(String(text || ''));
    }
    return Promise.reject(new Error('Clipboard nicht verfügbar'));
}

function bindJsonImport() {
    const wire = (inputId) => {
        const input = document.getElementById(inputId);
        if (!input) return;
        input.addEventListener('change', () => {
            const file = input.files && input.files[0];
            input.value = '';
            if (!file) return;
            const reader = new FileReader();
            reader.onload = () => {
                importDemoJsonText(String(reader.result || ''))
                    .catch((e) => toast(e && e.message ? e.message : String(e)));
            };
            reader.onerror = () => toast('Datei konnte nicht gelesen werden.');
            reader.readAsText(file, 'UTF-8');
        });
    };
    wire('saImportJson');
    wire('saImportJsonAdmin');
}

/**
 * @param {string} text
 */
async function importDemoJsonText(text) {
    const pack = parseDemoImportJson(text);
    const local = packToLocalPlanerState(pack);

    const stamOk = applyDemoStammdatenLocal(local.stammdaten);
    state.stammdaten = loadStammdaten();
    if (!state.stammdaten.subjects.length && local.stammdaten) {
        state.stammdaten = {
            subjects: local.stammdaten.subjects || [],
            classes: local.stammdaten.classes || [],
            teachers: local.stammdaten.teachers || [],
            students: local.stammdaten.students || []
        };
    }
    syncAccount();

    // Demo: als Admin, sonst sieht man ohne passende Lehrer-E-Mail nichts
    state.role = 'admin';
    persistRole('admin');
    state.roleHint = '';

    if (!state.siteUrl && local.siteDefault) {
        state.siteUrl = local.siteDefault;
        persistSiteUrl(state.siteUrl);
    }

    // Immer lokal anzeigen
    state.items = local.items;
    state.windows = local.windows;
    state.fachMeta = local.fachMeta;
    state.rules = local.rules;
    state.bootstrapped = true;
    state.error = '';
    state.localDemoOnly = true;
    state.view = 'dashboard';
    paint();

    const n = (local.counts && local.counts.schularbeiten) || local.items.length;
    toast(
        'Demo geladen (' +
            n +
            ' Schularbeiten)' +
            (stamOk ? ', Stammdaten übernommen' : '') +
            '.'
    );

    const writeSp = window.confirm(
        'Demo lokal geladen.\n\nJetzt auch auf SharePoint schreiben?\n\nSite: ' +
            (state.siteUrl || DEMO_SITE_DEFAULT) +
            '\n\n(Listen müssen existieren – sonst zuerst „Listen“ → Paket anlegen.)'
    );
    if (!writeSp) return;

    const webUrl = state.siteUrl || DEMO_SITE_DEFAULT;
    state.siteUrl = webUrl;
    persistSiteUrl(webUrl);
    state.loading = true;
    state.error = '';
    paint();
    try {
        await seedDemoSchularbeiten(webUrl, (msg) => console.log('[sa-demo]', msg), { pack: local.pack });
        state.localDemoOnly = false;
        await refreshData({ silent: true });
        toast('Demo auf SharePoint geschrieben und neu geladen.');
    } catch (e) {
        state.localDemoOnly = true;
        state.error =
            'Lokal geladen, SharePoint-Schreiben fehlgeschlagen: ' +
            (e && e.message ? e.message : String(e));
        toast(state.error);
        paint();
    } finally {
        state.loading = false;
        paint();
    }
}

function bindDetail() {
    root.querySelectorAll('[data-sa-detail]').forEach((el) => {
        el.addEventListener('click', () => {
            state.detailId = el.getAttribute('data-sa-detail');
            paint();
        });
    });
    const closeDetail = () => {
        state.detailId = null;
        paint();
    };
    const close = document.getElementById('saDetailClose');
    if (close) close.addEventListener('click', closeDetail);
    const close2 = document.getElementById('saDetailClose2');
    if (close2) close2.addEventListener('click', closeDetail);
    const modal = root.querySelector('.sa-modal');
    if (modal) {
        modal.addEventListener('click', (ev) => {
            if (ev.target === modal) closeDetail();
        });
    }
}

function syncAccount() {
    try {
        const info =
            typeof window.ms365AuthGetAccountInfo === 'function' ? window.ms365AuthGetAccountInfo() : null;
        const upn =
            typeof window.ms365AuthGetUserPrincipalName === 'function'
                ? window.ms365AuthGetUserPrincipalName()
                : '';
        state.accountEmail = String((info && (info.username || info.mail)) || upn || '')
            .trim()
            .toLowerCase();
        state.accountName = String((info && info.name) || '').trim();
    } catch {
        state.accountEmail = '';
        state.accountName = '';
    }
    state.teacherMatch = matchTeacherByEmail(state.accountEmail, state.stammdaten.teachers);
    state.studentMatch = matchStudentByEmail(state.accountEmail, state.stammdaten.students);
    if (!state.form.lehrerCode && state.teacherMatch) {
        state.form = emptyForm(defaultTeacherForm());
    }
}

async function ensurePlannerPermissionsConfig() {
    const site = String(state.siteUrl || loadSavedSiteUrl() || '').trim();
    if (!site) return;
    try {
        const listId =
            state.ctx && state.ctx.lists && state.ctx.lists.schularbeiten && state.ctx.lists.schularbeiten.id
                ? String(state.ctx.lists.schularbeiten.id)
                : '';
        await syncPlannerPermissionsFromSite(site, listId || undefined);
        state.entraGroupsConfigured = entraGroupsConfigured(loadEffectivePermissionsConfig());
    } catch {
        /* optional */
    }
}

function publishPermissionsToSharePoint() {
    const site = String(state.siteUrl || '').trim();
    if (!site) return;
    const listId =
        state.ctx && state.ctx.lists && state.ctx.lists.schularbeiten && state.ctx.lists.schularbeiten.id
            ? String(state.ctx.lists.schularbeiten.id)
            : '';
    publishPlannerPermissionsToSite(site, listId || undefined, loadEffectivePermissionsConfig()).catch(() => {});
}

async function tryRecoverPlanerAccess() {
    if (!state.planerAccessDenied) return;
    syncAccount();
    await ensurePlannerPermissionsConfig();
    await applyPlanerRoleFromEntra(state, {
        demoRoleOverride: state.demoRoleOverride,
        preferredDemoRole: state.demoRoleOverride ? state.role : null,
        preferredActiveRole: state.role
    });
    if (!state.planerAccessDenied) return;
    if (state.teacherMatch) {
        state.planerRoles = ['lehrer'];
        state.planerRoleSources = { lehrer: 'stammdaten' };
        state.role = 'lehrer';
        state.roleSource = 'stammdaten';
        state.planerAccessDenied = false;
        state.roleHint = 'Rolle aus Stammdaten (Lehrer-E-Mail).';
    }
}

async function resolvePlanerRole() {
    syncAccount();
    await ensurePlannerPermissionsConfig();
    await applyPlanerRoleFromEntra(state, {
        demoRoleOverride: state.demoRoleOverride,
        preferredDemoRole: state.demoRoleOverride ? state.role : null,
        preferredActiveRole: state.role
    });
    await tryRecoverPlanerAccess();
    const allowed = viewsForRole(state.role).map((v) => v.id);
    if (!allowed.includes(state.view)) state.view = 'dashboard';
}

async function refreshData(opts) {
    const silent = opts && opts.silent;
    state.error = '';
    if (!silent) {
        state.loading = true;
        paint();
    }
    try {
        syncAccount();
        state.stammdaten = loadStammdaten();
        syncAccount();
        if (!state.siteUrl) throw new Error('SharePoint-Site-URL fehlt.');
        state.ctx = await resolvePlanerContext(state.siteUrl);
        await ensurePlannerPermissionsConfig();
        await resolvePlanerRole();
        if (!state.schuljahr) state.schuljahr = loadSchuljahr() || state.schuljahr;
        const data = await loadAllPlanerData(state.ctx, { schuljahr: state.schuljahr });
        state.items = data.items;
        state.windows = data.windows;
        state.windowsAll = data.windowsAll || data.windows;
        state.allRules = data.allRules || [];
        state.fachMeta = data.fachMeta || [];
        state.fachMetaAll = data.fachMetaAll || data.fachMeta || [];
        state.listsMissingFachMeta = !(state.ctx.lists && state.ctx.lists.fachMeta);
        if (data.rules) state.rules = data.rules;
        state.bootstrapped = true;
        state.localDemoOnly = false;
        syncAccount();
        const entraMode = entraGroupsConfigured(loadEffectivePermissionsConfig());
        // Demo/IT ohne Entra-Gruppen: Lehrer-Filter leer → Hinweis (kein Auto-Admin bei Entra)
        if (
            !entraMode &&
            state.items.length &&
            state.role !== 'admin' &&
            state.role !== 'schueler'
        ) {
            const visible = filterSchularbeiten(
                state.items,
                state.filters,
                scopeFromState(state, { onlyMine: true })
            );
            if (!visible.length && !state.teacherMatch) {
                state.role = 'admin';
                persistRole('admin');
                state.roleSource = 'demo';
                state.roleHint =
                    'Anzeige als Admin: Ihre Anmeldung ist keiner Lehrkraft in den Stammdaten zugeordnet.';
            } else if (!visible.length) {
                state.roleHint =
                    state.items.length +
                    ' Schularbeiten geladen, aber keine zu Ihrer Anmeldung – bitte Rolle „Admin“ wählen.';
            } else if (!state.roleHint) {
                state.roleHint = '';
            }
        } else if (state.role === 'schueler') {
            const klasseOk = !!(state.studentMatch || state.demoKlasseCode);
            state.roleHint = klasseOk
                ? ''
                : 'Schüler-Ansicht: bitte Klasse wählen (Demo) oder E-Mail in den Stammdaten (students) hinterlegen.';
        } else {
            state.roleHint = '';
        }
        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function' && !silent) {
            window.ms365ActionLog.append({
                tool: 'schularbeiten-planer',
                action: 'load',
                target: state.siteUrl,
                summary: state.items.length + ' Schularbeiten geladen'
            });
        }
    } catch (e) {
        state.error = e && e.message ? e.message : String(e);
        throw e;
    } finally {
        state.loading = false;
        paint();
    }
}

function boot() {
    root = document.getElementById('saApp');
    if (!root) return;

    wirePermissionGroupPickersDelegated(root, PLANER_GROUP_FIELDS, () => {
        persistPickersToStorage(PLANER_GROUP_FIELDS);
        publishPermissionsToSharePoint();
    });

    state.siteUrl = loadSavedSiteUrl() || DEMO_SITE_DEFAULT;
    state.role = resolveRole();
    state.stammdaten = loadStammdaten();
    syncAccount();
    state.form = emptyForm(defaultTeacherForm());

    const params = new URLSearchParams(window.location.search || '');
    state.demoRoleOverride = params.get('demoRole') === '1';
    state.entraGroupsConfigured = entraGroupsConfigured(loadEffectivePermissionsConfig());

    const view = params.get('view');
    if (view && VIEWS.some((v) => v.id === view)) state.view = view;
    const roleQ = params.get('role');
    if (roleQ === 'admin' || roleQ === 'lehrer' || roleQ === 'schueler') {
        if (isPlanerDemoRoleUiEnabled(state.entraGroupsConfigured, state.demoRoleOverride)) {
            state.role = roleQ;
            persistRole(roleQ);
            state.roleSource = 'demo';
        }
    }
    const klasseQ = params.get('klasse');
    if (klasseQ) {
        state.demoKlasseCode = String(klasseQ).trim();
        persistDemoKlasseCode(state.demoKlasseCode);
    }
    if (state.role === 'schueler' && !viewsForRole('schueler').some((v) => v.id === state.view)) {
        state.view = 'dashboard';
    }

    paint();

    window.addEventListener('ms365-auth-widget-ready', placePlanerAuthWidget);

    window.addEventListener('ms365-auth-state-changed', () => {
        resolvePlanerRole()
            .then(() => {
                if (state.bootstrapped && state.siteUrl) {
                    return refreshData({ silent: true });
                }
                paint();
            })
            .catch((e) => {
                state.error = e && e.message ? e.message : String(e);
                paint();
            });
    });

    if (state.siteUrl) {
        refreshData().catch(() => {
            /* error already in state */
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}
