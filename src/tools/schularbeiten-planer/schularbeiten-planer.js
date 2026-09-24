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
    loadSavedSiteUrl,
    persistSiteUrl,
    emptyForm,
    filterSchularbeiten,
    labelMaps,
    persistPlanerSettings,
    viewsForRole,
    VIEWS,
    scopeFromState
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
    deleteFachMetaItem
} from './schularbeiten-planer-graph.js';
import { upsertSchulterminFromSchularbeit } from './schularbeiten-planer-sync.js';
import {
    seedDemoSchularbeiten,
    parseDemoImportJson,
    applyDemoStammdatenLocal,
    packToLocalPlanerState
} from './schularbeiten-planer-demo-seed.js';
import { newEntityId } from './schularbeiten-planer-schema.js';
import { DEMO_SITE_DEFAULT } from './schularbeiten-planer-demo-data.js';
import {
    renderApp,
    readFormFromDom,
    readFiltersFromDom,
    readFensterForm,
    readRulesForm
} from './schularbeiten-planer-ui.js';
import { buildIcs, downloadIcs } from './schularbeiten-planer-export.js';
import { validateSchularbeit } from './schularbeiten-planer-logic.js';

const state = createInitialState();
let root = null;

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else window.alert(msg);
}

function paint() {
    if (!root) return;
    renderApp(state, root);
    bindStatic();
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
            state.role = role;
            persistRole(role);
            state.roleHint = '';
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

    const filterIds = ['saFilterKlasse', 'saFilterFach', 'saFilterLehrer', 'saFilterStatus'];
    filterIds.forEach((id) => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('change', () => {
                state.filters = readFiltersFromDom();
                paint();
            });
        }
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
}

function bindFormLive() {
    const ids = ['saThema', 'saFach', 'saKlasse', 'saDatum', 'saDauer', 'saSemester', 'saLehrer', 'saNotiz'];
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
    const draft = {
        schularbeitId: state.form.schularbeitId || newEntityId('sa'),
        thema: form.thema,
        fachCode: form.fachCode,
        klasseCode: form.klasseCode,
        lehrerCode: form.lehrerCode,
        lehrerEmail: form.lehrerEmail,
        datum: form.datum,
        dauerMinuten: form.dauerMinuten,
        semester: form.semester,
        notiz: form.notiz,
        status: 'beantragt',
        beantragtVon: state.accountEmail || ''
    };
    const check = validateSchularbeit({
        draft,
        existing: state.items,
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
    try {
        if (state.editingItemId) {
            await updateSchularbeitItem(state.ctx, state.editingItemId, draft);
            toast('Antrag aktualisiert.');
        } else {
            await createSchularbeitItem(state.ctx, draft);
            toast('Antrag eingereicht.');
        }
        state.editingItemId = null;
        state.form = emptyForm(defaultTeacherForm());
        state.view = 'meine';
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
            if (!sa || sa.status !== 'beantragt') return;
            state.detailId = null;
            state.editingItemId = id;
            state.form = emptyForm({
                thema: sa.thema,
                fachCode: sa.fachCode,
                klasseCode: sa.klasseCode,
                lehrerCode: sa.lehrerCode,
                lehrerEmail: sa.lehrerEmail,
                datum: sa.datum,
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
            if (!id || !state.ctx) return;
            if (!window.confirm('Antrag wirklich löschen?')) return;
            state.detailId = null;
            deleteSchularbeitItem(state.ctx, id)
                .then(() => refreshData({ silent: true }))
                .then(() => toast('Gelöscht.'))
                .catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    });
}

function bindAdminActions() {
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
            updateSchularbeitItem(state.ctx, id, patch)
                .then(() => maybeSyncSchultermin(patch))
                .then((syncResult) =>
                    refreshData({ silent: true }).then(() => {
                        if (syncResult && syncResult.created) toast('Fixiert und in Schultermine angelegt.');
                        else if (syncResult) toast('Fixiert und Schultermine aktualisiert.');
                        else toast('Fixiert.');
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
            updateSchularbeitItem(state.ctx, id, patch)
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
}

function bindKalender() {
    const prev = document.getElementById('saCalPrev');
    const next = document.getElementById('saCalNext');
    const today = document.getElementById('saCalToday');
    if (prev) {
        prev.addEventListener('click', () => {
            state.calMonth -= 1;
            if (state.calMonth < 0) {
                state.calMonth = 11;
                state.calYear -= 1;
            }
            paint();
        });
    }
    if (next) {
        next.addEventListener('click', () => {
            state.calMonth += 1;
            if (state.calMonth > 11) {
                state.calMonth = 0;
                state.calYear += 1;
            }
            paint();
        });
    }
    if (today) {
        today.addEventListener('click', () => {
            const n = new Date();
            state.calYear = n.getFullYear();
            state.calMonth = n.getMonth();
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
        await updateSchularbeitItem(state.ctx, sa.itemId, {
            ...sa,
            schulterminKey: result.itemId
        });
    }
    return result;
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
            state.settings = persistPlanerSettings({
                syncSchultermine: !!(syncEl && syncEl.checked),
                schultermineList: listEl ? String(listEl.value || '').trim() || 'Schultermine' : 'Schultermine'
            });
            toast('Einstellungen gespeichert.');
        });
    }
    const addMeta = document.getElementById('saBtnAddFachMeta');
    if (addMeta) {
        addMeta.addEventListener('click', () => {
            const code = (document.getElementById('saFmCode') || {}).value || '';
            const farbe = (document.getElementById('saFmFarbe') || {}).value || '';
            const pro = Number((document.getElementById('saFmPro') || {}).value) || 2;
            const dauer = Number((document.getElementById('saFmDauer') || {}).value) || 100;
            if (!code || !state.ctx) {
                toast('Fach-Code wählen.');
                return;
            }
            const subj = state.stammdaten.subjects.find((s) => s.code === code);
            const payload = {
                fachCode: code,
                name: (subj && subj.name) || code,
                farbe: String(farbe).trim(),
                hatSchularbeiten: true,
                proSemester: pro,
                standardDauer: dauer
            };
            const existing = (state.fachMeta || []).find((m) => m.fachCode === code);
            const job = existing
                ? updateFachMetaItem(state.ctx, existing.itemId, payload)
                : createFachMetaItem(state.ctx, payload);
            job
                .then(() => refreshData({ silent: true }))
                .then(() => toast(existing ? 'Fach-Meta aktualisiert.' : 'Fach-Meta angelegt.'))
                .catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }
    const farbePicker = document.getElementById('saFmFarbePicker');
    const farbeText = document.getElementById('saFmFarbe');
    if (farbePicker && farbeText) {
        farbePicker.addEventListener('input', () => {
            farbeText.value = farbePicker.value;
        });
        farbeText.addEventListener('change', () => {
            const v = String(farbeText.value || '').trim();
            if (/^#?[0-9a-fA-F]{6}$/.test(v)) {
                farbePicker.value = v.charAt(0) === '#' ? v : '#' + v;
            }
        });
    }
    root.querySelectorAll('[data-sa-del-fachmeta]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-sa-del-fachmeta');
            if (!id || !state.ctx) return;
            if (!window.confirm('Fach-Meta löschen?')) return;
            deleteFachMetaItem(state.ctx, id)
                .then(() => refreshData({ silent: true }))
                .then(() => toast('Gelöscht.'))
                .catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    });

    const btnAdmin = document.getElementById('saBtnRoleAdmin');
    if (btnAdmin) {
        btnAdmin.addEventListener('click', () => {
            state.role = 'admin';
            persistRole('admin');
            state.roleHint = '';
            paint();
        });
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
        if (!state.siteUrl) throw new Error('SharePoint-Site-URL fehlt.');
        state.ctx = await resolvePlanerContext(state.siteUrl);
        const data = await loadAllPlanerData(state.ctx);
        state.items = data.items;
        state.windows = data.windows;
        state.fachMeta = data.fachMeta || [];
        state.listsMissingFachMeta = !(state.ctx.lists && state.ctx.lists.fachMeta);
        if (data.rules) state.rules = data.rules;
        state.bootstrapped = true;
        state.localDemoOnly = false;
        syncAccount();
        // Demo/IT: Daten vorhanden, aber Lehrer-Rolle filtert alles weg → Admin + Hinweis
        // Schüler-Rolle nie automatisch auf Admin umschalten
        if (state.items.length && state.role !== 'admin' && state.role !== 'schueler') {
            const visible = filterSchularbeiten(
                state.items,
                state.filters,
                scopeFromState(state, { onlyMine: true })
            );
            if (!visible.length && !state.teacherMatch) {
                state.role = 'admin';
                persistRole('admin');
                state.roleHint =
                    'Anzeige als Admin: Ihre Anmeldung ist keiner Lehrkraft in den Stammdaten zugeordnet.';
            } else if (!visible.length) {
                state.roleHint =
                    state.items.length +
                    ' Schularbeiten geladen, aber keine zu Ihrer Anmeldung – bitte Rolle „Admin“ wählen.';
            } else {
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

    state.siteUrl = loadSavedSiteUrl() || DEMO_SITE_DEFAULT;
    state.role = resolveRole();
    state.stammdaten = loadStammdaten();
    syncAccount();
    state.form = emptyForm(defaultTeacherForm());

    const params = new URLSearchParams(window.location.search || '');
    const view = params.get('view');
    if (view && VIEWS.some((v) => v.id === view)) state.view = view;
    const roleQ = params.get('role');
    if (roleQ === 'admin' || roleQ === 'lehrer' || roleQ === 'schueler') {
        state.role = roleQ;
        persistRole(roleQ);
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
