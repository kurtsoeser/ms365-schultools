/**
 * Entry / Wiring: Projektwochen (Phase 2–4)
 */
import {
    createInitialState,
    loadStammdaten,
    matchTeacherByEmail,
    loadSavedSiteUrl,
    persistSiteUrl,
    persistRole,
    emptyForm,
    formFromItem,
    pickActiveAktion,
    viewsForRole,
    canEditAngebot,
    scopeFromState,
    filterAngebote
} from './projektwochen-state.js';
import {
    resolvePwContext,
    loadAllPwData,
    createAngebotItem,
    updateAngebotItem,
    deleteAngebotItem,
    updateAktionItem
} from './projektwochen-graph.js';
import { validateAngebot, weekdayLabelDeFromIso } from './projektwochen-logic.js';
import { renderApp, readFormFromDom, readFiltersFromDom } from './projektwochen-ui.js';
import { buildLocalDemoState, getDemoSeedPackage, DEMO_SITE_DEFAULT } from './projektwochen-demo-data.js';
import {
    parseDemoImportJson,
    applyDemoStammdatenLocal,
    seedDemoProjektwochen
} from './projektwochen-demo-seed.js';
import {
    createProjektwochenLists,
    probeListsHealth
} from '../sharepoint/sharepoint-liste-projektwochen.js';
import {
    ensureBookingBusiness,
    syncAngebotToBookings,
    loadBookingsAttendees
} from './projektwochen-bookings.js';
import {
    downloadAttendeesCsv,
    downloadAngeboteCsv,
    buildAngeboteIcs,
    downloadIcs,
    buildWeekPlanPrintHtml,
    buildAttendeesPrintHtml,
    openPrintWindow
} from './projektwochen-export.js';
import { newEntityId } from './projektwochen-schema.js';

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
    root.querySelectorAll('[data-pw-view]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-pw-view') || 'dashboard';
            state.detailId = null;
            paint();
        });
    });

    root.querySelectorAll('[data-pw-role]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const raw = btn.getAttribute('data-pw-role') || 'lehrer';
            const role = raw === 'admin' ? 'admin' : raw === 'schueler' ? 'schueler' : 'lehrer';
            state.role = role;
            persistRole(role);
            state.roleHint = '';
            const allowed = viewsForRole(role).map((v) => v.id);
            if (!allowed.includes(state.view)) state.view = 'dashboard';
            paint();
        });
    });

    const loadBtn = document.getElementById('pwBtnLoad');
    if (loadBtn) {
        loadBtn.addEventListener('click', () => {
            const url = String((document.getElementById('pwSiteUrl') && document.getElementById('pwSiteUrl').value) || '').trim();
            state.siteUrl = url;
            persistSiteUrl(url);
            refreshData().catch((e) => {
                state.error = e && e.message ? e.message : String(e);
                state.loading = false;
                paint();
                toast(state.error);
            });
        });
    }

    const reloadBtn = document.getElementById('pwBtnReload');
    if (reloadBtn) {
        reloadBtn.addEventListener('click', () => {
            if (!state.siteUrl) {
                state.view = 'einstellungen';
                paint();
                toast('Bitte zuerst unter Einstellungen die SharePoint-Website hinterlegen.');
                return;
            }
            refreshData().catch((e) => {
                state.error = e && e.message ? e.message : String(e);
                state.loading = false;
                paint();
                toast(state.error);
            });
        });
    }

    const aktionSelect = document.getElementById('pwAktionSelect');
    if (aktionSelect) {
        aktionSelect.addEventListener('change', () => {
            const id = String(aktionSelect.value || '').trim();
            const found = (state.aktionen || []).find((a) => a && String(a.aktionId || '') === id);
            if (found) {
                state.aktion = found;
                paint();
            }
        });
    }

    const listsCreate = document.getElementById('pwBtnListsCreate');
    if (listsCreate) {
        listsCreate.addEventListener('click', () => {
            runListsCreate().catch((e) => toast(String((e && e.message) || e)));
        });
    }
    const listsHealth = document.getElementById('pwBtnListsHealth');
    if (listsHealth) {
        listsHealth.addEventListener('click', () => {
            runListsHealth().catch((e) => toast(String((e && e.message) || e)));
        });
    }

    const demoBtn = document.getElementById('pwBtnDemo');
    if (demoBtn) {
        demoBtn.addEventListener('click', () => {
            runDemoImport(getDemoSeedPackage()).catch((e) => toast(String((e && e.message) || e)));
        });
    }

    const importAdmin = document.getElementById('pwImportJsonAdmin');
    if (importAdmin) {
        importAdmin.addEventListener('change', () => {
            const file = importAdmin.files && importAdmin.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = () => {
                importDemoJsonText(String(reader.result || ''))
                    .catch((e) => toast(String((e && e.message) || e)))
                    .finally(() => {
                        importAdmin.value = '';
                    });
            };
            reader.readAsText(file, 'UTF-8');
        });
    }

    ['pwFilterStatus', 'pwFilterKat', 'pwFilterTag', 'pwFilterLehrer', 'pwFilterKlasse', 'pwFilterBuchung'].forEach(
        (id) => {
            const el = document.getElementById(id);
            if (el) {
                el.addEventListener('change', () => {
                    state.filters = readFiltersFromDom();
                    paint();
                });
            }
        }
    );
    const q = document.getElementById('pwFilterQ');
    if (q) {
        q.addEventListener('input', () => {
            state.filters = readFiltersFromDom();
            paint();
        });
    }
    const filterReset = document.getElementById('pwFilterReset');
    if (filterReset) {
        filterReset.addEventListener('click', () => {
            state.filters = { status: '', kategorie: '', tag: '', lehrer: '', klasse: '', q: '', buchung: '' };
            paint();
        });
    }

    const prev = document.getElementById('pwCalPrev');
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
    const next = document.getElementById('pwCalNext');
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

    root.querySelectorAll('[data-pw-detail]').forEach((btn) => {
        btn.addEventListener('click', () => {
            openDetail(btn.getAttribute('data-pw-detail'));
        });
    });

    const detailBack = document.getElementById('pwDetailBack');
    if (detailBack) {
        detailBack.addEventListener('click', () => {
            closeDetail();
        });
    }

    const detailLoadTn = document.getElementById('pwBtnDetailLoadTn');
    if (detailLoadTn) {
        detailLoadTn.addEventListener('click', () => {
            runLoadAttendees({ stayOnDetail: true }).catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    }

    root.querySelectorAll('[data-pw-edit]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-pw-edit');
            openDetail(id);
        });
    });

    root.querySelectorAll('[data-pw-del]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-pw-del');
            if (!window.confirm('Angebot löschen?')) return;
            deleteAngebot(id).catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    });

    root.querySelectorAll('[data-pw-approve]').forEach((btn) => {
        btn.addEventListener('click', () => {
            decide(btn.getAttribute('data-pw-approve'), 'freigegeben').catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });

    root.querySelectorAll('[data-pw-reject]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const reason = window.prompt('Ablehnungsgrund (optional):', '') || '';
            decide(btn.getAttribute('data-pw-reject'), 'abgelehnt', reason).catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });

    root.querySelectorAll('[data-pw-buchung-save]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-pw-buchung-save');
            const input = root.querySelector('[data-pw-buchung-id="' + id + '"]');
            const val = input ? String(input.value || '').trim() : '';
            saveBuchungAb(id, val).catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    });

    const saveAktion = document.getElementById('pwBtnAktionSave');
    if (saveAktion) {
        saveAktion.addEventListener('click', () => {
            saveAktionDefault().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }

    const saveBtn = document.getElementById('pwBtnSave');
    if (saveBtn) {
        saveBtn.addEventListener('click', () => {
            saveForm().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }

    const cancelEdit = document.getElementById('pwBtnCancelEdit');
    if (cancelEdit) {
        cancelEdit.addEventListener('click', () => {
            if (state.view === 'detail') {
                closeDetail();
                return;
            }
            state.editingItemId = null;
            state.form = emptyForm(prefillTeacher());
            state.view = 'meine';
            paint();
        });
    }

    const ensureBiz = document.getElementById('pwBtnEnsureBiz');
    if (ensureBiz) {
        ensureBiz.addEventListener('click', () => {
            runEnsureBusiness().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }
    const syncAll = document.getElementById('pwBtnSyncAll');
    if (syncAll) {
        syncAll.addEventListener('click', () => {
            const ids = state.items.filter((a) => a.status === 'freigegeben').map((a) => a.itemId);
            runSyncAngebote(ids).catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }
    const syncSel = document.getElementById('pwBtnSyncSelected');
    if (syncSel) {
        syncSel.addEventListener('click', () => {
            const ids = Array.from(root.querySelectorAll('.pw-sync-check:checked')).map((el) =>
                el.getAttribute('data-pw-sync-id')
            );
            runSyncAngebote(ids).catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }
    root.querySelectorAll('[data-pw-sync-one]').forEach((btn) => {
        btn.addEventListener('click', () => {
            runSyncAngebote([btn.getAttribute('data-pw-sync-one')]).catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });
    const loadTn = document.getElementById('pwBtnLoadTn');
    if (loadTn) {
        loadTn.addEventListener('click', () => {
            runLoadAttendees().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }
    const loadTn2 = document.getElementById('pwBtnLoadTn2');
    if (loadTn2) {
        loadTn2.addEventListener('click', () => {
            runLoadAttendees().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }
    const csvBtn = document.getElementById('pwBtnTnCsv');
    if (csvBtn) {
        csvBtn.addEventListener('click', () => {
            downloadAttendeesCsv(state.attendeeRows || [], 'projektwochen-teilnehmer.csv');
            toast('CSV exportiert.');
        });
    }
    const tnAngebot = document.getElementById('pwTnAngebot');
    if (tnAngebot) {
        tnAngebot.addEventListener('change', () => {
            state.attendeeFilter = {
                ...state.attendeeFilter,
                angebotId: String(tnAngebot.value || '')
            };
            paint();
        });
    }
    const tnKlasse = document.getElementById('pwTnKlasse');
    if (tnKlasse) {
        tnKlasse.addEventListener('change', () => {
            state.attendeeFilter = {
                ...state.attendeeFilter,
                klasse: String(tnKlasse.value || '')
            };
            paint();
        });
    }
    const tnQ = document.getElementById('pwTnQ');
    if (tnQ) {
        tnQ.addEventListener('input', () => {
            state.attendeeFilter = { ...state.attendeeFilter, q: String(tnQ.value || '') };
            paint();
        });
    }

    const expAng = document.getElementById('pwBtnExportAngebote');
    if (expAng) {
        expAng.addEventListener('click', () => {
            const items = filterAngebote(state.items, state.filters, scopeFromState(state), state.aktion);
            downloadAngeboteCsv(items);
            toast('Angebote-CSV exportiert.');
        });
    }
    const expIcs = document.getElementById('pwBtnExportIcs');
    if (expIcs) {
        expIcs.addEventListener('click', () => {
            const items = filterAngebote(state.items, state.filters, scopeFromState(state), state.aktion).filter(
                (a) => a.status === 'freigegeben'
            );
            const ics = buildAngeboteIcs(items, state.aktion);
            downloadIcs(ics, 'projektwochen.ics');
            toast('ICS exportiert (' + items.length + ').');
        });
    }
    const printPlan = document.getElementById('pwBtnPrintPlan');
    if (printPlan) {
        printPlan.addEventListener('click', () => {
            const items = filterAngebote(state.items, state.filters, scopeFromState(state), state.aktion);
            try {
                openPrintWindow(buildWeekPlanPrintHtml(items, state.aktion));
            } catch (e) {
                toast(e && e.message ? e.message : String(e));
            }
        });
    }
    const expTn = document.getElementById('pwBtnExportTn');
    if (expTn) {
        expTn.addEventListener('click', () => {
            downloadAttendeesCsv(state.attendeeRows || []);
            toast('Teilnehmer-CSV exportiert.');
        });
    }
    const printTn = document.getElementById('pwBtnPrintTn');
    if (printTn) {
        printTn.addEventListener('click', () => {
            try {
                openPrintWindow(buildAttendeesPrintHtml(state.attendeeRows || [], state.aktion));
            } catch (e) {
                toast(e && e.message ? e.message : String(e));
            }
        });
    }
}

function appendBookingsLog(msg) {
    state.bookingsLog = (state.bookingsLog ? state.bookingsLog + '\n' : '') + String(msg || '');
}

function prefillTeacher() {
    const t = state.teacherMatch;
    if (!t) return {};
    return {
        lehrerCode: t.code || '',
        lehrerEmail: t.email || state.accountEmail || ''
    };
}

function openDetail(rawId) {
    const id = String(rawId || '').trim();
    const item = state.items.find((a) => a.itemId === id || a.angebotId === id);
    if (!item) {
        toast('Angebot nicht gefunden.');
        return;
    }
    if (state.view !== 'detail') state.detailReturnView = state.view || 'dashboard';
    state.detailId = item.itemId || item.angebotId;
    state.view = 'detail';
    if (canEditAngebot(item, scopeFromState(state))) {
        state.editingItemId = item.itemId;
        state.form = formFromItem(item);
    } else {
        state.editingItemId = null;
    }
    paint();
}

function closeDetail() {
    const back = state.detailReturnView && state.detailReturnView !== 'detail' ? state.detailReturnView : 'dashboard';
    state.detailId = null;
    state.editingItemId = null;
    state.form = emptyForm(prefillTeacher());
    state.view = back;
    paint();
}

function applyLocalDemo(pack) {
    const packed = buildLocalDemoState(pack);
    state.aktionen = packed.aktionen;
    state.aktion = pickActiveAktion(packed.aktionen, packed.angebote);
    state.items = packed.angebote;
    if (packed.stammdaten) state.stammdaten = packed.stammdaten;
    state.localDemoOnly = true;
    state.bootstrapped = true;
    state.ctx = null;
    state.error = '';
    if (state.aktion && state.aktion.startdatum) {
        const d = new Date(state.aktion.startdatum + 'T12:00:00');
        if (!Number.isNaN(d.getTime())) {
            state.calYear = d.getFullYear();
            state.calMonth = d.getMonth();
        }
    }
    state.form = emptyForm(prefillTeacher());
    paint();
}

async function runListsCreate() {
    const url = String((document.getElementById('pwSiteUrl') && document.getElementById('pwSiteUrl').value) || state.siteUrl || '').trim();
    if (!url) {
        toast('Bitte SharePoint-Website eintragen.');
        return;
    }
    if (
        !window.confirm(
            'Projektwochen-Listen auf der Website anlegen bzw. fehlende Spalten ergänzen?\n\n' +
                url +
                '\n\nListen: PW-Aktionen, PW-Angebote'
        )
    ) {
        return;
    }
    state.siteUrl = url;
    persistSiteUrl(url);
    const logEl = document.getElementById('pwSetupLog');
    if (logEl) {
        logEl.hidden = false;
        logEl.textContent = '';
    }
    const write = (msg) => {
        if (logEl) {
            logEl.textContent += (logEl.textContent ? '\n' : '') + msg;
            logEl.scrollTop = logEl.scrollHeight;
        }
        console.log('[pw-lists]', msg);
    };
    await createProjektwochenLists(url, write);
    toast('Listen bereit.');
    await refreshData();
}

async function runListsHealth() {
    const url = String((document.getElementById('pwSiteUrl') && document.getElementById('pwSiteUrl').value) || state.siteUrl || '').trim();
    if (!url) {
        toast('Bitte SharePoint-Website eintragen.');
        return;
    }
    state.siteUrl = url;
    persistSiteUrl(url);
    const logEl = document.getElementById('pwSetupLog');
    if (logEl) {
        logEl.hidden = false;
        logEl.textContent = '';
    }
    const write = (msg) => {
        if (logEl) {
            logEl.textContent += (logEl.textContent ? '\n' : '') + msg;
            logEl.scrollTop = logEl.scrollHeight;
        }
    };
    const summary = await probeListsHealth(url, write);
    toast(summary && summary.ok ? 'Listen-Check OK' : 'Listen-Check: bitte Protokoll prüfen');
}

async function importDemoJsonText(text) {
    const pack = parseDemoImportJson(text);
    await runDemoImport(pack);
}

async function runDemoImport(pack) {
    const p = parseDemoImportJson(pack);
    const stamOk = applyDemoStammdatenLocal(p.stammdaten);
    applyLocalDemo(p);
    state.role = 'admin';
    persistRole('admin');
    if (!state.siteUrl && p.siteDefault) {
        state.siteUrl = String(p.siteDefault);
        persistSiteUrl(state.siteUrl);
    }
    paint();
    toast(
        'Demo lokal geladen (' +
            (p.counts && p.counts.angebote != null ? p.counts.angebote : p.angebote.length) +
            ' Angebote)' +
            (stamOk ? ' · Stammdaten übernommen' : '') +
            '.'
    );
    const webUrl = String(state.siteUrl || p.siteDefault || DEMO_SITE_DEFAULT).trim();
    const writeSpo = window.confirm(
        'Demo lokal geladen.\n\nJetzt auch auf SharePoint schreiben?\n\nSite: ' +
            webUrl +
            '\n\nVoraussetzung: Listen PW-Aktionen / PW-Angebote existieren.\nStammdaten liegen im Browser-Backup (tenant-settings).'
    );
    if (!writeSpo) return;
    state.loading = true;
    state.error = '';
    paint();
    try {
        await seedDemoProjektwochen(webUrl, (msg) => console.log('[pw-demo]', msg), { pack: p });
        state.siteUrl = webUrl;
        persistSiteUrl(webUrl);
        await refreshData();
        toast('Demo auf SharePoint geschrieben und neu geladen.');
    } catch (e) {
        state.localDemoOnly = true;
        state.loading = false;
        state.error = String((e && e.message) || e);
        paint();
        toast('SharePoint-Schreiben fehlgeschlagen – lokal bleibt Demo aktiv: ' + state.error);
    }
}

async function syncAccount() {
    try {
        if (typeof window.ms365AuthGetAccount === 'function') {
            const acc = window.ms365AuthGetAccount();
            if (acc) {
                state.accountEmail = String(acc.username || acc.mail || '').trim().toLowerCase();
                state.accountName = String(acc.name || '').trim();
            }
        }
    } catch {
        /* ignore */
    }
    state.stammdaten = loadStammdaten();
    state.teacherMatch = matchTeacherByEmail(state.accountEmail, state.stammdaten.teachers);
    if (state.role === 'lehrer' && !state.teacherMatch && state.bootstrapped && !state.localDemoOnly) {
        /* soft hint only */
    }
}

async function refreshData() {
    state.loading = true;
    state.error = '';
    paint();
    await syncAccount();
    const ctx = await resolvePwContext(state.siteUrl);
    state.ctx = ctx;
    const data = await loadAllPwData(ctx);
    state.aktionen = data.aktionen;
    state.aktion = pickActiveAktion(data.aktionen, data.angebote);
    state.items = data.angebote;
    state.localDemoOnly = false;
    state.bootstrapped = true;
    state.loading = false;
    if (state.aktion && state.aktion.startdatum) {
        const d = new Date(state.aktion.startdatum + 'T12:00:00');
        if (!Number.isNaN(d.getTime())) {
            state.calYear = d.getFullYear();
            state.calMonth = d.getMonth();
        }
    }
    if (!state.aktion) {
        state.roleHint = 'Keine PW-Aktion gefunden – bitte Setup/Seed prüfen.';
    }
    state.form = emptyForm(prefillTeacher());
    paint();
    toast('Projektwochen geladen (' + state.items.length + ' Angebote).');
}

async function saveForm() {
    const form = { ...state.form, ...readFormFromDom() };
    if (!form.tag && form.datum) form.tag = weekdayLabelDeFromIso(form.datum);
    if (!form.lehrerEmail && state.accountEmail) form.lehrerEmail = state.accountEmail;
    if (!form.lehrerCode && state.teacherMatch) form.lehrerCode = state.teacherMatch.code;
    state.form = form;

    const classCodes = ((state.stammdaten && state.stammdaten.classes) || []).map((c) => c.code).filter(Boolean);
    const peers = state.items.filter((i) => i.itemId !== state.editingItemId);
    const v = validateAngebot({
        draft: form,
        existing: peers,
        aktion: state.aktion,
        classCodes
    });
    if (!v.canSubmit) {
        paint();
        toast(v.errors[0] || 'Bitte Fehler prüfen.');
        return;
    }

    const payload = {
        ...form,
        angebotId: form.angebotId || newEntityId('ang'),
        aktionId: (state.aktion && state.aktion.aktionId) || '',
        beantragtVon: state.accountEmail || form.lehrerEmail || '',
        buchungAb: form.buchungAb || (state.aktion && state.aktion.buchungAbDefault) || ''
    };
    const existingEdit = state.editingItemId
        ? state.items.find((i) => i.itemId === state.editingItemId)
        : null;
    if (existingEdit && existingEdit.status && existingEdit.status !== 'entwurf') {
        payload.status = existingEdit.status === 'beantragt' ? 'beantragt' : existingEdit.status;
    } else {
        payload.status = 'beantragt';
    }
    const stayDetail = state.view === 'detail';

    if (state.localDemoOnly || !state.ctx) {
        if (state.editingItemId) {
            state.items = state.items.map((it) =>
                it.itemId === state.editingItemId ? { ...it, ...payload, itemId: state.editingItemId } : it
            );
        } else {
            state.items = state.items.concat([{ ...payload, itemId: 'local-' + payload.angebotId }]);
        }
        if (stayDetail) {
            state.form = formFromItem(
                state.items.find((i) => i.itemId === state.editingItemId) || payload
            );
            paint();
            toast('Lokal gespeichert (Demo).');
            return;
        }
        state.editingItemId = null;
        state.form = emptyForm(prefillTeacher());
        state.view = 'meine';
        paint();
        toast('Lokal gespeichert (Demo).');
        return;
    }

    if (state.editingItemId) {
        const existing = state.items.find((i) => i.itemId === state.editingItemId);
        await updateAngebotItem(state.ctx, state.editingItemId, {
            ...existing,
            ...payload,
            status:
                existing && existing.status !== 'entwurf' && existing.status !== 'beantragt'
                    ? existing.status
                    : payload.status
        });
    } else {
        await createAngebotItem(state.ctx, payload);
    }
    const keepId = state.editingItemId || state.detailId;
    await refreshData();
    if (stayDetail && keepId) {
        const again = state.items.find((a) => a.itemId === keepId || a.angebotId === keepId);
        state.view = 'detail';
        state.detailId = again ? again.itemId : keepId;
        if (again && canEditAngebot(again, scopeFromState(state))) {
            state.editingItemId = again.itemId;
            state.form = formFromItem(again);
        }
        paint();
        toast('Angebot gespeichert.');
        return;
    }
    state.editingItemId = null;
    state.view = 'meine';
    paint();
    toast('Angebot gespeichert.');
}

async function deleteAngebot(itemId) {
    const item = state.items.find((a) => a.itemId === itemId);
    if (!item || !canEditAngebot(item, scopeFromState(state))) {
        toast('Löschen nicht erlaubt.');
        return;
    }
    if (state.localDemoOnly || !state.ctx) {
        state.items = state.items.filter((a) => a.itemId !== itemId);
        if (state.view === 'detail') closeDetail();
        else paint();
        toast('Gelöscht (Demo).');
        return;
    }
    await deleteAngebotItem(state.ctx, itemId);
    await refreshData();
    if (state.view === 'detail') closeDetail();
    else paint();
    toast('Gelöscht.');
}

async function decide(itemId, status, reason) {
    const item = state.items.find((a) => a.itemId === itemId);
    if (!item) return;
    const stayDetail = state.view === 'detail';
    const patch = {
        ...item,
        status,
        ablehnungsGrund: status === 'abgelehnt' ? reason || '' : '',
        freigegebenVon: status === 'freigegeben' ? state.accountEmail || '' : item.freigegebenVon,
        freigegebenAm: status === 'freigegeben' ? new Date().toISOString() : item.freigegebenAm,
        buchungAb: item.buchungAb || (state.aktion && state.aktion.buchungAbDefault) || ''
    };
    if (state.localDemoOnly || !state.ctx) {
        state.items = state.items.map((a) => (a.itemId === itemId ? patch : a));
        if (stayDetail) {
            state.detailId = itemId;
            if (canEditAngebot(patch, scopeFromState(state))) {
                state.editingItemId = itemId;
                state.form = formFromItem(patch);
            } else {
                state.editingItemId = null;
            }
        }
        paint();
        toast(status === 'freigegeben' ? 'Freigegeben (Demo).' : 'Abgelehnt (Demo).');
        return;
    }
    await updateAngebotItem(state.ctx, itemId, patch);
    await refreshData();
    if (stayDetail) {
        const again = state.items.find((a) => a.itemId === itemId);
        state.view = 'detail';
        state.detailId = itemId;
        if (again && canEditAngebot(again, scopeFromState(state))) {
            state.editingItemId = again.itemId;
            state.form = formFromItem(again);
        } else {
            state.editingItemId = null;
        }
        paint();
    }
    toast(status === 'freigegeben' ? 'Freigegeben.' : 'Abgelehnt.');
}

async function saveBuchungAb(itemId, buchungAb) {
    const item = state.items.find((a) => a.itemId === itemId);
    if (!item) return;
    const patch = { ...item, buchungAb };
    if (state.localDemoOnly || !state.ctx) {
        state.items = state.items.map((a) => (a.itemId === itemId ? patch : a));
        paint();
        toast('Buchungsstart gespeichert (Demo).');
        return;
    }
    await updateAngebotItem(state.ctx, itemId, patch);
    await refreshData();
    toast('Buchungsstart gespeichert.');
}

async function saveAktionDefault() {
    if (!state.aktion) return;
    const el = document.getElementById('pwAktionBuchungDefault');
    const buchungAbDefault = el ? String(el.value || '').trim() : '';
    const patch = { ...state.aktion, buchungAbDefault };
    if (state.localDemoOnly || !state.ctx) {
        state.aktion = patch;
        state.aktionen = state.aktionen.map((a) => (a.itemId === patch.itemId ? patch : a));
        paint();
        toast('Aktion gespeichert (Demo).');
        return;
    }
    await updateAktionItem(state.ctx, state.aktion.itemId, patch);
    await refreshData();
    toast('Aktion gespeichert.');
}

async function runEnsureBusiness() {
    if (!state.aktion) {
        toast('Keine aktive Aktion.');
        return;
    }
    state.bookingsBusy = true;
    state.bookingsLog = '';
    paint();
    const log = (m) => {
        appendBookingsLog(m);
        const el = document.getElementById('pwBookingsLog');
        if (el) el.textContent = state.bookingsLog;
    };
    try {
        if (state.localDemoOnly || !state.ctx) {
            const fakeId = 'demo-biz-' + (state.aktion.aktionId || 'pw');
            state.aktion = {
                ...state.aktion,
                bookingsBusinessId: fakeId,
                bookingsBusinessName: state.aktion.title || 'Demo Business'
            };
            state.aktionen = state.aktionen.map((a) =>
                a.itemId === state.aktion.itemId ? state.aktion : a
            );
            log('Demo: Business gebunden als ' + fakeId);
            toast('Demo-Business gesetzt.');
            return;
        }
        const biz = await ensureBookingBusiness(state.aktion, {}, log);
        const patch = {
            ...state.aktion,
            bookingsBusinessId: biz.id,
            bookingsBusinessName: biz.displayName || state.aktion.title
        };
        await updateAktionItem(state.ctx, state.aktion.itemId, patch);
        state.aktion = patch;
        state.aktionen = state.aktionen.map((a) => (a.itemId === patch.itemId ? patch : a));
        log('Business-ID in SharePoint gespeichert.');
        toast('Bookings-Business bereit.');
    } finally {
        state.bookingsBusy = false;
        paint();
    }
}

async function runSyncAngebote(itemIds) {
    const ids = (itemIds || []).filter(Boolean);
    if (!ids.length) {
        toast('Keine Angebote ausgewählt.');
        return;
    }
    if (!state.aktion) {
        toast('Keine aktive Aktion.');
        return;
    }
    state.bookingsBusy = true;
    state.bookingsLog = '';
    paint();
    const log = (m) => {
        appendBookingsLog(m);
        const el = document.getElementById('pwBookingsLog');
        if (el) el.textContent = state.bookingsLog;
    };

    try {
        let businessId = String(state.aktion.bookingsBusinessId || '').trim();
        if (!businessId) {
            if (state.localDemoOnly || !state.ctx) {
                businessId = 'demo-biz-' + (state.aktion.aktionId || 'pw');
                state.aktion = {
                    ...state.aktion,
                    bookingsBusinessId: businessId,
                    bookingsBusinessName: state.aktion.title
                };
            } else {
                const biz = await ensureBookingBusiness(state.aktion, {}, log);
                businessId = biz.id;
                const patch = {
                    ...state.aktion,
                    bookingsBusinessId: biz.id,
                    bookingsBusinessName: biz.displayName
                };
                await updateAktionItem(state.ctx, state.aktion.itemId, patch);
                state.aktion = patch;
            }
        }

        let ok = 0;
        let fail = 0;
        for (let i = 0; i < ids.length; i++) {
            const item = state.items.find((a) => a.itemId === ids[i]);
            if (!item) continue;
            log('— Sync: ' + item.title);
            try {
                if (item.status !== 'freigegeben') {
                    log('  übersprungen (nicht freigegeben)');
                    continue;
                }
                let result;
                if (state.localDemoOnly || !state.ctx) {
                    result = {
                        serviceId: 'demo-svc-' + (item.angebotId || newEntityId('svc')),
                        bookingUrl: 'https://outlook.office.com/bookings/demo',
                        created: true
                    };
                    log('  Demo-Service ' + result.serviceId);
                } else {
                    result = await syncAngebotToBookings(businessId, item, state.aktion, log);
                }
                const patched = {
                    ...item,
                    bookingsServiceId: result.serviceId,
                    bookingsBookingUrl: result.bookingUrl || item.bookingsBookingUrl,
                    syncStatus: 'ok',
                    syncFehler: '',
                    syncAm: new Date().toISOString()
                };
                if (state.localDemoOnly || !state.ctx) {
                    state.items = state.items.map((a) => (a.itemId === item.itemId ? patched : a));
                } else {
                    await updateAngebotItem(state.ctx, item.itemId, patched);
                    state.items = state.items.map((a) => (a.itemId === item.itemId ? patched : a));
                }
                ok++;
            } catch (e) {
                fail++;
                const msg = e && e.message ? e.message : String(e);
                log('  FEHLER: ' + msg);
                const patched = {
                    ...item,
                    syncStatus: 'fehler',
                    syncFehler: msg,
                    syncAm: new Date().toISOString()
                };
                if (state.localDemoOnly || !state.ctx) {
                    state.items = state.items.map((a) => (a.itemId === item.itemId ? patched : a));
                } else {
                    try {
                        await updateAngebotItem(state.ctx, item.itemId, patched);
                    } catch {
                        /* ignore secondary */
                    }
                    state.items = state.items.map((a) => (a.itemId === item.itemId ? patched : a));
                }
            }
        }
        toast('Sync fertig: ' + ok + ' ok, ' + fail + ' Fehler.');
    } finally {
        state.bookingsBusy = false;
        paint();
    }
}

async function runLoadAttendees(opts) {
    const stayOnDetail = !!(opts && opts.stayOnDetail) && state.view === 'detail';
    if (!state.aktion) {
        toast('Keine aktive Aktion.');
        return;
    }
    const businessId = String(state.aktion.bookingsBusinessId || '').trim();
    state.bookingsBusy = true;
    paint();
    const log = (m) => appendBookingsLog(m);
    try {
        if (state.localDemoOnly || !businessId || businessId.indexOf('demo-') === 0) {
            // Demo-TN aus freigegebenen Angeboten simulieren
            const rows = [];
            const occupancy = {};
            state.items
                .filter((a) => a.status === 'freigegeben')
                .forEach((a, idx) => {
                    const sid = a.bookingsServiceId || 'demo-svc-' + a.angebotId;
                    occupancy[sid] = {
                        filled: Math.min(a.kapazitaet, 3 + (idx % 5)),
                        max: a.kapazitaet,
                        angebotId: a.angebotId,
                        title: a.title
                    };
                    for (let n = 0; n < occupancy[sid].filled; n++) {
                        rows.push({
                            appointmentId: 'demo-ap-' + idx + '-' + n,
                            serviceId: sid,
                            angebotId: a.angebotId,
                            angebotTitle: a.title,
                            datum: a.datum,
                            name: 'Schüler ' + (idx + 1) + '-' + (n + 1),
                            email: 'schueler.' + idx + n + '@kurtrocks.onmicrosoft.com',
                            phone: '',
                            klasse: ['1AK', '2AK', '3AK'][n % 3]
                        });
                    }
                });
            state.attendeeRows = rows;
            state.occupancy = occupancy;
            state.attendeeLoadedAt = new Date().toLocaleString('de-AT');
            log('Demo: ' + rows.length + ' Teilnehmerzeilen simuliert.');
            toast('Demo-Teilnehmer geladen.');
            if (!stayOnDetail) state.view = 'teilnehmer';
            return;
        }
        const data = await loadBookingsAttendees(
            businessId,
            state.aktion,
            state.items,
            (state.stammdaten && state.stammdaten.students) || [],
            log
        );
        state.attendeeRows = data.rows;
        state.occupancy = data.occupancy;
        state.attendeeLoadedAt = new Date().toLocaleString('de-AT');
        toast(data.rows.length + ' Teilnehmer geladen.');
        if (!stayOnDetail) state.view = 'teilnehmer';
    } finally {
        state.bookingsBusy = false;
        paint();
    }
}

function bootFromQuery() {
    try {
        const q = new URLSearchParams(window.location.search);
        const view = q.get('view');
        const role = q.get('role');
        if (role === 'admin' || role === 'lehrer' || role === 'schueler') {
            state.role = role;
            persistRole(role);
        }
        const allowed = viewsForRole(state.role).map((v) => v.id);
        if (view && allowed.includes(view)) state.view = view;
    } catch {
        /* ignore */
    }
}

async function init() {
    root = document.getElementById('pwApp');
    if (!root) return;
    bootFromQuery();
    state.siteUrl = loadSavedSiteUrl();
    await syncAccount();
    state.form = emptyForm(prefillTeacher());
    paint();
    // Gespeicherte Site: Daten laden, ohne die URL allen Rollen zu zeigen
    if (state.siteUrl && !state.localDemoOnly) {
        refreshData().catch((e) => {
            state.error = e && e.message ? e.message : String(e);
            state.loading = false;
            paint();
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        init().catch((e) => console.error(e));
    });
} else {
    init().catch((e) => console.error(e));
}
