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
import {
    validateAngebot,
    weekdayLabelDeFromIso,
    resolvePwOwnership,
    assertPwOwnership,
    isPwAccessDeniedError,
    friendlyPwSpoError
} from './projektwochen-logic.js';
import {
    isPwLocalOnlyItem,
    mergePwSpoSyncMeta,
    shouldSkipPwSpoPush
} from './projektwochen-sync.js';
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
    friendlyBookingsError,
    summarizePwSyncBatch
} from './projektwochen-bookings-logic.js';
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
import {
    applyPwRoleFromEntra,
    isPwDemoRoleUiEnabled,
    readDemoRoleOverrideFromUrl
} from './projektwochen-entra-role.js';
import { hydrateStammdatenFromLiveSite } from '../../shared/school-stammdaten-spo-lists.js';
import {
    toolSiteUrlFromLicenseResult,
    applyToolSiteUrlLocally,
    bootstrapSchoolSiteFromLicense
} from '../../shared/tenant-school-site-bootstrap.js';

const STAMM_READ_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
];
const STAMM_RW_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

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
    import('../../shared/frontend-planner-chrome-policy.js')
        .then((m) => {
            if (m && typeof m.refreshFrontendPlannerChromeAudience === 'function') {
                return m.refreshFrontendPlannerChromeAudience();
            }
        })
        .catch(() => {});
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
            const demoUi = isPwDemoRoleUiEnabled(
                state.entraGroupsConfigured,
                state.demoRoleOverride
            );
            const raw = btn.getAttribute('data-pw-role') || 'lehrer';
            if (!demoUi && !(state.planerRoles || []).includes(raw)) return;
            const role = raw === 'admin' ? 'admin' : raw === 'schueler' ? 'schueler' : 'lehrer';
            state.role = role;
            if (demoUi) state.demoRoleOverride = true;
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
    state.spoSync = mergePwSpoSyncMeta(state.spoSync, {
        ok: true,
        source: 'demo',
        connected: false,
        error: '',
        siteUrl: String(state.siteUrl || ''),
        lists: '',
        lastPullAt: ''
    });
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

function markPwPullLive(ctx) {
    state.spoSync = mergePwSpoSyncMeta(state.spoSync, {
        ok: true,
        source: 'live',
        connected: true,
        error: '',
        siteUrl: String((ctx && ctx.webUrl) || state.siteUrl || ''),
        lists: 'PW-Aktionen, PW-Angebote',
        lastPullAt: new Date().toISOString()
    });
}

function markPwPushLive() {
    state.spoSync = mergePwSpoSyncMeta(state.spoSync, {
        ok: true,
        source: 'live',
        connected: true,
        error: '',
        lastPushAt: new Date().toISOString()
    });
}

function markPwSyncSoftFail(msg, opts) {
    const keepDemo = !!(opts && opts.keepDemo);
    state.spoSync = mergePwSpoSyncMeta(state.spoSync, {
        ok: false,
        source: keepDemo || state.localDemoOnly ? 'demo' : 'local',
        connected: false,
        error: friendlyPwSpoError(msg),
        siteUrl: String(state.siteUrl || '')
    });
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
        // Roh-403 nicht in Banner – Demo bleibt nutzbar
        state.error = '';
        paint();
        toast('Demo bleibt lokal (SharePoint-Schreiben nicht möglich). Für die Vorführung reicht das.');
        console.warn('[pw-demo] SharePoint seed failed', e);
    }
}

/**
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
 * Lehrer/Klassen live hydratisieren; bei Fehler lokalen Stand behalten.
 * @param {string} siteId
 */
async function hydrateLiveStammdaten(siteId) {
    const local = loadStammdaten();
    if (state.localDemoOnly || !siteId) {
        state.stammdaten = state.stammdaten && (state.stammdaten.teachers || []).length ? state.stammdaten : local;
        state.stammdatenLive = {
            ok: false,
            error: siteId ? '' : 'Keine Site-URL',
            classCount: (state.stammdaten.classes || []).length,
            teacherCount: (state.stammdaten.teachers || []).length,
            missing: [],
            source: 'local',
            siteUrl: String(state.siteUrl || '')
        };
        return;
    }
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api || typeof api.getGraphToken !== 'function') {
        state.stammdaten = local;
        state.stammdatenLive = {
            ok: false,
            error: '',
            classCount: (local.classes || []).length,
            teacherCount: (local.teachers || []).length,
            missing: [],
            source: 'local',
            siteUrl: String(state.siteUrl || '')
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
            local
        });
        const liveClasses = (result.live && result.live.classes) || [];
        const liveTeachers = (result.live && result.live.teachers) || [];
        const missing = (result.live && result.live.missing) || [];
        const usedLive = !!(liveClasses.length || liveTeachers.length);
        // Nie leeren Stand übernehmen, wenn lokal schon Daten da waren
        if (
            (result.stammdaten &&
                ((result.stammdaten.classes || []).length || (result.stammdaten.teachers || []).length)) ||
            (local.classes || []).length ||
            (local.teachers || []).length
        ) {
            state.stammdaten = result.stammdaten || local;
            if (!(state.stammdaten.classes || []).length && (local.classes || []).length) {
                state.stammdaten = { ...state.stammdaten, classes: local.classes };
            }
            if (!(state.stammdaten.teachers || []).length && (local.teachers || []).length) {
                state.stammdaten = { ...state.stammdaten, teachers: local.teachers };
            }
        } else {
            state.stammdaten = local;
        }
        const liveSource = (result.live && result.live.source) || '';
        let err = '';
        if (!usedLive && result.error) err = friendlyPwSpoError(result.error);
        state.stammdatenLive = {
            ok: usedLive,
            error: err,
            classCount: (state.stammdaten.classes || []).length,
            teacherCount: (state.stammdaten.teachers || []).length,
            missing,
            source: usedLive
                ? liveSource === 'spo-rest'
                    ? 'sharepoint-spo'
                    : 'sharepoint'
                : 'local-fallback',
            siteUrl: String(state.siteUrl || '')
        };
    } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        state.stammdaten = local;
        state.stammdatenLive = {
            ok: (local.classes || []).length > 0,
            error: friendlyPwSpoError(msg),
            classCount: (local.classes || []).length,
            teacherCount: (local.teachers || []).length,
            missing: [],
            source: 'local-fallback',
            siteUrl: String(state.siteUrl || '')
        };
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
    if (!state.stammdaten || (!(state.stammdaten.teachers || []).length && !(state.stammdaten.classes || []).length)) {
        state.stammdaten = loadStammdaten();
    }
    state.teacherMatch = matchTeacherByEmail(state.accountEmail, state.stammdaten.teachers);
    await applyPwRoleFromEntra(state, {
        demoRoleOverride: state.demoRoleOverride,
        preferredActiveRole: state.role
    });
    const allowed = viewsForRole(state.role).map((v) => v.id);
    if (!allowed.includes(state.view)) state.view = 'dashboard';
}

/**
 * Lehrer-Browser oft ohne lokale IT-Einrichtung: Site-URL aus Lizenz.
 * @returns {Promise<string>}
 */
async function ensurePwSiteFromLicense() {
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
            token = await api.acquireLicenseToken();
        } else if (typeof window.ms365AuthAcquireIdToken === 'function') {
            token = await window.ms365AuthAcquireIdToken();
        }
        if (!token) return '';
        const result = await api.fetchLicenseMe(token);
        if (!result || result.allowed !== true) return '';
        const url = toolSiteUrlFromLicenseResult(result);
        if (!url) {
            state.roleHint =
                'Keine Schul-Site in der Lizenz. IT: unter Admin → Schule die SharePoint-URL setzen.';
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

/** @returns {Promise<string>} */
async function ensurePwSiteUrl() {
    let url = String(state.siteUrl || '').trim() || loadSavedSiteUrl();
    if (url) {
        state.siteUrl = url;
        return url;
    }
    url = await ensurePwSiteFromLicense();
    if (url) return url;
    url = loadSavedSiteUrl();
    if (url) {
        state.siteUrl = url;
        return url;
    }
    return '';
}

async function refreshData() {
    const hadLocalDemo = !!(state.localDemoOnly && (state.items || []).length);
    const keepItems = hadLocalDemo ? state.items.slice() : null;
    const keepAktionen = hadLocalDemo ? (state.aktionen || []).slice() : null;
    const keepAktion = hadLocalDemo ? state.aktion : null;

    state.loading = true;
    state.error = '';
    paint();
    try {
        await syncAccount();
        const siteUrl = await ensurePwSiteUrl();
        if (!siteUrl) {
            throw new Error('Keine SharePoint-Website hinterlegt.');
        }
        state.siteUrl = siteUrl;
        const ctx = await resolvePwContext(siteUrl);
        state.ctx = ctx;
        await hydrateLiveStammdaten(ctx.siteId);
        state.teacherMatch = matchTeacherByEmail(state.accountEmail, state.stammdaten.teachers);
        const data = await loadAllPwData(ctx);
        state.aktionen = data.aktionen;
        state.aktion = pickActiveAktion(data.aktionen, data.angebote);
        state.items = data.angebote;
        state.localDemoOnly = false;
        state.bootstrapped = true;
        state.loading = false;
        markPwPullLive(ctx);
        if (state.aktion && state.aktion.startdatum) {
            const d = new Date(state.aktion.startdatum + 'T12:00:00');
            if (!Number.isNaN(d.getTime())) {
                state.calYear = d.getFullYear();
                state.calMonth = d.getMonth();
            }
        }
        if (!state.aktion) {
            state.roleHint = 'Keine PW-Aktion gefunden – bitte System/Setup prüfen.';
        }
        state.form = emptyForm(prefillTeacher());
        paint();
        toast('Projektwochen geladen (' + state.items.length + ' Angebote).');
    } catch (e) {
        state.loading = false;
        const msg = e && e.message ? e.message : String(e);
        if (hadLocalDemo && keepItems) {
            state.items = keepItems;
            state.aktionen = keepAktionen || [];
            state.aktion = keepAktion;
            state.localDemoOnly = true;
            state.bootstrapped = true;
            state.error = '';
            markPwSyncSoftFail(msg, { keepDemo: true });
            console.warn('[pw] refresh skipped – local demo kept', e);
            toast('SharePoint nicht erreichbar – Demo bleibt lokal.');
        } else if (isPwAccessDeniedError(msg)) {
            state.error = '';
            markPwSyncSoftFail(msg, { keepDemo: !!state.localDemoOnly });
            toast(friendlyPwSpoError(msg));
            console.warn('[pw] soft 403', e);
        } else {
            markPwSyncSoftFail(msg, { keepDemo: !!state.localDemoOnly });
            state.error = friendlyPwSpoError(msg);
        }
        paint();
    }
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

    const ownership = resolvePwOwnership(
        {
            ...form,
            beantragtVon: state.accountEmail || form.lehrerEmail || ''
        },
        {
            accountEmail: state.accountEmail,
            teachers: (state.stammdaten && state.stammdaten.teachers) || []
        }
    );
    try {
        assertPwOwnership(ownership);
    } catch (e) {
        toast(e && e.message ? e.message : String(e));
        return;
    }
    const payload = {
        ...form,
        ...ownership,
        angebotId: form.angebotId || newEntityId('ang'),
        aktionId: (state.aktion && state.aktion.aktionId) || '',
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

    if (shouldSkipPwSpoPush(state)) {
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

    // local-* nie per PATCH – bei Context als Create nach SPO promoten
    if (state.editingItemId && !isPwLocalOnlyItem(state.editingItemId)) {
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
    markPwPushLive();
    const keepId = state.editingItemId || state.detailId;
    const keepAngebotId = payload.angebotId;
    await refreshData();
    if (stayDetail && (keepId || keepAngebotId)) {
        const again = state.items.find(
            (a) =>
                (keepId && !isPwLocalOnlyItem(keepId) && a.itemId === keepId) ||
                (keepAngebotId && a.angebotId === keepAngebotId)
        );
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
    if (shouldSkipPwSpoPush(state) || isPwLocalOnlyItem(itemId)) {
        state.items = state.items.filter((a) => a.itemId !== itemId);
        if (state.view === 'detail') closeDetail();
        else paint();
        toast(shouldSkipPwSpoPush(state) ? 'Gelöscht (Demo).' : 'Lokal gelöscht.');
        return;
    }
    await deleteAngebotItem(state.ctx, itemId);
    markPwPushLive();
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
    if (shouldSkipPwSpoPush(state) || isPwLocalOnlyItem(itemId)) {
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
    markPwPushLive();
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
    if (shouldSkipPwSpoPush(state) || isPwLocalOnlyItem(itemId)) {
        state.items = state.items.map((a) => (a.itemId === itemId ? patch : a));
        paint();
        toast('Buchungsstart gespeichert (Demo).');
        return;
    }
    await updateAngebotItem(state.ctx, itemId, patch);
    markPwPushLive();
    await refreshData();
    toast('Buchungsstart gespeichert.');
}

async function saveAktionDefault() {
    if (!state.aktion) return;
    const el = document.getElementById('pwAktionBuchungDefault');
    const buchungAbDefault = el ? String(el.value || '').trim() : '';
    const patch = { ...state.aktion, buchungAbDefault };
    if (shouldSkipPwSpoPush(state) || isPwLocalOnlyItem(state.aktion.itemId)) {
        state.aktion = patch;
        state.aktionen = state.aktionen.map((a) => (a.itemId === patch.itemId ? patch : a));
        paint();
        toast('Aktion gespeichert (Demo).');
        return;
    }
    await updateAktionItem(state.ctx, state.aktion.itemId, patch);
    markPwPushLive();
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
        let skipped = 0;
        /** @type {string[]} */
        const errors = [];
        for (let i = 0; i < ids.length; i++) {
            const item = state.items.find((a) => a.itemId === ids[i]);
            if (!item) continue;
            log('— Sync: ' + item.title);
            try {
                if (item.status !== 'freigegeben') {
                    skipped++;
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
                if (shouldSkipPwSpoPush(state) || isPwLocalOnlyItem(item.itemId)) {
                    state.items = state.items.map((a) => (a.itemId === item.itemId ? patched : a));
                } else {
                    await updateAngebotItem(state.ctx, item.itemId, patched);
                    markPwPushLive();
                    state.items = state.items.map((a) => (a.itemId === item.itemId ? patched : a));
                }
                ok++;
            } catch (e) {
                fail++;
                const raw = e && e.message ? e.message : String(e);
                const msg = friendlyBookingsError(raw);
                errors.push(msg);
                log('  FEHLER: ' + msg);
                const patched = {
                    ...item,
                    syncStatus: 'fehler',
                    syncFehler: msg,
                    syncAm: new Date().toISOString()
                };
                if (shouldSkipPwSpoPush(state) || isPwLocalOnlyItem(item.itemId)) {
                    state.items = state.items.map((a) => (a.itemId === item.itemId ? patched : a));
                } else {
                    try {
                        await updateAngebotItem(state.ctx, item.itemId, patched);
                        markPwPushLive();
                    } catch {
                        /* ignore secondary */
                    }
                    state.items = state.items.map((a) => (a.itemId === item.itemId ? patched : a));
                }
            }
        }
        const summary = summarizePwSyncBatch({ ok, fail, skipped, errors });
        state.bookingsSyncSummary = summary;
        toast(summary.headline + (summary.detail ? ' – ' + summary.detail : ''));
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
        state.demoRoleOverride = readDemoRoleOverrideFromUrl(false);
        // ?role= nur in IT-Vorschau
        if (
            state.demoRoleOverride &&
            (role === 'admin' || role === 'lehrer' || role === 'schueler')
        ) {
            state.role = role;
            persistRole(role);
        }
        const allowed = viewsForRole(state.role).map((v) => v.id);
        if (view && allowed.includes(view)) state.view = view;
        // Demo-Daten nur per ?demo=1 (Opt-in)
        if (q.get('demo') === '1') {
            state._bootDemo = true;
        }
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
    if (state._bootDemo) {
        delete state._bootDemo;
        applyLocalDemo(getDemoSeedPackage());
        toast('Demo lokal geladen (?demo=1).');
        return;
    }
    paint();
    const url = await ensurePwSiteUrl();
    if (url && !state.localDemoOnly) {
        await refreshData();
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        init().catch((e) => console.error(e));
    });
} else {
    init().catch((e) => console.error(e));
}
