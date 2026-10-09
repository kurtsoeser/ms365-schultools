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
    canEditItem
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
import { validateAktivitaet, DEFAULT_RULES } from './schulaktivitaeten-planer-logic.js';
import {
    renderApp,
    readFormFromDom,
    readFiltersFromDom,
    readRulesForm
} from './schulaktivitaeten-planer-ui.js';
import { buildIcs, downloadIcs } from './schulaktivitaeten-planer-export.js';
import {
    getDemoSeedPackage,
    buildLocalDemoState,
    DEMO_SITE_DEFAULT
} from './schulaktivitaeten-planer-demo-data.js';
import {
    parseDemoImportJson,
    applyDemoStammdatenLocal,
    seedDemoSchulaktivitaeten,
    resetDemoSchulaktivitaeten
} from './schulaktivitaeten-planer-demo-seed.js';

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

function currentAccount() {
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
        if (acc) return { email: acc.username || acc.mail || '', name: acc.name || '' };
    } catch {
        /* ignore */
    }
    return { email: '', name: '' };
}

function syncAccount() {
    const a = currentAccount();
    state.accountEmail = String(a.email || a.username || '')
        .trim()
        .toLowerCase();
    state.accountName = String(a.name || '').trim();
    state.stammdaten = loadStammdaten();
    const match = matchTeacherByEmail(state.stammdaten.teachers, state.accountEmail);
    if (match && state.view === 'antrag' && !state.form.lehrerCode) {
        state.form.lehrerCode = match.code || '';
        state.form.lehrerEmail = state.accountEmail;
    }
}

async function refreshData() {
    state.loading = true;
    state.error = '';
    paint();
    try {
        syncAccount();
        const ctx = await resolveAktContext(state.siteUrl);
        state.ctx = ctx;
        const data = await loadAllAktData(ctx);
        state.items = data.items || [];
        state.rules = data.rules || { ...DEFAULT_RULES };
        state.localDemoOnly = false;
        state.loading = false;
        paint();
    } catch (e) {
        state.loading = false;
        state.error = e && e.message ? e.message : String(e);
        paint();
    }
}

function bindStatic() {
    root.querySelectorAll('[data-akt-view]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-akt-view') || 'dashboard';
            state.detailId = null;
            if (state.view === 'antrag' && !state.editingItemId) state.form = emptyForm(prefillTeacher());
            paint();
        });
    });

    root.querySelectorAll('[data-akt-view-jump]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-akt-view-jump') || 'dashboard';
            paint();
        });
    });

    root.querySelectorAll('[data-akt-role]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const role = btn.getAttribute('data-akt-role') === 'admin' ? 'admin' : 'lehrer';
            state.role = role;
            persistRole(role);
            const allowed = viewsForRole(role).map((v) => v.id);
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

    ['aktFilterKlasse', 'aktFilterTyp', 'aktFilterStatus', 'aktFilterLehrer'].forEach((id) => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('change', () => {
                state.filters = readFiltersFromDom();
                paint();
            });
        }
    });
    const reset = document.getElementById('aktFilterReset');
    if (reset) {
        reset.addEventListener('click', () => {
            state.filters = { klasse: '', typ: '', status: '', lehrer: '' };
            paint();
        });
    }

    const form = document.getElementById('aktForm');
    if (form) {
        form.addEventListener('submit', (ev) => {
            ev.preventDefault();
            saveForm().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
        form.addEventListener('input', () => {
            state.form = { ...state.form, ...readFormFromDom() };
        });
        form.addEventListener('change', () => {
            state.form = { ...state.form, ...readFormFromDom() };
            paint();
        });
    }
    const formReset = document.getElementById('aktFormReset');
    if (formReset) {
        formReset.addEventListener('click', () => {
            state.editingItemId = null;
            state.form = emptyForm(prefillTeacher());
            paint();
        });
    }

    const rulesForm = document.getElementById('aktRulesForm');
    if (rulesForm) {
        rulesForm.addEventListener('submit', (ev) => {
            ev.preventDefault();
            saveRules().catch((e) => toast(e && e.message ? e.message : String(e)));
        });
    }

    root.querySelectorAll('[data-akt-detail]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.detailId = btn.getAttribute('data-akt-detail');
            paint();
        });
    });
    const close = document.getElementById('aktDetailClose');
    if (close) {
        close.addEventListener('click', () => {
            state.detailId = null;
            paint();
        });
    }

    root.querySelectorAll('[data-akt-edit]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-akt-edit');
            const it = state.items.find((x) => x.itemId === id);
            if (!it || !canEditItem(it, state)) return;
            state.editingItemId = it.itemId;
            state.form = emptyForm(it);
            state.detailId = null;
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

    root.querySelectorAll('[data-akt-approve]').forEach((btn) => {
        btn.addEventListener('click', () => {
            decide(btn.getAttribute('data-akt-approve'), 'genehmigt').catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });
    root.querySelectorAll('[data-akt-reject]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const grund = window.prompt('Ablehnungsgrund (optional):', '') || '';
            decide(btn.getAttribute('data-akt-reject'), 'abgelehnt', grund).catch((e) =>
                toast(e && e.message ? e.message : String(e))
            );
        });
    });

    const prev = document.getElementById('aktCalPrev');
    const next = document.getElementById('aktCalNext');
    if (prev) {
        prev.addEventListener('click', () => {
            state.calMonth -= 1;
            if (state.calMonth < 1) {
                state.calMonth = 12;
                state.calYear -= 1;
            }
            paint();
        });
    }
    if (next) {
        next.addEventListener('click', () => {
            state.calMonth += 1;
            if (state.calMonth > 12) {
                state.calMonth = 1;
                state.calYear += 1;
            }
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

    const wireDemo = (id) => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', () => runDemoImport(getDemoSeedPackage()).catch((e) => toast(String((e && e.message) || e))));
    };
    wireDemo('aktBtnDemo');
    wireDemo('aktBtnDemoPanel');

    const wireReset = (id) => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', () => runDemoReset().catch((e) => toast(String((e && e.message) || e))));
    };
    wireReset('aktBtnDemoReset');
    wireReset('aktBtnDemoResetPanel');

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
    const match = matchTeacherByEmail(state.stammdaten.teachers, state.accountEmail);
    return {
        lehrerCode: match ? match.code || '' : '',
        lehrerEmail: state.accountEmail || ''
    };
}

async function saveForm() {
    const draft = { ...state.form, ...readFormFromDom() };
    if (!draft.lehrerEmail) draft.lehrerEmail = state.accountEmail;
    if (!draft.beantragtVon) draft.beantragtVon = state.accountEmail;
    const teacher = (state.stammdaten.teachers || []).find((t) => t.code === draft.lehrerCode);
    if (teacher && teacher.email) draft.lehrerEmail = String(teacher.email).toLowerCase();

    const rules = state.rules || DEFAULT_RULES;
    const check = validateAktivitaet({
        draft: { ...draft, aktivitaetId: state.form.aktivitaetId || '' },
        existing: state.items,
        rules: {
            minVorlaufTage: rules.minVorlaufTage,
            maxGleichzeitigProKlasse: rules.maxGleichzeitigProKlasse
        }
    });
    if (!check.ok) {
        state.form = draft;
        paint();
        throw new Error(check.errors[0] || 'Validierung fehlgeschlagen.');
    }

    if (state.localDemoOnly || !state.ctx) {
        if (state.editingItemId) {
            state.items = state.items.map((it) =>
                it.itemId === state.editingItemId
                    ? {
                          ...it,
                          ...draft,
                          aktivitaetId: it.aktivitaetId || draft.aktivitaetId || newEntityId('akt'),
                          status: it.status || 'beantragt'
                      }
                    : it
            );
            toast('Antrag aktualisiert (Demo).');
        } else {
            const id = newEntityId('akt');
            state.items = [
                {
                    itemId: 'local-' + id,
                    aktivitaetId: id,
                    ...draft,
                    status: 'beantragt',
                    beantragtVon: state.accountEmail,
                    _localOnly: true
                },
                ...state.items
            ];
            toast('Antrag gestellt (Demo).');
        }
        state.editingItemId = null;
        state.form = emptyForm(prefillTeacher());
        state.view = 'liste';
        state.localDemoOnly = true;
        paint();
        return;
    }

    state.loading = true;
    paint();
    try {
        if (state.editingItemId) {
            const prev = state.items.find((x) => x.itemId === state.editingItemId);
            await updateAktivitaetItem(state.ctx, state.editingItemId, {
                ...prev,
                ...draft,
                aktivitaetId: (prev && prev.aktivitaetId) || draft.aktivitaetId || newEntityId('akt'),
                status: (prev && prev.status) || 'beantragt'
            });
            toast('Antrag aktualisiert.');
        } else {
            await createAktivitaetItem(state.ctx, {
                ...draft,
                aktivitaetId: newEntityId('akt'),
                status: 'beantragt',
                beantragtVon: state.accountEmail
            });
            toast('Antrag gestellt.');
        }
        state.editingItemId = null;
        state.form = emptyForm(prefillTeacher());
        state.view = 'liste';
        await refreshData();
    } catch (e) {
        state.loading = false;
        throw e;
    }
}

async function decide(itemId, status, ablehnungsGrund) {
    const it = state.items.find((x) => x.itemId === itemId);
    if (!it) throw new Error('Eintrag nicht gefunden.');
    const patch = {
        ...it,
        status,
        ablehnungsGrund: status === 'abgelehnt' ? ablehnungsGrund || '' : '',
        genehmigtVon: state.accountEmail,
        genehmigtAm: new Date().toISOString()
    };
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

    if (state.localDemoOnly || !state.ctx) {
        state.items = state.items.map((row) => (row.itemId === itemId ? { ...row, ...patch } : row));
        toast(status === 'genehmigt' ? 'Genehmigt (Demo).' : 'Abgelehnt (Demo).');
        state.detailId = null;
        state.localDemoOnly = true;
        paint();
        return;
    }

    await updateAktivitaetItem(state.ctx, itemId, patch);
    toast(status === 'genehmigt' ? 'Genehmigt.' : 'Abgelehnt.');
    state.detailId = null;
    await refreshData();
}

async function deleteItem(itemId) {
    const it = state.items.find((x) => x.itemId === itemId);
    if (!it || !canEditItem(it, state)) throw new Error('Löschen nicht erlaubt.');
    if (!window.confirm('Antrag wirklich löschen?')) return;

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
            itemId: (state.rules && state.rules.itemId) || 'local-rw'
        };
        state.localDemoOnly = true;
        toast('Regelwerk gespeichert (Demo).');
        paint();
        return;
    }

    await updateRegelwerkItem(state.ctx, state.rules.itemId, patch);
    toast('Regelwerk gespeichert.');
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
    state.view = 'dashboard';
    state.role = 'admin';
    persistRole('admin');
    if (!state.siteUrl && local.siteDefault) {
        state.siteUrl = local.siteDefault;
        persistSiteUrl(state.siteUrl);
    }
    // Kalender auf Schuljahrbeginn
    state.calYear = 2026;
    state.calMonth = 10;
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

function boot() {
    root = document.getElementById('aktApp');
    if (!root) return;
    state.role = resolveRole();
    state.siteUrl = loadSavedSiteUrl() || state.siteUrl;
    state.stammdaten = loadStammdaten();
    state.form = emptyForm(prefillTeacher());
    paint();
    if (state.siteUrl) {
        refreshData().catch(() => {
            /* error already in state */
        });
    }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
