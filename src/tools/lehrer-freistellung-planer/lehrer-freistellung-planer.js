/**
 * Freistellungen Lehrkräfte – Antrag, Direktions-Freigabe, Kalender, iCal.
 */
import {
    createInitialState,
    loadStammdaten,
    matchTeacherByEmail,
    resolveRole,
    persistRole,
    persistSiteUrl,
    persistSetupCfg,
    scopeFromState,
    emptyForm,
    prefillLehrerFormIdentity,
    loadDemoItems,
    saveDemoItems,
    viewsForRole
} from './lfr-state.js';
import { validateAntrag, normalizeStatus, toIsoDateOnly } from './lfr-logic.js';
import { newAntragId } from './lfr-schema.js';
import {
    resolveLfrContext,
    loadAllItems,
    createItem,
    updateItem,
    createListIfMissing
} from './lfr-graph.js';
import { buildDemoItems } from './lfr-demo-data.js';
import {
    renderApp,
    readFormFromDom,
    readFiltersFromDom,
    readSetupFromDom
} from './lfr-ui.js';
import { buildIcs, downloadIcs } from './lfr-export.js';
import { syncLfrItemsToOutlookCalendar, loadCalendarTargetFromSetup } from './lfr-calendar-sync.js';
import {
    applyLfrRoleFromEntra,
    isLfrDemoRoleUiEnabled,
    readDemoRoleOverrideFromUrl
} from './lfr-entra-role.js';

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
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const a = window.ms365AuthGetAccountInfo();
            if (a) {
                return {
                    email: String(a.username || a.upn || a.mail || a.email || '').trim(),
                    name: String(a.name || '').trim()
                };
            }
        }
    } catch {
        /* ignore */
    }
    try {
        const upn =
            typeof window.ms365AuthGetUserPrincipalName === 'function'
                ? window.ms365AuthGetUserPrincipalName()
                : '';
        if (upn) return { email: String(upn).trim(), name: '' };
    } catch {
        /* ignore */
    }
    return { email: '', name: '' };
}

function syncAccount() {
    const a = currentAccount();
    state.accountEmail = String(a.email || '').trim().toLowerCase();
    state.accountName = String(a.name || '').trim();
    state.stammdaten = loadStammdaten();
    prefillLehrerFormIdentity(state);
}

async function syncAccountAndRole() {
    syncAccount();
    await applyLfrRoleFromEntra(state, {
        demoRoleOverride: state.demoRoleOverride,
        preferredActiveRole: state.role
    });
    if (!viewsForRole(state.role).some((v) => v.id === state.view)) {
        state.view = viewsForRole(state.role)[0]?.id || 'dashboard';
    }
    prefillLehrerFormIdentity(state);
}

function findItem(key) {
    const k = String(key || '');
    return state.items.find((it) => it.antragId === k || it.itemId === k) || null;
}

async function refreshData() {
    state.loading = true;
    state.error = '';
    paint();
    await syncAccountAndRole();
    try {
        if (!state.siteUrl) {
            state.items = loadDemoItems().length
                ? loadDemoItems()
                : buildDemoItems(state.accountEmail, state.accountName);
            state.localDemoOnly = true;
            saveDemoItems(state.items);
            state.loading = false;
            paint();
            return;
        }
        const ctx = await resolveLfrContext(state.siteUrl, {
            listName: state.listName,
            listId: state.listId
        });
        state.ctx = ctx;
        state.listId = ctx.list.id;
        persistSetupCfg({ listId: ctx.list.id, listName: ctx.list.title, siteUrl: state.siteUrl });
        state.items = await loadAllItems(ctx);
        state.localDemoOnly = false;
        state.loading = false;
        paint();
    } catch (e) {
        state.loading = false;
        const msg = e && e.message ? e.message : String(e);
        state.error = msg;
        if (!state.items.length) {
            state.items = buildDemoItems(state.accountEmail, state.accountName);
            state.localDemoOnly = true;
            saveDemoItems(state.items);
        }
        paint();
    }
}

async function submitAntrag() {
    syncAccount();
    const fromDom = readFormFromDom();
    const draft = {
        ...state.form,
        ...fromDom,
        lehrerEmail: state.accountEmail || fromDom.lehrerEmail || state.form.lehrerEmail,
        lehrerName:
            state.accountName ||
            fromDom.lehrerName ||
            state.form.lehrerName ||
            (matchTeacherByEmail(state.stammdaten.teachers, state.accountEmail) || {}).name ||
            ''
    };
    const v = validateAntrag(draft);
    if (!v.ok) {
        toast(v.errors.join('\n'));
        return;
    }
    const row = {
        ...draft,
        antragId: newAntragId(),
        status: 'Ausstehend'
    };
    state.loading = true;
    paint();
    try {
        if (state.localDemoOnly || !state.ctx) {
            row.itemId = 'local-' + Date.now();
            state.items = [row, ...state.items];
            saveDemoItems(state.items);
        } else {
            const created = await createItem(state.ctx, row);
            state.items = [created, ...state.items.filter((x) => x.itemId !== created.itemId)];
        }
        const match = matchTeacherByEmail(state.stammdaten.teachers, state.accountEmail);
        state.form = emptyForm({
            lehrerEmail: state.accountEmail,
            lehrerName: state.accountName || (match && match.name) || ''
        });
        state.view = 'meine';
        state.loading = false;
        toast('Antrag gespeichert – Freigabe durch die Direktion ausstehend.');
        paint();
    } catch (e) {
        state.loading = false;
        toast(e && e.message ? e.message : String(e));
        paint();
    }
}

async function decideItem(key, approved) {
    const it = findItem(key);
    if (!it) return;
    const today = toIsoDateOnly(new Date()) || '';
    const actor = state.accountName || state.accountEmail || 'Direktion';
    const next = {
        ...it,
        status: approved ? 'Genehmigt' : 'Abgelehnt',
        genehmigtVon: approved ? actor : '',
        genehmigtAm: approved ? today : '',
        abgelehntVon: approved ? '' : actor,
        abgelehntAm: approved ? '' : today
    };
    state.loading = true;
    paint();
    try {
        if (state.localDemoOnly || !state.ctx || !it.itemId || String(it.itemId).startsWith('local-')) {
            state.items = state.items.map((x) =>
                x.antragId === it.antragId || x.itemId === it.itemId ? next : x
            );
            saveDemoItems(state.items);
        } else {
            await updateItem(state.ctx, it.itemId, next);
            state.items = state.items.map((x) => (x.itemId === it.itemId ? { ...x, ...next } : x));
        }
        state.loading = false;
        toast(approved ? 'Genehmigt.' : 'Abgelehnt.');
        paint();
    } catch (e) {
        state.loading = false;
        toast(e && e.message ? e.message : String(e));
        paint();
    }
}

function bindStatic() {
    if (!root) return;
    root.querySelectorAll('[data-lfr-view]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-lfr-view') || 'dashboard';
            if (!viewsForRole(state.role).some((v) => v.id === state.view)) {
                state.view = viewsForRole(state.role)[0]?.id || 'dashboard';
            }
            if (state.view === 'antrag') syncAccount();
            paint();
        });
    });
    root.querySelectorAll('[data-lfr-view-jump]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-lfr-view-jump') || 'dashboard';
            if (state.view === 'antrag') syncAccount();
            paint();
        });
    });
    root.querySelectorAll('[data-lfr-role]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const demoUi = isLfrDemoRoleUiEnabled(
                state.entraGroupsConfigured,
                state.demoRoleOverride
            );
            if (!demoUi && !(state.planerRoles || []).includes(btn.getAttribute('data-lfr-role') || '')) {
                return;
            }
            const r = btn.getAttribute('data-lfr-role');
            state.role = r === 'direktion' ? 'direktion' : 'lehrer';
            if (demoUi) state.demoRoleOverride = true;
            persistRole(state.role);
            if (!viewsForRole(state.role).some((v) => v.id === state.view)) {
                state.view = viewsForRole(state.role)[0]?.id || 'dashboard';
            }
            paint();
        });
    });
    const submit = document.getElementById('lfrFormSubmit');
    if (submit) submit.addEventListener('click', () => submitAntrag());
    const reset = document.getElementById('lfrFormReset');
    if (reset) {
        reset.addEventListener('click', () => {
            syncAccount();
            const match = matchTeacherByEmail(state.stammdaten.teachers, state.accountEmail);
            state.form = emptyForm({
                lehrerEmail: state.accountEmail,
                lehrerName: state.accountName || (match && match.name) || ''
            });
            paint();
        });
    }
    ['lfrFilterStatus', 'lfrFilterKat', 'lfrFilterQ'].forEach((id) => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('change', () => {
                state.filters = readFiltersFromDom(state);
                paint();
            });
            if (id === 'lfrFilterQ') el.addEventListener('input', () => {
                state.filters = readFiltersFromDom(state);
                paint();
            });
        }
    });
    const prev = document.getElementById('lfrCalPrev');
    const next = document.getElementById('lfrCalNext');
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
    root.querySelectorAll('[data-lfr-approve]').forEach((btn) => {
        btn.addEventListener('click', () => decideItem(btn.getAttribute('data-lfr-approve'), true));
    });
    root.querySelectorAll('[data-lfr-reject]').forEach((btn) => {
        btn.addEventListener('click', () => decideItem(btn.getAttribute('data-lfr-reject'), false));
    });
    const icsBtn = document.getElementById('lfrBtnIcs');
    if (icsBtn) {
        icsBtn.addEventListener('click', () => {
            const approved = state.items.filter((i) => normalizeStatus(i.status) === 'Genehmigt');
            downloadIcs(buildIcs(approved), 'lehrer-freistellungen.ics');
        });
    }
    async function runOutlookSync(opts) {
        const o = opts || {};
        const buttons = [
            document.getElementById('lfrBtnOutlookSync'),
            document.getElementById('lfrBtnOutlookSyncMe')
        ].filter(Boolean);
        buttons.forEach((b) => {
            b.disabled = true;
        });
        try {
            syncAccount();
            const target = loadCalendarTargetFromSetup();
            const res = await syncLfrItemsToOutlookCalendar(state.items, {
                calendarUser: o.useMe ? state.accountEmail : target.calendarUser,
                calendarId: o.useMe ? '' : target.calendarId,
                signedInEmail: state.accountEmail,
                useMe: !!o.useMe,
                onlyApproved: true
            });
            if (res.fail) {
                toast('Sync: ' + res.ok + ' ok, ' + res.fail + ' Fehler. ' + (res.errors[0] || ''));
            } else {
                toast(
                    res.ok +
                        ' Termin(e) im Outlook-Kalender aktualisiert' +
                        (o.useMe ? ' (dein Kalender).' : '.')
                );
            }
        } catch (e) {
            toast(e && e.message ? e.message : String(e));
        } finally {
            buttons.forEach((b) => {
                b.disabled = false;
            });
            const setupBtn = document.getElementById('lfrBtnOutlookSync');
            if (setupBtn && !loadCalendarTargetFromSetup().calendarUser) setupBtn.disabled = true;
        }
    }
    const outlookBtn = document.getElementById('lfrBtnOutlookSync');
    if (outlookBtn) {
        outlookBtn.addEventListener('click', () => {
            if (!loadCalendarTargetFromSetup().calendarUser) {
                toast('Bitte im IT-Setup (Schritt 4) den Kalender-Besitzer eintragen – oder „In meinen Kalender“ nutzen.');
                return;
            }
            runOutlookSync({ useMe: false });
        });
    }
    const outlookMeBtn = document.getElementById('lfrBtnOutlookSyncMe');
    if (outlookMeBtn) {
        outlookMeBtn.addEventListener('click', () => {
            if (!state.accountEmail) {
                toast('Bitte zuerst anmelden.');
                return;
            }
            runOutlookSync({ useMe: true });
        });
    }
    const saveSetup = document.getElementById('lfrBtnSaveSetup');
    if (saveSetup) {
        saveSetup.addEventListener('click', () => {
            const s = readSetupFromDom(state);
            state.siteUrl = s.siteUrl;
            state.listName = s.listName || state.listName;
            state.listId = s.listId;
            persistSiteUrl(s.siteUrl);
            persistSetupCfg(s);
            toast('Einstellungen gespeichert.');
        });
    }
    const ensureList = document.getElementById('lfrBtnEnsureList');
    if (ensureList) {
        ensureList.addEventListener('click', async () => {
            const s = readSetupFromDom(state);
            state.siteUrl = s.siteUrl;
            state.listName = s.listName || state.listName;
            persistSiteUrl(s.siteUrl);
            persistSetupCfg(s);
            state.loading = true;
            paint();
            try {
                let ctx = await resolveLfrContext(state.siteUrl, { listName: state.listName, listId: s.listId });
                const log = (msg) => toast(msg);
                ctx = await createListIfMissing(ctx, state.listName, (m) => log(m));
                state.ctx = ctx;
                state.listId = ctx.list.id;
                persistSetupCfg({ listId: ctx.list.id, listName: ctx.list.title, siteUrl: state.siteUrl });
                await refreshData();
                toast('Liste bereit.');
            } catch (e) {
                state.loading = false;
                toast(e && e.message ? e.message : String(e));
                paint();
            }
        });
    }
    const reload = document.getElementById('lfrBtnReload');
    if (reload) reload.addEventListener('click', () => refreshData());
    const refreshTop = document.getElementById('lfrBtnRefresh');
    if (refreshTop) refreshTop.addEventListener('click', () => refreshData());
}

function boot() {
    root = document.getElementById('lfrApp');
    if (!root) return;
    state.demoRoleOverride = readDemoRoleOverrideFromUrl(false);
    state.role = resolveRole();
    state.form = emptyForm();
    window.addEventListener('ms365-auth-state-changed', () => {
        void syncAccountAndRole().then(() => paint());
    });
    void refreshData();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}
