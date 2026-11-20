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
    clearDemoItems,
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
    readSetupFromDom,
    readCalendarSetupFromDom
} from './lfr-ui.js';
import { persistLfrApproverEmail } from './lfr-permissions.js';
import { applyLfrListPermissions } from './lfr-list-permissions.js';
import { normalizeSystemTab } from '../../shared/system-admin-shell.js';
import { buildIcs, downloadIcs } from './lfr-export.js';
import { syncLfrItemsToOutlookCalendar, loadCalendarTargetFromSetup } from './lfr-calendar-sync.js';
import {
    applyLfrRoleFromEntra,
    isLfrDemoRoleUiEnabled,
    readDemoRoleOverrideFromUrl,
    shellPreferredLfrRole
} from './lfr-entra-role.js';
import { startOfWeekMonday } from '../../shared/ms365-calendar-shell.js';
import { announcePlannerState } from '../../shared/app-role-chrome.js';

const state = createInitialState();
let root = null;

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else window.alert(msg);
}

function paint() {
    if (!root) return;
    announcePlannerState('lehrer-freistellung', state);
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

/**
 * MSAL steht oft erst nach dem ersten Paint. Vorher ist das Konto leer
 * und die App speichert fälschlich „Lehrkraft“.
 * @returns {Promise<{ email: string, name: string }>}
 */
async function waitForAuthAccount() {
    try {
        if (typeof window.ms365AuthEnsureInitialized === 'function') {
            await window.ms365AuthEnsureInitialized();
        }
    } catch {
        /* ignore */
    }
    const immediate = currentAccount();
    if (immediate.email) return immediate;
    const started = Date.now();
    while (Date.now() - started < 4500) {
        await new Promise((resolve) => setTimeout(resolve, 120));
        const next = currentAccount();
        if (next.email) return next;
    }
    return currentAccount();
}

/** @type {Promise<void>} */
let accountSyncQueue = Promise.resolve();

/**
 * @param {{ preferShellDefault?: boolean, preferredActiveRole?: string|null }} [opts]
 */
function enqueueAccountSync(opts) {
    const job = accountSyncQueue.then(() => syncAccountAndRole(opts));
    accountSyncQueue = job.then(
        () => {},
        () => {}
    );
    return job;
}

/**
 * @param {{ preferShellDefault?: boolean, preferredActiveRole?: string|null }} [opts]
 */
async function syncAccountAndRole(opts) {
    const account = await waitForAuthAccount();
    state.accountEmail = String(account.email || '').trim().toLowerCase();
    state.accountName = String(account.name || '').trim();
    state.stammdaten = loadStammdaten();
    prefillLehrerFormIdentity(state);
    // Demo-Direktor → Stammdaten Schulleitung, bevor Rollen aus Stammdaten gelesen werden
    try {
        const { ensureSignedInMs365DemoPortalAccountInLocalStammdaten } = await import(
            '../../shared/demo-portal-teacher.js'
        );
        ensureSignedInMs365DemoPortalAccountInLocalStammdaten();
        state.stammdaten = loadStammdaten();
    } catch {
        /* ignore */
    }
    // Lehrer-Browser: leere lokale Liste von der Tool-Site nachziehen
    try {
        const { ensureLiveSchoolStammdatenForBrowser } = await import(
            '../../shared/ensure-live-school-stammdaten.js'
        );
        const { loadEffectiveTeachersForPersona } = await import(
            '../../shared/live-school-stammdaten-cache.js'
        );
        if (!loadEffectiveTeachersForPersona().length) {
            await ensureLiveSchoolStammdatenForBrowser({ force: false });
            state.stammdaten = loadStammdaten();
        }
    } catch {
        /* ignore */
    }
    const preferShellDefault = !!(opts && opts.preferShellDefault);
    const hasPreferred = opts && Object.prototype.hasOwnProperty.call(opts, 'preferredActiveRole');
    await applyLfrRoleFromEntra(state, {
        demoRoleOverride: state.demoRoleOverride,
        preferShellDefault,
        ...(hasPreferred
            ? { preferredActiveRole: opts.preferredActiveRole }
            : preferShellDefault
              ? { preferredActiveRole: null }
              : { preferredActiveRole: state.role })
    });
    if (!viewsForRole(state.role).some((v) => v.id === state.view)) {
        state.view = viewsForRole(state.role)[0]?.id || 'dashboard';
    }
    prefillLehrerFormIdentity(state);
}

async function onItShellChanged() {
    const shellRole = shellPreferredLfrRole(state.planerRoles || []);
    await enqueueAccountSync({
        preferShellDefault: true,
        preferredActiveRole: shellRole
    });
    paint();
}

function findItem(key) {
    const k = String(key || '');
    return state.items.find((it) => it.antragId === k || it.itemId === k) || null;
}

let loadGen = 0;

async function refreshData() {
    const gen = ++loadGen;
    state.loading = true;
    state.error = '';
    paint();
    await enqueueAccountSync({ preferShellDefault: true });
    if (gen !== loadGen) return;
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
        if (gen !== loadGen) return;
        state.localDemoOnly = false;
        state.loading = false;
        paint();
    } catch (e) {
        if (gen !== loadGen) return;
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
        state.detailId = null;
        state.loading = false;
        toast(approved ? 'Genehmigt.' : 'Abgelehnt.');
        paint();
    } catch (e) {
        state.loading = false;
        toast(e && e.message ? e.message : String(e));
        paint();
    }
}

function alignCalendarToToday() {
    const now = new Date();
    state.calYear = now.getFullYear();
    state.calMonth = now.getMonth() + 1;
    state.calWeekAnchor = startOfWeekMonday(now).toISOString();
}

function bindStatic() {
    if (!root) return;
    root.querySelectorAll('[data-lfr-view]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-lfr-view') || 'dashboard';
            state.detailId = null;
            if (!viewsForRole(state.role).some((v) => v.id === state.view)) {
                state.view = viewsForRole(state.role)[0]?.id || 'dashboard';
            }
            if (state.view === 'einrichtung' && !state.systemTab) state.systemTab = 'site';
            if (state.view === 'antrag') syncAccount();
            paint();
        });
    });
    root.querySelectorAll('[data-lfr-view-jump]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-lfr-view-jump') || 'dashboard';
            state.detailId = null;
            if (state.view === 'antrag') syncAccount();
            paint();
        });
    });
    root.querySelectorAll('[data-lfr-sys-tab]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.systemTab = normalizeSystemTab(btn.getAttribute('data-lfr-sys-tab'));
            paint();
        });
    });
    root.querySelectorAll('[data-lfr-role]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const demoUi = isLfrDemoRoleUiEnabled(
                state.entraGroupsConfigured,
                state.demoRoleOverride
            );
            if (!demoUi) return;
            const r = btn.getAttribute('data-lfr-role');
            state.role = r === 'direktion' ? 'direktion' : 'lehrer';
            state.demoRoleOverride = true;
            persistRole(state.role);
            if (!viewsForRole(state.role).some((v) => v.id === state.view)) {
                state.view = viewsForRole(state.role)[0]?.id || 'dashboard';
            }
            paint();
        });
    });
    root.querySelectorAll('[data-lfr-detail]').forEach((el) => {
        const open = () => {
            state.detailId = el.getAttribute('data-lfr-detail');
            paint();
        };
        el.addEventListener('click', (ev) => {
            if (el.tagName === 'TR' && ev.target instanceof Element && ev.target.closest('button, a, input, select, textarea')) {
                return;
            }
            open();
        });
        if (el.tagName === 'TR') {
            el.addEventListener('keydown', (ev) => {
                if (ev.key !== 'Enter' && ev.key !== ' ') return;
                ev.preventDefault();
                open();
            });
        }
    });
    root.querySelectorAll('[data-lfr-close-detail]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.detailId = null;
            paint();
        });
    });
    const antragForm = document.getElementById('lfrAntragForm');
    if (antragForm) antragForm.addEventListener('submit', (ev) => ev.preventDefault());
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
            if (id === 'lfrFilterQ') {
                el.addEventListener('input', () => {
                    const start = el instanceof HTMLInputElement ? el.selectionStart : null;
                    state.filters = readFiltersFromDom(state);
                    paint();
                    const next = document.getElementById('lfrFilterQ');
                    if (next instanceof HTMLInputElement) {
                        next.focus();
                        if (typeof start === 'number') next.setSelectionRange(start, start);
                    }
                });
            }
        }
    });
    root.querySelectorAll('[data-lfr-cal-mode]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const mode = String(btn.getAttribute('data-lfr-cal-mode') || 'month');
            state.calMode = mode === 'agenda' || mode === 'week' || mode === 'month' ? mode : 'month';
            paint();
        });
    });
    root.querySelectorAll('[data-lfr-cal-nav]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const delta = Number(btn.getAttribute('data-lfr-cal-nav') || '0');
            const mode = String(state.calMode || 'month');
            if (mode === 'week' || mode === 'agenda') {
                if (delta === 0) {
                    alignCalendarToToday();
                } else {
                    const anchor = startOfWeekMonday(new Date(state.calWeekAnchor || Date.now()));
                    anchor.setDate(anchor.getDate() + delta * 7);
                    state.calWeekAnchor = startOfWeekMonday(anchor).toISOString();
                    state.calYear = anchor.getFullYear();
                    state.calMonth = anchor.getMonth() + 1;
                }
            } else if (delta === 0) {
                alignCalendarToToday();
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
            paint();
        });
    });
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
    if (reload) {
        reload.addEventListener('click', () => {
            const s = readSetupFromDom(state);
            if (s.siteUrl) {
                state.siteUrl = s.siteUrl;
                state.listName = s.listName || state.listName;
                state.listId = s.listId;
                persistSiteUrl(s.siteUrl);
                persistSetupCfg(s);
            }
            refreshData();
        });
    }
    const saveApprover = document.getElementById('lfrBtnSaveApprover');
    if (saveApprover) {
        saveApprover.addEventListener('click', () => {
            const el = document.getElementById('lfrSysApproverEmail');
            const mail = el ? String(el.value || '').trim() : '';
            persistLfrApproverEmail(mail);
            toast(mail ? 'Genehmiger gespeichert.' : 'Genehmiger geleert.');
            paint();
        });
    }
    const applyPerms = document.getElementById('lfrBtnApplyListPerms');
    if (applyPerms && !applyPerms.dataset.lfrWired) {
        applyPerms.dataset.lfrWired = '1';
        applyPerms.addEventListener('click', () => {
            const site = String(state.siteUrl || '').trim();
            if (!site) {
                toast('Zuerst SharePoint-URL setzen.');
                return;
            }
            const logEl = document.getElementById('lfrPermsApplyLog');
            const lines = [];
            const write = (m) => {
                lines.push(String(m || ''));
                if (logEl) {
                    logEl.hidden = false;
                    logEl.textContent = lines.join('\n');
                }
            };
            applyPerms.disabled = true;
            write('Listenrechte werden geschrieben …');
            applyLfrListPermissions(site, state.listName || 'Lehrer-Freistellungen', {
                listId: state.listId,
                flowServiceAccount: ''
            }, write)
                .then(() => {
                    toast('Listenrechte angewendet – Anträge neu laden.');
                    return refreshData();
                })
                .catch((e) => {
                    const msg = e && e.message ? e.message : String(e);
                    write('Fehler: ' + msg);
                    toast(msg);
                })
                .finally(() => {
                    applyPerms.disabled = false;
                });
        });
    }
    const saveCal = document.getElementById('lfrBtnSaveCalendar');
    if (saveCal) {
        saveCal.addEventListener('click', () => {
            const c = readCalendarSetupFromDom(state);
            state.outlookCalendarUser = c.outlookCalendarUser;
            state.outlookCalendarId = c.outlookCalendarId;
            persistSetupCfg(c);
            toast('Kalender-Ziel gespeichert.');
            paint();
        });
    }
    const demoReset = document.getElementById('lfrBtnDemoReset');
    if (demoReset) {
        demoReset.addEventListener('click', () => {
            if (
                !window.confirm(
                    'Lokale Demo-Daten wirklich löschen? SharePoint-Daten bleiben unverändert.'
                )
            ) {
                return;
            }
            clearDemoItems();
            state.items = [];
            state.localDemoOnly = false;
            toast('Lokale Demo-Daten zurückgesetzt.');
            void refreshData();
        });
    }
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
        void refreshData();
    });
    window.addEventListener('ms365-it-shell-changed', () => {
        void onItShellChanged();
    });
    window.addEventListener('keydown', (ev) => {
        if (ev.key === 'Escape' && state.detailId) {
            state.detailId = null;
            paint();
        }
    });
    void refreshData();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}
