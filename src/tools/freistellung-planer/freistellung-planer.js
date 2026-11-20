/**
 * Entry / Wiring: Freistellungs-Planer
 */
import {
    createInitialState,
    loadStammdaten,
    resolveRole,
    persistRole,
    persistSiteUrl,
    emptyForm,
    filterItems,
    scopeFromState,
    resolveKvForClass,
    applyKvFromClass,
    prefillStudentFreistellungForm,
    loadSetupCfg,
    loadSavedSiteUrl,
    resolveFreistellungSiteUrl,
    matchStudentByEmail,
    refreshStudentMatchFromStammdaten,
    clearStudentKlassePick,
    studentKlasseFromRecord,
    matchKvByClassHeadEmail,
    classRowsForKvScope,
    resolveStudentKlasseCode,
    isStudentKlasseLocked,
    persistDemoKlasseCode,
    loadDemoKlasseCode,
    viewsForRole,
    useStudentPlanerChrome,
    useKvPlanerChrome,
    useDirektionPlanerChrome,
    useMinimalPlanerChrome,
    isLikelyFreistellungStaffAccount,
    canDeleteFreistellung,
    itemVisibleForRole,
    effectiveAccountEmail
} from './freistellung-planer-state.js';
import { hydrateStammdatenFromLiveSite } from '../../shared/school-stammdaten-spo-lists.js';
import { announcePlannerState } from '../../shared/app-role-chrome.js';
import {
    toolSiteUrlFromLicenseResult,
    applyToolSiteUrlLocally,
    bootstrapSchoolSiteFromLicense
} from '../../shared/tenant-school-site-bootstrap.js';
import { startOfWeekMonday } from '../../shared/ms365-calendar-shell.js';
import {
    applyPlanerRoleFromEntra,
    isPlanerDemoRoleUiEnabled,
    finalizePlanerRoles
} from './freistellung-planer-entra-role.js';
import {
    accountIsDirektionPlannerUser,
    accountIsPlannerUserInList
} from './freistellung-planer-direktion-users.js';
import {
    entraGroupsConfigured,
    loadPermissionsConfig,
    loadEffectivePermissionsConfig
} from './freistellung-planer-permissions.js';
import { syncPlannerPermissionsFromSite } from './freistellung-planer-remote-config.js';
import { ensureFreistellungPlanerSchoolContext } from './freistellung-planer-bootstrap.js';
import { wireFreistellungAccessDebug } from './freistellung-planer-access-debug-ui.js';
import {
    mergeKategorieChoices,
    loadExtraKategorien
} from './freistellung-planer-kategorien.js';
import {
    readPendingUploadFiles,
    uploadFreistellungNachweise,
    saveNachweiseOnItem
} from './freistellung-planer-nachweise.js';
import {
    patchFreistellungKategorieColumn,
    patchFreistellungKlasseColumn,
    fetchFreistellungKlasseColumnChoices
} from './freistellung-planer-graph.js';
import { wireFreistellungKategorienAdmin } from './freistellung-kategorien-ui.js';
import { wireFreistellungKlassenAdmin } from './freistellung-klassen-ui.js';
import {
    resolveFrContext,
    loadAllFreistellungenForPlaner,
    createFreistellungItem,
    updateFreistellungStatus,
    deleteFreistellungItem,
    probeFreistellungListRead
} from './freistellung-planer-graph.js';
import { LIST_TITLE_DEFAULT } from './freistellung-planer-schema.js';
import {
    validateFreistellung,
    deriveKvClassCodesFromFreistellungItems,
    assertFrOwnership
} from './freistellung-planer-logic.js';
import {
    renderApp,
    readFormFromDom,
    readFiltersFromDom,
    bindFreistellungUploadField,
    clearFreistellungUploadField,
    wireFreistellungStudentKvSearch
} from './freistellung-planer-ui.js';
import { downloadFreistellungCsv } from './freistellung-planer-export.js';
import {
    buildLocalDemoItems,
    getDemoSeedPackage,
    parseDemoImportJson,
    itemsFromDemoPack,
    DEMO_SITE_DEFAULT,
    DEMO_SEED_TAG,
    DEMO_KV_PERSONA
} from './freistellung-planer-demo.js';
import {
    applyDemoStammdatenLocal,
    seedDemoFreistellungen,
    resetDemoFreistellungen
} from './freistellung-planer-demo-seed.js';
import { resolveStudentKlasseForAccount } from './freistellung-planer-student-klasse.js';

const state = createInitialState();
let root = null;

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else window.alert(msg);
}

function syncPageHeader() {
    const page = document.querySelector('.fr-page');
    if (!page) return;
    const student = useStudentPlanerChrome(state);
    const kv = useKvPlanerChrome(state);
    const direktion = useDirektionPlanerChrome(state);
    page.setAttribute('data-fr-student', student ? '1' : '0');
    page.setAttribute('data-fr-kv', kv ? '1' : '0');
    page.setAttribute('data-fr-direktion', direktion ? '1' : '0');
}

async function ensureKlasseChoicesLoaded() {
    if (state.role !== 'schueler') return false;
    const catalog = (loadEffectivePermissionsConfig().classCatalog || []).filter(
        (r) => r && r.code && !/^[1-5]AHW$/i.test(String(r.code))
    );
    if (catalog.length) {
        state.klasseColumnChoices = catalog.map((r) => ({
            code: r.code,
            name: r.name || r.code
        }));
        return true;
    }
    let ctx = state.ctx;
    if (!ctx || !ctx.list || !ctx.list.id) {
        const site = resolveFreistellungSiteUrl(state.siteUrl);
        if (!site) return false;
        try {
            ctx = await resolveFrContext(site, {
                listName: state.listName,
                listId: state.listId
            });
            state.ctx = ctx;
            if (ctx && ctx.list && ctx.list.id) state.listId = String(ctx.list.id);
        } catch {
            return false;
        }
    }
    try {
        const fromSp = await fetchFreistellungKlasseColumnChoices(ctx);
        state.klasseColumnChoices = (fromSp || []).filter(
            (r) => r && r.code && !/^[1-5]AHW$/i.test(String(r.code))
        );
    } catch {
        state.klasseColumnChoices = [];
    }
    return !!(state.klasseColumnChoices && state.klasseColumnChoices.length);
}

async function afterStudentAntragViewOpen() {
    if (state.role !== 'schueler') return;
    await ensureKlasseChoicesLoaded();
    prefillStudentFreistellungForm(state);
    paint();
}

function paint() {
    if (!root) return;
    announcePlannerState('freistellung', state);
    if (state.role === 'schueler') prefillStudentFreistellungForm(state);
    syncPageHeader();
    renderApp(state, root);
    bindStatic();
    wireFreistellungAccessDebug(root, state);
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
    try {
        const info =
            typeof window.ms365AuthGetAccountInfo === 'function' ? window.ms365AuthGetAccountInfo() : null;
        const upn =
            typeof window.ms365AuthGetUserPrincipalName === 'function'
                ? window.ms365AuthGetUserPrincipalName()
                : '';
        const a = currentAccount();
        state.accountEmail = String((info && (info.username || info.mail)) || a.email || upn || '')
            .trim()
            .toLowerCase();
        state.accountName = String((info && info.name) || a.name || '').trim();
    } catch {
        state.accountEmail = '';
        state.accountName = '';
    }
    state.stammdaten = loadStammdaten();
    state.studentMatch = matchStudentByEmail(state.accountEmail, state.stammdaten.students);
    clearStudentKlassePick(state.accountEmail);
    if (state.studentMatch && state.studentMatch.klasseSource === 'local-pick') {
        state.studentMatch = matchStudentByEmail(state.accountEmail, state.stammdaten.students);
    }
    if (studentKlasseFromRecord(state.studentMatch)) {
        state.demoKlasseCode = '';
        persistDemoKlasseCode('');
    }
    state.kvMatch = matchKvByClassHeadEmail(state.accountEmail, classRowsForKvScope(state));
    const dirMail = String(state.emailDirektion || loadSetupCfg().emailDirektion || '')
        .trim()
        .toLowerCase();
    state.direktionMatch = !!(dirMail && state.accountEmail && dirMail === state.accountEmail);
    if (state.role === 'schueler') {
        prefillStudentFreistellungForm(state);
    }
}

function friendlySpoError(e) {
    const msg = String((e && e.message) || e || '');
    if (useMinimalPlanerChrome(state)) {
        if (/access denied|403|forbidden/i.test(msg)) {
            if (useDirektionPlanerChrome(state)) {
                return 'Die Freistellungsliste ist für dieses Konto noch nicht erreichbar. Site und Berechtigungen liegen unter Systemadministration.';
            }
            return useKvPlanerChrome(state)
                ? 'Die Freistellungsliste ist für Ihr Konto noch nicht erreichbar. Bitte die Schulleitung oder IT informieren.'
                : 'Die Freistellungsliste ist für Ihr Konto noch nicht erreichbar. Bitte Klassenvorstand oder Sekretariat informieren.';
        }
        return 'Daten konnten nicht geladen werden. Bitte später erneut versuchen.';
    }
    if (/access denied|403|forbidden/i.test(msg)) {
        return (
            msg +
            ' – Im Freistellungen-Setup „Liste anlegen / prüfen“ oder „Nur Berechtigungen“ ausführen (Direktion/KV/Schüler-Gruppen). Planer-Gruppen: Listen-Beschreibung oder SiteAssets/ms365/freistellung-planer-groups.json'
        );
    }
    return msg;
}

async function ensurePlannerPermissionsConfig(listIdHint) {
    const setup = loadSetupCfg();
    const site = resolveFreistellungSiteUrl(state.siteUrl || setup.siteUrl);
    if (site) state.siteUrl = site;
    if (!site) return;
    let listId = String(listIdHint || state.listId || setup.listId || '').trim();
    const listName =
        String(state.listName || setup.listName || LIST_TITLE_DEFAULT).trim() || LIST_TITLE_DEFAULT;
    if (!listId) {
        try {
            const ctx = await resolveFrContext(site, { listName });
            if (ctx && ctx.list && ctx.list.id) {
                listId = String(ctx.list.id);
                state.listId = listId;
            }
        } catch {
            /* Liste erst nach Anmeldung / ohne Rechte */
        }
    }
    try {
        await syncPlannerPermissionsFromSite(site, listId || undefined, {
            listName,
            listId: listId || setup.listId || ''
        });
    } catch {
        /* optional – Rolle ggf. ohne Remote */
    }
}

function applyRecoveredPlanerRoles(roles, sources) {
    const finalized = finalizePlanerRoles(roles, sources);
    state.planerRoles = finalized.roles;
    state.planerRoleSources = finalized.sources;
    state.role = finalized.roles[0] || '';
    state.roleSource = (finalized.sources && finalized.sources[state.role]) || 'stammdaten';
    state.planerAccessDenied = finalized.roles.length === 0;
    state.roleHint = '';
    state.roleHintPublic = '';
    state.roleHintStaff = '';
}

/**
 * Nach Laden der Gruppen-Config aus SharePoint: Stammdaten, Liste, erneuter Entra-Versuch.
 */
async function tryRecoverPlanerAccess() {
    if (!state.planerAccessDenied) return;

    syncAccount();
    const site = resolveFreistellungSiteUrl(state.siteUrl);
    if (site) state.siteUrl = site;
    await ensurePlannerPermissionsConfig(state.listId);
    await applyPlanerRoleFromEntra(state, {
        demoRoleOverride: state.demoRoleOverride,
        preferredDemoRole: state.demoRoleOverride ? state.role : null,
        preferredActiveRole: state.role
    });
    if (!state.planerAccessDenied) return;

    const cfg = loadEffectivePermissionsConfig();

    if (state.kvMatch) {
        applyRecoveredPlanerRoles(['kv'], { kv: 'stammdaten' });
        return;
    }

    if (accountIsPlannerUserInList(state.accountEmail, cfg.kvUsers)) {
        applyRecoveredPlanerRoles(['kv'], { kv: 'kv-user' });
        return;
    }

    if (accountIsPlannerUserInList(state.accountEmail, cfg.schuelerUsers)) {
        applyRecoveredPlanerRoles(['schueler'], { schueler: 'schueler-user' });
        return;
    }

    if (
        state.direktionMatch ||
        accountIsDirektionPlannerUser(state.accountEmail, cfg.direktionUsers)
    ) {
        applyRecoveredPlanerRoles(
            ['direktion'],
            {
                direktion: state.direktionMatch ? 'setup-direktion' : 'direktion-user'
            }
        );
        return;
    }

    if (state.studentMatch) {
        try {
            const ctx = await resolveFrContext(state.siteUrl, {
                listName: state.listName,
                listId: state.listId
            });
            if (await probeFreistellungListRead(ctx)) {
                applyRecoveredPlanerRoles(['schueler'], { schueler: 'stammdaten' });
                return;
            }
        } catch {
            /* ignore */
        }
    }

    try {
        const ctx = await resolveFrContext(state.siteUrl, {
            listName: state.listName,
            listId: state.listId
        });
        if (!(await probeFreistellungListRead(ctx))) return;

        if (!isLikelyFreistellungStaffAccount(state)) {
            applyRecoveredPlanerRoles(['schueler'], { schueler: 'list-access' });
            return;
        }

        if (entraGroupsConfigured(cfg) && isLikelyFreistellungStaffAccount(state) && cfg.groupKvId) {
            applyRecoveredPlanerRoles(['kv'], { kv: 'list-access' });
            return;
        }
    } catch {
        /* ignore */
    }
}

async function grantSchuelerIfFreistellungListReadable() {
    if (!state.planerAccessDenied) return;
    if (isLikelyFreistellungStaffAccount(state)) return;
    const site = resolveFreistellungSiteUrl(state.siteUrl);
    if (!site) return;
    state.siteUrl = site;
    try {
        const ctx = await resolveFrContext(state.siteUrl, {
            listName: state.listName,
            listId: state.listId
        });
        if (!ctx || !ctx.list || !ctx.list.id) return;
        state.listId = String(ctx.list.id);
        if (!(await probeFreistellungListRead(ctx))) return;
        await syncPlannerPermissionsFromSite(state.siteUrl, state.listId, {
            listName: state.listName,
            listId: state.listId
        });
        await applyPlanerRoleFromEntra(state, {
            demoRoleOverride: state.demoRoleOverride,
            preferredDemoRole: state.demoRoleOverride ? state.role : null,
            preferredActiveRole: state.role
        });
        if (!state.planerAccessDenied) return;
        applyRecoveredPlanerRoles(['schueler'], { schueler: 'list-access' });
        state.stammdaten = loadStammdaten();
        await resolveStudentKlasseForAccount(state);
    } catch {
        /* ignore */
    }
}

async function resolvePlanerRole(listIdHint) {
    syncAccount();
    if (typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn()) {
        const boot = await ensureFreistellungPlanerSchoolContext({
            siteUrl: state.siteUrl,
            listId: listIdHint || state.listId,
            listName: state.listName
        });
        if (boot && boot.siteUrl) state.siteUrl = boot.siteUrl;
        if (boot && boot.listId) state.listId = boot.listId;
    }
    const site = resolveFreistellungSiteUrl(state.siteUrl);
    if (site) state.siteUrl = site;
    await ensurePlannerPermissionsConfig(listIdHint);
    await applyPlanerRoleFromEntra(state, {
        demoRoleOverride: state.demoRoleOverride,
        preferredDemoRole: state.demoRoleOverride ? state.role : null,
        preferredActiveRole: state.role
    });
    await tryRecoverPlanerAccess();
    await grantSchuelerIfFreistellungListReadable();
    if (!state.planerAccessDenied && state.role === 'schueler') {
        refreshStudentMatchFromStammdaten(state);
        await resolveStudentKlasseForAccount(state);
        await ensureKlasseChoicesLoaded();
    }
    const allowed = viewsForRole(state.role, state).map((v) => v.id);
    if (!allowed.includes(state.view)) state.view = allowed[0] || 'meine';
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
 * Klassen/Lehrer aus SharePoint-Live-Listen (Fallback: lokaler Cache).
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
        const local = loadStammdaten();
        const result = await hydrateStammdatenFromLiveSite({
            token: tok,
            siteId: String(siteId),
            siteWebUrl: String(state.siteUrl || '').trim(),
            graph: api,
            local
        });
        state.stammdaten = {
            classes: (result.stammdaten && result.stammdaten.classes) || local.classes || [],
            teachers: (result.stammdaten && result.stammdaten.teachers) || local.teachers || [],
            students: (result.stammdaten && result.stammdaten.students) || local.students || []
        };
        const liveClasses = (result.live && result.live.classes) || [];
        const liveTeachers = (result.live && result.live.teachers) || [];
        const missing = (result.live && result.live.missing) || [];
        const usedLive = !!(liveClasses.length || liveTeachers.length);
        let err = '';
        if (!result.ok && !liveClasses.length) err = result.error || 'Live-Laden fehlgeschlagen';
        else if (missing.length && !liveClasses.length) err = 'Listen fehlen: ' + missing.join(', ');
        state.stammdatenLive = {
            ok: liveClasses.length > 0,
            error: err,
            classCount: liveClasses.length,
            teacherCount: liveTeachers.length,
            missing,
            siteUrl: String(state.siteUrl || '').trim(),
            source: usedLive ? 'sharepoint' : result.ok ? 'local-fallback' : 'error'
        };
    } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        state.stammdatenLive = {
            ok: false,
            error: /nicht angemeldet|anmelden/i.test(msg)
                ? 'Bitte anmelden und danach neu laden – sonst bleiben die Listen leer.'
                : msg,
            classCount: (state.stammdaten.classes || []).length,
            teacherCount: (state.stammdaten.teachers || []).length,
            missing: [],
            source: 'error'
        };
    }
}

async function ensureFrSiteFromLicense() {
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
                'Keine Schul-Site in der Lizenz (Tool-Site-URL). IT: unter Admin → Schule die SharePoint-URL setzen.';
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
async function ensureFrSiteUrl() {
    let url = String(state.siteUrl || '').trim() || loadSavedSiteUrl();
    if (url) {
        state.siteUrl = url;
        return url;
    }
    url = await ensureFrSiteFromLicense();
    if (url) return url;
    url = loadSavedSiteUrl();
    if (url) {
        state.siteUrl = url;
        return url;
    }
    return '';
}

function applyDemoKvPersona() {
    if (!state.demoPersonaEmail) return;
    const email = String(state.demoPersonaEmail).trim().toLowerCase();
    state.kvMatch = matchKvByClassHeadEmail(email, classRowsForKvScope(state));
    if (!state.demoRoleOverride) {
        state.role = 'kv';
        state.planerRoles = ['kv', 'schueler'];
        state.planerAccessDenied = false;
    }
}

function ensureDemoPersonaInStammdaten() {
    const teachers = Array.isArray(state.stammdaten.teachers) ? state.stammdaten.teachers : [];
    if (!teachers.some((t) => String(t.email || '').toLowerCase() === DEMO_KV_PERSONA.email)) {
        state.stammdaten.teachers = [
            ...teachers,
            {
                code: DEMO_KV_PERSONA.code,
                name: DEMO_KV_PERSONA.name,
                email: DEMO_KV_PERSONA.email
            }
        ];
    }
}

function alignCalendarToToday() {
    const now = new Date();
    state.calYear = now.getFullYear();
    state.calMonth = now.getMonth() + 1;
    state.calWeekAnchor = startOfWeekMonday(now).toISOString();
    state.calMode = state.calMode || 'month';
}

async function refreshData() {
    state.loading = true;
    state.error = '';
    state.info = '';
    paint();
    try {
        syncAccount();
        if (!String(state.siteUrl || '').trim()) {
            const viaLicense = await ensureFrSiteUrl();
            if (!viaLicense) {
                const setup = loadSetupCfg();
                const fallback = loadSavedSiteUrl(setup);
                if (fallback) state.siteUrl = fallback;
            }
        }
        const setup = loadSetupCfg();
        if (setup.listName) state.listName = setup.listName;
        if (setup.listId) state.listId = setup.listId;
        if (setup.emailDirektion) state.emailDirektion = setup.emailDirektion;

        if (String(state.siteUrl || '').trim()) {
            try {
                const ctxEarly = await resolveFrContext(state.siteUrl, {
                    listName: state.listName,
                    listId: state.listId
                });
                if (ctxEarly && ctxEarly.list && ctxEarly.list.id) {
                    state.listId = String(ctxEarly.list.id);
                }
                if (ctxEarly && ctxEarly.siteId) {
                    await hydrateLiveStammdaten(ctxEarly.siteId);
                }
            } catch {
                /* Liste erst nach Rollen-Sync / ohne Rechte */
            }
        }

        await resolvePlanerRole(state.listId);
        applyDemoKvPersona();
        refreshKategorieChoicesState();

        if (state.planerAccessDenied) {
            await tryRecoverPlanerAccess();
            const allowedAfterRecover = viewsForRole(state.role, state).map((v) => v.id);
            if (!allowedAfterRecover.includes(state.view)) {
                state.view = allowedAfterRecover[0] || 'dashboard';
            }
        }

        if (state.planerAccessDenied) {
            state.ctx = null;
            if (!(state.localDemoOnly && (state.items || []).length)) {
                state.items = [];
            }
            return;
        }

        const ctx = await resolveFrContext(state.siteUrl, {
            listName: state.listName,
            listId: state.listId
        });
        const listIdFromCtx = ctx && ctx.list && ctx.list.id ? String(ctx.list.id) : '';
        if (listIdFromCtx) {
            state.listId = listIdFromCtx;
            await ensurePlannerPermissionsConfig(listIdFromCtx);
            await applyPlanerRoleFromEntra(state, {
                demoRoleOverride: state.demoRoleOverride,
                preferredDemoRole: state.demoRoleOverride ? state.role : null,
                preferredActiveRole: state.role
            });
            applyDemoKvPersona();
            const allowed = viewsForRole(state.role, state).map((v) => v.id);
            if (!allowed.includes(state.view)) state.view = allowed[0] || 'meine';
            if (state.planerAccessDenied) {
                state.ctx = null;
                if (!(state.localDemoOnly && (state.items || []).length)) {
                    state.items = [];
                }
                return;
            }
        }
        state.ctx = ctx;
        if (ctx && ctx.siteId) await hydrateLiveStammdaten(ctx.siteId);
        try {
            state.klasseColumnChoices = await fetchFreistellungKlasseColumnChoices(ctx);
        } catch {
            state.klasseColumnChoices = [];
        }
        state.localDemoOnly = false;
        state.items = await loadAllFreistellungenForPlaner(ctx, state);
        state.info = '';
        const scopeEmail = effectiveAccountEmail(state) || state.accountEmail;
        if (state.role === 'kv' || state.demoPersonaEmail) {
            state.kvClassCodesFromItems = deriveKvClassCodesFromFreistellungItems(
                state.items,
                scopeEmail,
                state.demoPersonaName || state.accountName
            );
            state.kvMatch = matchKvByClassHeadEmail(scopeEmail, classRowsForKvScope(state));
        }
    } catch (e) {
        state.loading = false;
        // Lokale Demo nicht mit SharePoint-403 übermalen
        if (state.localDemoOnly && (state.items || []).length) {
            state.error = '';
            console.warn('[fr] refresh skipped – local demo kept', e);
        } else {
            state.error = friendlySpoError(e) || 'Laden fehlgeschlagen';
            state.ctx = null;
        }
    } finally {
        state.loading = false;
        paint();
    }
}

function exportCsv() {
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, state.filters, scope);
    downloadFreistellungCsv(items);
    toast('CSV exportiert (' + items.length + ' Zeilen).');
}

function applyLocalDemo(pack) {
    syncAccount();
    const stamOk = applyDemoStammdatenLocal(pack.stammdaten);
    if (stamOk) state.stammdaten = loadStammdaten();
    ensureDemoPersonaInStammdaten();
    state.localDemoOnly = true;
    state.ctx = null;
    state.error = '';
    state.demoPersonaEmail = DEMO_KV_PERSONA.email;
    state.demoPersonaName = DEMO_KV_PERSONA.name;
    state.demoRoleOverride = state.demoRoleOverride || false;
    applyDemoKvPersona();
    const usePackRows = pack && Array.isArray(pack.freistellungen) && pack.freistellungen.length;
    state.items = usePackRows
        ? itemsFromDemoPack(pack, {
              accountEmail: DEMO_KV_PERSONA.email,
              accountName: DEMO_KV_PERSONA.name
          })
        : buildLocalDemoItems({
              accountEmail: DEMO_KV_PERSONA.email,
              accountName: DEMO_KV_PERSONA.name,
              classes: state.stammdaten.classes
          });
    state.demoSeedTag = (pack && pack.seedTag) || DEMO_SEED_TAG;
    state.view = 'dashboard';
    alignCalendarToToday();
    if (!state.siteUrl && (pack.siteUrl || pack.siteDefault || DEMO_SITE_DEFAULT)) {
        state.siteUrl = String(pack.siteUrl || pack.siteDefault || DEMO_SITE_DEFAULT).trim();
        persistSiteUrl(state.siteUrl);
    }
    const c = pack.counts || {};
    const label = (pack && pack.title) || 'Demo SJ ' + (pack.schoolYear || '2026/27');
    state.info =
        label +
        ' lokal (' +
        state.items.length +
        ' Anträge' +
        (c.ausstehend != null ? ', ' + c.ausstehend + ' offen' : '') +
        ').';
    return stamOk;
}

async function runPackImport(pack) {
    const stamOk = applyLocalDemo(pack);
    paint();
    const tag = state.demoSeedTag || DEMO_SEED_TAG;
    toast(
        'Paket geladen (' +
            state.items.length +
            ' Anträge)' +
            (stamOk ? ', Stammdaten übernommen' : '') +
            '.'
    );

    const siteInput = root && root.querySelector('#frSiteUrl');
    const siteUrl =
        String((siteInput && siteInput.value) || state.siteUrl || pack.siteUrl || DEMO_SITE_DEFAULT).trim();
    if (!siteUrl) return;

    const writeSp = window.confirm(
        'Daten auch auf SharePoint schreiben?\n\n' +
            'Site: ' +
            siteUrl +
            '\n' +
            state.items.length +
            ' Einträge (Upsert über id:fr-… in Beschreibung, Tag ' +
            tag +
            ').\n\n' +
            'Hinweis: Bei aktivem Freistellungs-Flow können Approvals für NEUE Einträge mit Status Ausstehend starten. Flow vorher pausieren empfohlen.\n' +
            'Bereits vorhandene Test-IDs werden aktualisiert (kein neuer Trigger).'
    );
    if (!writeSp) return;

    state.siteUrl = siteUrl;
    persistSiteUrl(siteUrl);
    state.loading = true;
    state.info = 'Schreibe auf SharePoint …';
    paint();
    try {
        const setup = loadSetupCfg();
        await seedDemoFreistellungen(siteUrl, (msg) => console.log('[fr-demo]', msg), {
            pack,
            listName: state.listName || setup.listName,
            listId: state.listId || setup.listId
        });
        state.localDemoOnly = false;
        await refreshData();
        toast('SharePoint-Seed fertig und neu geladen.');
    } catch (e) {
        state.localDemoOnly = true;
        state.loading = false;
        // Roh-403 nicht in die Banner-Zeile – lokal bleibt die Demo voll nutzbar.
        state.error = '';
        paint();
        toast('Demo bleibt lokal (SharePoint-Schreiben nicht möglich). Für die Vorführung reicht das.');
        console.warn('[fr-demo] SharePoint seed failed', e);
    }
}

async function runDemoImport() {
    await runPackImport(getDemoSeedPackage());
}

async function runJsonImportFromFile(file) {
    if (!file) return;
    const text = await file.text();
    const pack = parseDemoImportJson(text);
    await runPackImport(pack);
}

async function runDemoReset() {
    const localOnly = state.localDemoOnly && !state.ctx;
    const siteUrl = String(state.siteUrl || DEMO_SITE_DEFAULT).trim();
    const tag = state.demoSeedTag || DEMO_SEED_TAG;
    const ok = window.confirm(
        localOnly
            ? 'Lokale Demo-Daten leeren?'
            : 'Alle Demo-/Test-Einträge (Tag ' +
                  tag +
                  ') auf SharePoint löschen?\n\nSite: ' +
                  siteUrl +
                  '\n\nEchte (nicht-Demo) Einträge bleiben erhalten.'
    );
    if (!ok) return;

    if (localOnly) {
        state.items = [];
        state.localDemoOnly = false;
        state.demoPersonaEmail = '';
        state.demoPersonaName = '';
        state.info = 'Lokale Demo geleert.';
        paint();
        toast('Lokale Demo geleert.');
        return;
    }

    state.loading = true;
    paint();
    try {
        const setup = loadSetupCfg();
        const result = await resetDemoFreistellungen(
            siteUrl,
            (msg) => console.log('[fr-demo-reset]', msg),
            {
                listName: state.listName || setup.listName,
                listId: state.listId || setup.listId,
                seedTag: tag
            }
        );
        state.localDemoOnly = false;
        await refreshData();
        toast('Demo zurückgesetzt (' + (result.deleted || 0) + ' gelöscht).');
    } catch (e) {
        state.loading = false;
        state.error = String((e && e.message) || e);
        paint();
        toast(state.error);
    }
}

function refreshKategorieChoicesState() {
    state.kategorieChoices = mergeKategorieChoices(loadExtraKategorien());
}

async function submitAntrag() {
    if (state.planerAccessDenied) {
        toast('Keine Berechtigung für Anträge.');
        return;
    }
    const draft = readFormFromDom(root);
    state.form = { ...state.form, ...draft };
    let pendingFiles = [];
    const filesEl = root.querySelector('#frFormFiles');
    try {
        if (filesEl && filesEl.files && filesEl.files.length) {
            pendingFiles = readPendingUploadFiles(filesEl.files);
        }
    } catch (e) {
        toast(e && e.message ? e.message : String(e));
        return;
    }
    const ownership = assertFrOwnership(
        { ...state.form, _allowedKategorien: state.kategorieChoices },
        { accountEmail: state.accountEmail, accountName: state.accountName }
    );
    if (!ownership.ok) {
        toast(ownership.error || 'Antragsteller fehlt.');
        return;
    }
    state.form = { ...state.form, ...ownership.draft };
    const check = validateFreistellung({
        draft: { ...state.form, _allowedKategorien: state.kategorieChoices }
    });
    if (!check.ok) {
        toast(check.errors[0] || 'Bitte Formular prüfen.');
        return;
    }
    if (state.localDemoOnly || !state.ctx) {
        const id = 'demo-' + Date.now();
        const demoNachweise = pendingFiles.map((f) => ({ name: f.name, url: '' }));
        const ownerEmail = ownership.draft.authorEmail || state.accountEmail;
        state.items = [
            {
                itemId: id,
                titel: draft.schuelerName + ' (' + draft.klasse + ')',
                schuelerName: draft.schuelerName,
                klasse: draft.klasse,
                beginn: draft.beginn,
                ende: draft.ende,
                status: 'Ausstehend',
                kategorie: draft.kategorie,
                beschreibung: draft.beschreibung,
                bemerkungen: '',
                kvEmail: draft.kvEmail,
                kvName: state.form.kvName || '',
                authorEmail: ownerEmail,
                beantragtVon: ownerEmail,
                authorName: ownership.draft.authorName || state.accountName || '',
                dayCount: check.dayCount,
                multiDay: check.path.multiDay,
                approvalLabel: check.path.label,
                nachweise: demoNachweise
            },
            ...state.items
        ];
        state.form = emptyForm({ schuelerName: state.accountName });
        prefillStudentFreistellungForm(state);
        clearFreistellungUploadField(root);
        state.view = 'meine';
        toast(
            'Demo: Antrag lokal gespeichert' +
                (demoNachweise.length ? ' (' + demoNachweise.length + ' Datei(en) nur als Namen).' : '.')
        );
        paint();
        return;
    }
    state.loading = true;
    state.info =
        pendingFiles.length
            ? 'Antrag wird gespeichert und ' + pendingFiles.length + ' Datei(en) hochgeladen …'
            : '';
    paint();
    try {
        const created = await createFreistellungItem(state.ctx, ownership.draft);
        if (pendingFiles.length && created && created.itemId) {
            const links = await uploadFreistellungNachweise(state.ctx, created.itemId, pendingFiles);
            await saveNachweiseOnItem(state.ctx, created.itemId, links);
        }
        clearFreistellungUploadField(root);
        toast(
            'Antrag eingereicht' +
                (pendingFiles.length ? ' inkl. ' + pendingFiles.length + ' Anhang/Anhänge' : '') +
                ' – Genehmigung startet über Microsoft Approvals.'
        );
        state.form = emptyForm({ schuelerName: state.accountName });
        prefillStudentFreistellungForm(state);
        state.view = state.role === 'schueler' ? 'meine' : 'liste';
        await refreshData();
    } catch (e) {
        state.error = String((e && e.message) || e);
        state.loading = false;
        paint();
    }
}

async function setStatus(itemId, status) {
    if (state.localDemoOnly || !state.ctx) {
        state.items = state.items.map((it) =>
            String(it.itemId) === String(itemId) ? { ...it, status } : it
        );
        toast('Demo: Status → ' + status);
        paint();
        return;
    }
    try {
        await updateFreistellungStatus(state.ctx, itemId, {
            status,
            role: state.role,
            actorName: state.accountName,
            actorEmail: state.accountEmail
        });
        toast('Status aktualisiert: ' + status);
        await refreshData();
    } catch (e) {
        toast(String((e && e.message) || e));
    }
}

async function deleteAntrag(itemId) {
    if (!canDeleteFreistellung(state)) {
        toast('Löschen nur für Direktion/Admin.');
        return;
    }
    const it = (state.items || []).find((x) => String(x.itemId) === String(itemId));
    if (!it || !itemVisibleForRole(state, it, { scopeAll: true })) {
        toast('Antrag nicht gefunden.');
        return;
    }
    const label = String(it.schuelerName || it.titel || 'Antrag').trim() || 'Antrag';
    const ok = window.confirm(
        'Antrag „' + label + '“ wirklich löschen?\n\nDas kann nicht rückgängig gemacht werden.'
    );
    if (!ok) return;

    if (state.localDemoOnly || !state.ctx) {
        state.items = state.items.filter((x) => String(x.itemId) !== String(itemId));
        state.detailId = null;
        toast('Gelöscht (Demo).');
        paint();
        return;
    }

    state.loading = true;
    paint();
    try {
        await deleteFreistellungItem(state.ctx, itemId);
        state.detailId = null;
        toast('Antrag gelöscht.');
        await refreshData();
    } catch (e) {
        state.loading = false;
        state.error = String((e && e.message) || e);
        paint();
        toast(state.error);
    }
}

function bindStatic() {
    root.querySelectorAll('[data-fr-view]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-fr-view') || 'dashboard';
            if (state.view === 'antrag') {
                state.formStep = 1;
                if (state.role === 'schueler') {
                    afterStudentAntragViewOpen();
                    return;
                }
            }
            paint();
        });
    });
    root.querySelectorAll('[data-fr-sys-tab]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = String(btn.getAttribute('data-fr-sys-tab') || 'site');
            state.systemTab =
                id === 'perms' || id === 'kalender' || id === 'wartung' || id === 'danger' ? id : 'site';
            paint();
        });
    });
    root.querySelectorAll('[data-fr-view-jump]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-fr-view-jump') || 'dashboard';
            const fs = btn.getAttribute('data-fr-filter-status');
            if (fs != null) {
                state.filters = {
                    klasse: '',
                    status: String(fs),
                    kategorie: '',
                    multiDay: '',
                    q: ''
                };
                state.filterMultiOpen = { klasse: false, status: false, kategorie: false };
            }
            if (state.view === 'antrag') {
                state.formStep = 1;
                if (state.role === 'schueler') {
                    afterStudentAntragViewOpen();
                    return;
                }
            }
            paint();
        });
    });
    root.querySelectorAll('[data-fr-status-preset]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.filters = {
                ...state.filters,
                status: String(btn.getAttribute('data-fr-status-preset') || '')
            };
            paint();
        });
    });
    root.querySelectorAll('[data-fr-role]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const raw = btn.getAttribute('data-fr-role') || 'schueler';
            const role = raw === 'direktion' ? 'direktion' : raw === 'kv' ? 'kv' : 'schueler';
            const demoUi = isPlanerDemoRoleUiEnabled(state.entraGroupsConfigured, state.demoRoleOverride);
            const entitled = demoUi ? ['direktion', 'kv', 'schueler'] : state.planerRoles || [];
            if (!entitled.includes(role)) return;
            state.role = role;
            state.roleSource =
                (state.planerRoleSources && state.planerRoleSources[role]) || (demoUi ? 'demo' : 'entra');
            persistRole(role);
            state.roleHint = '';
            state.planerAccessDenied = false;
            const allowedViews = viewsForRole(role, state).map((v) => v.id);
            if (!allowedViews.includes(state.view)) {
                state.view = allowedViews[0] || (role === 'schueler' ? 'meine' : 'dashboard');
            }
            if (state.view === 'antrag' && role === 'schueler') {
                afterStudentAntragViewOpen();
                return;
            }
            paint();
        });
    });

    const demoKlasse = root.querySelector('#frDemoKlasse');
    if (demoKlasse) {
        demoKlasse.addEventListener('change', () => {
            state.demoKlasseCode = String(demoKlasse.value || '').trim();
            persistDemoKlasseCode(state.demoKlasseCode);
            if (state.demoKlasseCode) {
                state.form = { ...state.form, klasse: state.demoKlasseCode };
                applyKvFromClass(state);
            }
            paint();
        });
    }

    const site = root.querySelector('#frSiteUrl');
    const btnLoad = root.querySelector('#frBtnLoad');
    if (btnLoad) {
        btnLoad.addEventListener('click', () => {
            if (site) {
                state.siteUrl = String(site.value || '').trim();
                persistSiteUrl(state.siteUrl);
            }
            refreshData();
        });
    }
    const btnCsv = root.querySelector('#frBtnCsv');
    if (btnCsv) btnCsv.addEventListener('click', exportCsv);
    const btnCsvInline = root.querySelector('#frBtnCsvInline');
    if (btnCsvInline) btnCsvInline.addEventListener('click', exportCsv);
    const btnDemo = root.querySelector('#frBtnDemo');
    if (btnDemo) btnDemo.addEventListener('click', () => runDemoImport().catch((e) => toast(String((e && e.message) || e))));
    const btnJson = root.querySelector('#frBtnJsonImport');
    const jsonFile = root.querySelector('#frJsonImportFile');
    if (btnJson && jsonFile) {
        btnJson.addEventListener('click', () => jsonFile.click());
        jsonFile.addEventListener('change', () => {
            const f = jsonFile.files && jsonFile.files[0];
            jsonFile.value = '';
            if (!f) return;
            runJsonImportFromFile(f).catch((e) => toast(String((e && e.message) || e)));
        });
    }
    const btnDemoReset = root.querySelector('#frBtnDemoReset');
    if (btnDemoReset) {
        btnDemoReset.addEventListener('click', () =>
            runDemoReset().catch((e) => toast(String((e && e.message) || e)))
        );
    }

    if (state.view === 'administration' && state.role === 'direktion') {
        const siteCtx = () => ({
            siteUrl: state.siteUrl,
            listId: state.listId || (loadSetupCfg() && loadSetupCfg().listId) || ''
        });
        wireFreistellungKlassenAdmin(
            {
                mountId: 'frPlanerKlassenList',
                syncBtnId: 'frPlanerKlassenSyncSp',
                reloadBtnId: 'frPlanerKlassenReload',
                cleanBtnId: 'frPlanerKlassenClean'
            },
            siteCtx
        );
        wireFreistellungKategorienAdmin(
            {
                listId: 'frPlanerKatExtraList',
                addId: 'frPlanerKatExtraAdd',
                newInputId: 'frPlanerKatExtraNew',
                syncListBtnId: 'frPlanerKatSyncSp'
            },
            siteCtx
        );
    }

    wireFreistellungStudentKvSearch(root);

    function syncAntragFormFromDom() {
        state.form = { ...state.form, ...readFormFromDom(root) };
    }

    function validateCurrentAntragStep() {
        syncAntragFormFromDom();
        const step = Math.min(4, Math.max(1, Number(state.formStep) || 1));
        const f = state.form || {};
        if (step === 1) {
            if (!String(f.schuelerName || '').trim()) return 'Bitte den Namen angeben.';
            if (!String(f.klasse || '').trim()) return 'Bitte eine Klasse wählen.';
        } else if (step === 2) {
            if (!String(f.beginn || '').trim()) return 'Bitte Beginn angeben.';
            if (!String(f.ende || '').trim()) return 'Bitte Ende angeben.';
        } else if (step === 3) {
            if (!String(f.kategorie || '').trim()) return 'Bitte eine Kategorie wählen.';
        }
        return '';
    }

    root.querySelectorAll('[data-fr-step-next]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const err = validateCurrentAntragStep();
            if (err) {
                toast(err);
                return;
            }
            applyKvFromClass(state);
            state.formStep = Math.min(4, (Number(state.formStep) || 1) + 1);
            paint();
        });
    });
    root.querySelectorAll('[data-fr-step-prev]').forEach((btn) => {
        btn.addEventListener('click', () => {
            syncAntragFormFromDom();
            state.formStep = Math.max(1, (Number(state.formStep) || 1) - 1);
            paint();
        });
    });
    root.querySelectorAll('[data-fr-goto-step]').forEach((btn) => {
        btn.addEventListener('click', () => {
            syncAntragFormFromDom();
            const target = Math.min(4, Math.max(1, Number(btn.getAttribute('data-fr-goto-step') || '1')));
            const current = Math.min(4, Math.max(1, Number(state.formStep) || 1));
            // Vorwärts nur bis zum nächsten offenen Schritt (kein Überspringen ohne Eingabe)
            if (target > current) {
                const err = validateCurrentAntragStep();
                if (err) {
                    toast(err);
                    return;
                }
                state.formStep = Math.min(target, current + 1);
            } else {
                state.formStep = target;
            }
            applyKvFromClass(state);
            paint();
        });
    });

    const form = root.querySelector('#frAntragForm');
    if (form) {
        bindFreistellungUploadField(root);
        form.addEventListener('submit', (ev) => {
            ev.preventDefault();
            syncAntragFormFromDom();
            submitAntrag();
        });
        const klasse = root.querySelector('#frFormKlasse');
        if (klasse) {
            const onKlasseChange = () => {
                syncAntragFormFromDom();
                if (state.role === 'schueler' && state.form.klasse && !isStudentKlasseLocked(state)) {
                    state.demoKlasseCode = String(state.form.klasse).trim();
                    persistDemoKlasseCode(state.demoKlasseCode);
                }
                applyKvFromClass(state);
                paint();
            };
            klasse.addEventListener('change', onKlasseChange);
        }
        ['frFormBeginn', 'frFormEnde'].forEach((id) => {
            const el = root.querySelector('#' + id);
            if (el) {
                el.addEventListener('change', () => {
                    syncAntragFormFromDom();
                    paint();
                });
            }
        });
        const reset = root.querySelector('#frBtnFormReset');
        if (reset) {
            reset.addEventListener('click', () => {
                state.form = emptyForm({ schuelerName: state.accountName });
                state.formStep = 1;
                prefillStudentFreistellungForm(state);
                clearFreistellungUploadField(root);
                paint();
            });
        }
    }

    const filters = root.querySelector('[data-fr-filters]');
    if (filters) {
        const apply = () => {
            state.filters = readFiltersFromDom(root);
            paint();
        };
        const q = filters.querySelector('#frFilterQ');
        if (q) {
            q.addEventListener('input', () => {
                clearTimeout(q._frT);
                q._frT = setTimeout(apply, 200);
            });
            q.addEventListener('change', apply);
        }
        filters.querySelectorAll('select').forEach((el) => {
            el.addEventListener('change', apply);
        });
        const reset = root.querySelector('#frFilterReset');
        if (reset) {
            reset.addEventListener('click', () => {
                state.filterMultiOpen = { klasse: false, status: false, kategorie: false };
                state.filterMultiQuery = { klasse: '', status: '', kategorie: '' };
                state.filters = { klasse: '', status: '', kategorie: '', multiDay: '', q: '' };
                paint();
            });
        }
    }

    if (!state.filterMultiOpen || typeof state.filterMultiOpen !== 'object') {
        state.filterMultiOpen = { klasse: false, status: false, kategorie: false };
    }
    if (!state.filterMultiQuery || typeof state.filterMultiQuery !== 'object') {
        state.filterMultiQuery = { klasse: '', status: '', kategorie: '' };
    }
    root.querySelectorAll('[data-fr-filter-multi]').forEach((box) => {
        const key = String(box.getAttribute('data-fr-filter-multi') || '').trim();
        if (!key) return;
        box.addEventListener('toggle', () => {
            state.filterMultiOpen = { ...state.filterMultiOpen, [key]: !!box.open };
        });
        box.addEventListener('change', (ev) => {
            const t = ev.target;
            if (!t || t.type !== 'checkbox') return;
            state.filterMultiOpen = { ...state.filterMultiOpen, [key]: true };
            state.filters = readFiltersFromDom(root);
            paint();
        });
        const qEl = box.querySelector('[data-fr-filter-multi-q="' + key + '"]');
        if (qEl) {
            qEl.addEventListener('input', () => {
                const needle = String(qEl.value || '');
                state.filterMultiQuery = { ...state.filterMultiQuery, [key]: needle };
                const low = needle.trim().toLowerCase();
                box.querySelectorAll('.fr-filter-multi__item').forEach((row) => {
                    const hay = String(row.textContent || '')
                        .trim()
                        .toLowerCase();
                    row.classList.toggle('is-filtered-out', !!(low && hay.indexOf(low) === -1));
                });
            });
        }
        const clearBtn = box.querySelector('[data-fr-filter-multi-clear="' + key + '"]');
        if (clearBtn) {
            clearBtn.addEventListener('click', (ev) => {
                ev.preventDefault();
                state.filterMultiOpen = { ...state.filterMultiOpen, [key]: true };
                state.filters = { ...state.filters, [key]: '' };
                paint();
            });
        }
    });

    const calToday = root.querySelector('#frCalToday');
    if (calToday) {
        calToday.addEventListener('click', () => {
            alignCalendarToToday();
            state.view = 'kalender';
            paint();
        });
    }
    root.querySelectorAll('[data-fr-cal-mode]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const mode = String(btn.getAttribute('data-fr-cal-mode') || 'month');
            state.calMode = mode === 'agenda' || mode === 'week' || mode === 'month' ? mode : 'month';
            paint();
        });
    });
    root.querySelectorAll('[data-fr-cal-nav]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const delta = Number(btn.getAttribute('data-fr-cal-nav') || '0');
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
    // Legacy month buttons (falls noch im DOM)
    const calPrev = root.querySelector('#frCalPrev');
    const calNext = root.querySelector('#frCalNext');
    if (calPrev) {
        calPrev.addEventListener('click', () => {
            state.calMonth -= 1;
            if (state.calMonth < 1) {
                state.calMonth = 12;
                state.calYear -= 1;
            }
            paint();
        });
    }
    if (calNext) {
        calNext.addEventListener('click', () => {
            state.calMonth += 1;
            if (state.calMonth > 12) {
                state.calMonth = 1;
                state.calYear += 1;
            }
            paint();
        });
    }

    root.querySelectorAll('[data-fr-detail]').forEach((el) => {
        const open = () => {
            state.detailId = el.getAttribute('data-fr-detail');
            paint();
        };
        el.addEventListener('click', (ev) => {
            if (
                el.tagName === 'TR' &&
                ev.target instanceof Element &&
                ev.target.closest('button, a, input, select, textarea')
            ) {
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
    root.querySelectorAll('[data-fr-close-detail]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.detailId = null;
            paint();
        });
    });
    root.querySelectorAll('[data-fr-approve]').forEach((btn) => {
        btn.addEventListener('click', () => setStatus(btn.getAttribute('data-fr-approve'), 'Genehmigt'));
    });
    root.querySelectorAll('[data-fr-reject]').forEach((btn) => {
        btn.addEventListener('click', () => setStatus(btn.getAttribute('data-fr-reject'), 'Abgelehnt'));
    });
    root.querySelectorAll('[data-fr-delete]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-fr-delete');
            deleteAntrag(id).catch((e) => toast(String((e && e.message) || e)));
        });
    });
}

function boot() {
    refreshKategorieChoicesState();
    root = document.getElementById('frApp');
    if (!root) return;

    const params = new URLSearchParams(window.location.search || '');
    state.demoRoleOverride = params.get('demoRole') === '1';
    state.entraGroupsConfigured = entraGroupsConfigured(loadPermissionsConfig());
    state.demoKlasseCode = loadDemoKlasseCode();
    const roleQ = params.get('role');
    if (roleQ === 'direktion' || roleQ === 'kv' || roleQ === 'schueler') {
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

    syncAccount();
    const siteBoot = resolveFreistellungSiteUrl(state.siteUrl);
    if (siteBoot) state.siteUrl = siteBoot;
    if (state.form.klasse) {
        const kv = resolveKvForClass(state, state.form.klasse);
        if (kv) {
            state.form.kvEmail = kv.email;
            state.form.kvName = kv.name;
        }
    }

    resolvePlanerRole()
        .then(() => paint())
        .catch(() => paint());

    window.addEventListener('ms365-auth-state-changed', () => {
        void (async () => {
            try {
                await resolvePlanerRole();
                const url = await ensureFrSiteUrl();
                if (url && !state.planerAccessDenied) await refreshData();
                else paint();
            } catch (e) {
                if (!(state.localDemoOnly && (state.items || []).length)) {
                    state.error = friendlySpoError(e) || 'Laden fehlgeschlagen';
                }
                paint();
            }
        })();
    });
    window.addEventListener('ms365-tenant-settings-changed', () => {
        if (typeof window.ms365AuthIsLoggedIn !== 'function' || !window.ms365AuthIsLoggedIn()) return;
        const site = resolveFreistellungSiteUrl(state.siteUrl);
        if (site) state.siteUrl = site;
        resolvePlanerRole(state.listId)
            .then(() => {
                if (!state.planerAccessDenied && state.siteUrl) return refreshData();
                paint();
            })
            .catch(() => paint());
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

    void (async () => {
        const url = await ensureFrSiteUrl();
        if (url) {
            refreshData();
        } else {
            if (!useMinimalPlanerChrome(state)) {
                state.info =
                    state.roleHint ||
                    'SharePoint-Site fehlt in diesem Browser. IT: Tool-Site-URL in der Lizenz setzen – oder Demo laden.';
            }
            resolvePlanerRole().then(() => paint());
        }
    })();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}
