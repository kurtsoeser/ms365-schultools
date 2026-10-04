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
    matchStudentByEmail,
    studentKlasseFromRecord,
    matchKvByClassHeadEmail,
    resolveStudentKlasseCode,
    persistDemoKlasseCode,
    loadDemoKlasseCode,
    viewsForRole,
    useStudentPlanerChrome,
    useKvPlanerChrome,
    useDirektionPlanerChrome,
    useMinimalPlanerChrome,
    isLikelyFreistellungStaffAccount
} from './freistellung-planer-state.js';
import {
    applyPlanerRoleFromEntra,
    isPlanerDemoRoleUiEnabled,
    finalizePlanerRoles
} from './freistellung-planer-entra-role.js';
import {
    accountIsDirektionPlannerUser,
    accountIsPlannerUserInList
} from './freistellung-planer-direktion-users.js';
import { entraGroupsConfigured, loadPermissionsConfig } from './freistellung-planer-permissions.js';
import { syncPlannerPermissionsFromSite } from './freistellung-planer-remote-config.js';
import {
    mergeKategorieChoices,
    loadExtraKategorien
} from './freistellung-planer-kategorien.js';
import {
    readPendingUploadFiles,
    uploadFreistellungNachweise,
    saveNachweiseOnItem
} from './freistellung-planer-nachweise.js';
import { patchFreistellungKategorieColumn } from './freistellung-planer-graph.js';
import { wireFreistellungKategorienAdmin } from './freistellung-kategorien-ui.js';
import {
    resolveFrContext,
    loadAllFreistellungen,
    createFreistellungItem,
    updateFreistellungStatus,
    probeFreistellungListRead
} from './freistellung-planer-graph.js';
import { LIST_TITLE_DEFAULT } from './freistellung-planer-schema.js';
import { validateFreistellung } from './freistellung-planer-logic.js';
import {
    renderApp,
    readFormFromDom,
    readFiltersFromDom,
    bindFreistellungUploadField,
    clearFreistellungUploadField
} from './freistellung-planer-ui.js';
import { downloadFreistellungCsv } from './freistellung-planer-export.js';
import {
    buildLocalDemoItems,
    getDemoSeedPackage,
    parseDemoImportJson,
    itemsFromDemoPack,
    DEMO_SITE_DEFAULT,
    DEMO_SEED_TAG
} from './freistellung-planer-demo.js';
import {
    applyDemoStammdatenLocal,
    seedDemoFreistellungen,
    resetDemoFreistellungen
} from './freistellung-planer-demo-seed.js';

const state = createInitialState();
let root = null;

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else window.alert(msg);
}

function placePlanerAuthWidget() {
    const slot = document.getElementById('frNavAuthSlot');
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

function paint() {
    if (!root) return;
    if (state.role === 'schueler') prefillStudentFreistellungForm(state);
    syncPageHeader();
    renderApp(state, root);
    bindStatic();
    placePlanerAuthWidget();
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
    if (studentKlasseFromRecord(state.studentMatch)) {
        state.demoKlasseCode = '';
        persistDemoKlasseCode('');
    }
    state.kvMatch = matchKvByClassHeadEmail(state.accountEmail, state.stammdaten.classes);
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
    const site = String(state.siteUrl || setup.siteUrl || loadSavedSiteUrl(setup) || '').trim();
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
    if (!state.planerAccessDenied || !state.siteUrl) return;

    syncAccount();
    await ensurePlannerPermissionsConfig(state.listId);
    await applyPlanerRoleFromEntra(state, {
        demoRoleOverride: state.demoRoleOverride,
        preferredDemoRole: state.demoRoleOverride ? state.role : null,
        preferredActiveRole: state.role
    });
    if (!state.planerAccessDenied) return;

    const cfg = loadPermissionsConfig();

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

async function resolvePlanerRole(listIdHint) {
    syncAccount();
    await ensurePlannerPermissionsConfig(listIdHint);
    await applyPlanerRoleFromEntra(state, {
        demoRoleOverride: state.demoRoleOverride,
        preferredDemoRole: state.demoRoleOverride ? state.role : null,
        preferredActiveRole: state.role
    });
    await tryRecoverPlanerAccess();
    const allowed = viewsForRole(state.role, state).map((v) => v.id);
    if (!allowed.includes(state.view)) state.view = allowed[0] || 'meine';
}

async function refreshData() {
    state.loading = true;
    state.error = '';
    state.info = '';
    paint();
    try {
        syncAccount();
        if (!String(state.siteUrl || '').trim()) {
            const setup = loadSetupCfg();
            const fallback = loadSavedSiteUrl(setup);
            if (fallback) state.siteUrl = fallback;
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
            } catch {
                /* Liste erst nach Rollen-Sync / ohne Rechte */
            }
        }

        await resolvePlanerRole(state.listId);
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
            state.items = [];
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
            const allowed = viewsForRole(state.role, state).map((v) => v.id);
            if (!allowed.includes(state.view)) state.view = allowed[0] || 'meine';
            if (state.planerAccessDenied) {
                state.ctx = null;
                state.items = [];
                return;
            }
        }
        state.ctx = ctx;
        state.localDemoOnly = false;
        state.items = await loadAllFreistellungen(ctx);
        if (!useMinimalPlanerChrome(state)) {
            state.info =
                'Liste „' +
                (ctx.list.name || state.listName) +
                '“ · ' +
                state.items.length +
                ' Einträge';
        }
    } catch (e) {
        state.error = friendlySpoError(e) || 'Laden fehlgeschlagen';
        state.ctx = null;
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
    state.localDemoOnly = true;
    state.ctx = null;
    state.error = '';
    const usePackRows = pack && Array.isArray(pack.freistellungen) && pack.freistellungen.length;
    state.items = usePackRows
        ? itemsFromDemoPack(pack, {
              accountEmail: state.accountEmail,
              accountName: state.accountName
          })
        : buildLocalDemoItems({
              accountEmail: state.accountEmail,
              accountName: state.accountName,
              classes: state.stammdaten.classes
          });
    state.demoSeedTag = (pack && pack.seedTag) || DEMO_SEED_TAG;
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
        state.error = String((e && e.message) || e);
        paint();
        toast('SharePoint-Seed fehlgeschlagen – lokale Ansicht bleibt: ' + state.error);
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
                authorEmail: state.accountEmail,
                beantragtVon: state.accountEmail,
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
        const created = await createFreistellungItem(state.ctx, state.form);
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
        await refreshData();
        state.view = state.role === 'schueler' ? 'meine' : 'liste';
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
        await updateFreistellungStatus(state.ctx, itemId, status);
        toast('Status aktualisiert: ' + status);
        await refreshData();
    } catch (e) {
        toast(String((e && e.message) || e));
    }
}

function bindStatic() {
    root.querySelectorAll('[data-fr-view]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-fr-view') || 'dashboard';
            if (state.view === 'antrag' && state.role === 'schueler') prefillStudentFreistellungForm(state);
            paint();
        });
    });
    root.querySelectorAll('[data-fr-view-jump]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-fr-view-jump') || 'dashboard';
            if (state.view === 'antrag' && state.role === 'schueler') prefillStudentFreistellungForm(state);
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
            if (state.view === 'antrag' && role === 'schueler') prefillStudentFreistellungForm(state);
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
        wireFreistellungKategorienAdmin(
            {
                listId: 'frPlanerKatExtraList',
                addId: 'frPlanerKatExtraAdd',
                newInputId: 'frPlanerKatExtraNew',
                syncListBtnId: 'frPlanerKatSyncSp'
            },
            () => ({
                siteUrl: state.siteUrl,
                listId: state.listId || (loadSetupCfg() && loadSetupCfg().listId) || ''
            })
        );
    }

    const form = root.querySelector('#frAntragForm');
    if (form) {
        bindFreistellungUploadField(root);
        form.addEventListener('submit', (ev) => {
            ev.preventDefault();
            submitAntrag();
        });
        const klasse = root.querySelector('#frFormKlasse');
        if (klasse) {
            klasse.addEventListener('change', () => {
                state.form = { ...state.form, ...readFormFromDom(root) };
                if (state.role === 'schueler' && !state.studentMatch && state.form.klasse) {
                    state.demoKlasseCode = String(state.form.klasse).trim();
                    persistDemoKlasseCode(state.demoKlasseCode);
                }
                applyKvFromClass(state);
                paint();
            });
        }
        ['frFormBeginn', 'frFormEnde'].forEach((id) => {
            const el = root.querySelector('#' + id);
            if (el) {
                el.addEventListener('change', () => {
                    state.form = { ...state.form, ...readFormFromDom(root) };
                    paint();
                });
            }
        });
        const reset = root.querySelector('#frBtnFormReset');
        if (reset) {
            reset.addEventListener('click', () => {
                state.form = emptyForm({ schuelerName: state.accountName });
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
        filters.querySelectorAll('select, input').forEach((el) => {
            el.addEventListener('change', apply);
            if (el.tagName === 'INPUT') {
                el.addEventListener('input', () => {
                    clearTimeout(el._frT);
                    el._frT = setTimeout(apply, 200);
                });
            }
        });
        const reset = root.querySelector('#frFilterReset');
        if (reset) {
            reset.addEventListener('click', () => {
                state.filters = { klasse: '', status: '', kategorie: '', multiDay: '', q: '' };
                paint();
            });
        }
    }

    root.querySelectorAll('[data-fr-detail]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.detailId = btn.getAttribute('data-fr-detail');
            paint();
        });
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
    const setup = loadSetupCfg();
    if (!state.siteUrl && setup.siteUrl) state.siteUrl = setup.siteUrl;
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

    window.addEventListener('ms365-auth-widget-ready', placePlanerAuthWidget);
    window.addEventListener('ms365-auth-state-changed', () => {
        resolvePlanerRole()
            .then(() => {
                if (state.ctx && state.siteUrl) return refreshData();
                paint();
            })
            .catch((e) => {
                state.error = String((e && e.message) || e);
                paint();
            });
    });

    if (state.siteUrl) {
        refreshData();
    } else {
        if (!useMinimalPlanerChrome(state)) {
            state.info =
                'Site-URL aus Freistellungen-Setup übernehmen oder eintragen, dann „Laden“. Alternativ Demo nutzen.';
        }
        resolvePlanerRole().then(() => paint());
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}
