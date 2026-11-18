/**
 * SharePoint: Schulaktivitäten-Listen anlegen.
 */
import {
    LIST_TITLES,
    AKTIVITAETEN_COLUMNS,
    REGELWERK_COLUMNS,
    REQUIRED_COLUMNS,
    DEFAULT_REGELWERK_FIELDS,
    toGraphColumnBody
} from '../schulaktivitaeten-planer/schulaktivitaeten-planer-schema.js';
import {
    ensureSchoolStammdatenShellLists,
    healthSchoolStammdatenShellLists,
    resolveStammdatenSpoListTitlesFromWindow,
    seedSchoolStammdatenListsFromLocal
} from '../../shared/school-stammdaten-spo-lists.js';
import {
    applySchulaktivitaetenPackagePermissions,
    normalizeAktPermissionsConfig,
    loadAktPermissionsConfig,
    loadEffectiveAktPermissionsConfig,
    saveAktPermissionsConfig,
    aktEntraGroupsConfigured
} from '../schulaktivitaeten-planer/schulaktivitaeten-planer-permissions.js';
import {
    htmlAktSetupEntraPermGrid,
    fillAktPermissionsPickers,
    remountAktSetupPermPickers,
    persistAktPickersToStorage,
    readAktPermissionsFromPickers
} from '../schulaktivitaeten-planer/schulaktivitaeten-permissions-ui.js';
import {
    SPAKT_WIZARD_STEP_COUNT,
    SPAKT_WIZARD_STEP_LABELS,
    clampWizardStep,
    loadWizardStep,
    saveWizardStep,
    validateWizardStep,
    buildWizardSummary
} from './sharepoint-liste-schulaktivitaeten-wizard.js';
import {
    persistSiteUrl,
    loadSavedSiteUrl
} from '../schulaktivitaeten-planer/schulaktivitaeten-planer-state.js';

function rememberAktSiteUrl(webUrl) {
    const url = String(webUrl || '').trim();
    if (!url) return;
    persistSiteUrl(url);
}

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

function G() {
    const api = window.ms365SpoGraph;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

function $(id) {
    return document.getElementById(id);
}

function log(msg) {
    const el = $('spaktLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + msg;
    el.scrollTop = el.scrollHeight;
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

async function ensureToken() {
    return await G().getGraphToken(SCOPES);
}

async function findListByTitle(token, siteId, listTitle) {
    const title = String(listTitle || '').trim();
    const path =
        G().graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await G().graphJson('GET', path, token, undefined, 'v1.0');
    return ((data && data.value) || [])[0] || null;
}

function sameChoiceSet(a, b) {
    const left = (Array.isArray(a) ? a : []).map((x) => String(x)).sort();
    const right = (Array.isArray(b) ? b : []).map((x) => String(x)).sort();
    if (left.length !== right.length) return false;
    return left.every((v, i) => v === right[i]);
}

async function addMissingColumns(siteId, listId, token, defs, write) {
    const colsPath = G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns?$top=200';
    const colsData = await G().graphJson('GET', colsPath, token, undefined, 'v1.0');
    const existingCols = (colsData && colsData.value) || [];
    const byName = new Map(
        existingCols
            .map((c) => [String((c && c.name) || '').trim(), c])
            .filter((pair) => pair[0])
    );
    const base = G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
    let added = 0;
    let updated = 0;
    for (let i = 0; i < defs.length; i++) {
        const def = defs[i];
        const existing = byName.get(def.name);
        if (!existing) {
            await G().graphJson('POST', base, token, toGraphColumnBody(def), 'v1.0');
            added++;
            write('  + Spalte ' + def.name);
            await G().sleep(120);
            continue;
        }
        if (def.choice && existing.id) {
            const current = (existing.choice && existing.choice.choices) || [];
            const wanted = def.choice.choices || [];
            if (!sameChoiceSet(current, wanted)) {
                await G().graphJson(
                    'PATCH',
                    base + '/' + encodeURIComponent(existing.id),
                    token,
                    { choice: { allowTextEntry: false, choices: [...wanted] } },
                    'v1.0'
                );
                updated++;
                write('  ~ Choice-Spalte ' + def.name + ' aktualisiert');
                await G().sleep(120);
            }
        }
    }
    return { added, updated };
}

async function ensureList(token, siteId, title, columnDefs, write) {
    let list = await findListByTitle(token, siteId, title);
    if (!list || !list.id) {
        write('Erstelle Liste „' + title + '" …');
        const created = await G().graphJson(
            'POST',
            G().graphPathSite(siteId) + '/lists',
            token,
            { displayName: title, list: { template: 'genericList' } },
            'v1.0'
        );
        const listId = created && created.id ? String(created.id) : '';
        if (!listId) throw new Error('Listen-ID fehlt für „' + title + '“.');
        list = { id: listId, displayName: title, webUrl: created.webUrl || '' };
        write('Liste angelegt: ' + (list.webUrl || listId));
    } else {
        write('Liste „' + title + '" bereits vorhanden.');
    }
    write('Prüfe Spalten für „' + title + '" …');
    const result = await addMissingColumns(siteId, list.id, token, columnDefs, write);
    const parts = [];
    if (result.added) parts.push(result.added + ' Spalte(n) ergänzt');
    if (result.updated) parts.push(result.updated + ' Choice-Spalte(n) aktualisiert');
    write(parts.length ? parts.join(', ') + '.' : 'Spalten vollständig.');
    return list;
}

async function seedRegelwerkIfEmpty(token, siteId, listId, write) {
    const itemsPath =
        G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/items?$expand=fields&$top=5';
    const data = await G().graphJson('GET', itemsPath, token, undefined, 'v1.0');
    const items = (data && data.value) || [];
    if (items.length) {
        write('Regelwerk enthält bereits Einträge – Seed übersprungen.');
        return false;
    }
    write('Lege Standard-Regelwerk an …');
    await G().graphJson(
        'POST',
        G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/items',
        token,
        { fields: { ...DEFAULT_REGELWERK_FIELDS } },
        'v1.0'
    );
    write('Seed: „' + DEFAULT_REGELWERK_FIELDS.Title + '"');
    return true;
}

function readStammdatenListTitles() {
    return resolveStammdatenSpoListTitlesFromWindow();
}

/**
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 * @param {object} [permOpts] normalizeAktPermissionsConfig-kompatibel
 */
export async function createSchulaktivitaetenLists(webUrl, logFn, permOpts) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');
    const o = normalizeAktPermissionsConfig({
        ...loadAktPermissionsConfig(),
        ...(permOpts || {})
    });

    const token = await ensureToken();
    write('Löse Website auf …');
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    write('Site: ' + (site.displayName || siteId));

    write('— Stammdaten-Listen (Lehrer / Klassen inkl. KV) —');
    const titles = readStammdatenListTitles();
    let stammdaten = null;
    let stammdatenSeed = null;
    try {
        stammdatenSeed = await seedSchoolStammdatenListsFromLocal({
            token,
            siteId,
            siteWebUrl: url,
            graph: G(),
            write,
            titles,
            personFields: true,
            removeOrphans: false
        });
        stammdaten = stammdatenSeed && stammdatenSeed.shell ? stammdatenSeed.shell : null;
    } catch (e) {
        write('Hinweis Stammdaten-Befüllung: ' + (e && e.message ? e.message : e));
        stammdaten = await ensureSchoolStammdatenShellLists({
            token,
            siteId,
            graph: G(),
            write,
            titles
        });
    }

    write('— Schulaktivitäten-Paket —');
    const regelwerk = await ensureList(token, siteId, LIST_TITLES.regelwerk, REGELWERK_COLUMNS, write);
    await seedRegelwerkIfEmpty(token, siteId, regelwerk.id, write);
    const aktivitaeten = await ensureList(
        token,
        siteId,
        LIST_TITLES.aktivitaeten,
        AKTIVITAETEN_COLUMNS,
        write
    );

    if (!o.skipPerms) {
        try {
            await applySchulaktivitaetenPackagePermissions(url, o, write);
        } catch (e) {
            write('Hinweis Berechtigungen: ' + (e && e.message ? e.message : e));
        }
    } else {
        write('Berechtigungen übersprungen (Haken gesetzt).');
    }

    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
        window.ms365ActionLog.append({
            tool: 'sharepoint',
            action: 'create-schulaktivitaeten-lists',
            target: url,
            summary:
                'Schulaktivitäten-Paket (Lehrer/Klassen befüllt, Regelwerk, Schulaktivitaeten)' +
                (o.skipPerms ? ' ohne Rechte' : '')
        });
    }

    return {
        siteId,
        stammdatenSeed,
        lists: {
            regelwerk: { id: regelwerk.id, webUrl: regelwerk.webUrl || '' },
            aktivitaeten: { id: aktivitaeten.id, webUrl: aktivitaeten.webUrl || '' },
            stammdaten: (stammdaten && stammdaten.lists) || {}
        }
    };
}

/**
 * Nur Lehrer-/Klassenlisten aus lokalen Stammdaten befüllen (ohne AKT-Paket/Rechte).
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 */
export async function seedSchulaktivitaetenStammdatenLists(webUrl, logFn) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');
    const token = await ensureToken();
    write('Löse Website auf …');
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    write('Site: ' + (site.displayName || siteId));
    return await seedSchoolStammdatenListsFromLocal({
        token,
        siteId,
        siteWebUrl: url,
        graph: G(),
        write,
        titles: readStammdatenListTitles(),
        personFields: true,
        removeOrphans: false
    });
}

export async function probeSchulaktivitaetenSite(webUrl) {
    const token = await ensureToken();
    const site = await G().resolveSiteFromWebUrl(token, String(webUrl || '').trim());
    return site;
}

export async function healthSchulaktivitaetenLists(webUrl) {
    const token = await ensureToken();
    const site = await G().resolveSiteFromWebUrl(token, String(webUrl || '').trim());
    const siteId = site && site.id ? String(site.id) : '';
    const report = [];
    const stammTitles = readStammdatenListTitles();
    const stammHealth = await healthSchoolStammdatenShellLists({
        token,
        siteId,
        graph: G(),
        titles: stammTitles
    });
    report.push(...stammHealth);
    for (const [key, title] of Object.entries(LIST_TITLES)) {
        const list = await findListByTitle(token, siteId, title);
        if (!list) {
            report.push({ title, ok: false, detail: 'fehlt', key });
            continue;
        }
        const colsPath =
            G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(list.id) + '/columns?$top=200';
        const colsData = await G().graphJson('GET', colsPath, token, undefined, 'v1.0');
        const names = new Set(((colsData && colsData.value) || []).map((c) => c.name));
        const required = REQUIRED_COLUMNS[title] || [];
        const missing = required.filter((n) => !names.has(n));
        report.push({
            title,
            ok: missing.length === 0,
            detail: missing.length ? 'fehlt: ' + missing.join(', ') : 'ok',
            key
        });
    }
    return report;
}

function readPermissionsFromForm() {
    return normalizeAktPermissionsConfig({
        ...readAktPermissionsFromPickers(),
        skipPerms: !!($('spaktSkipPerms') && $('spaktSkipPerms').checked)
    });
}

function openProtocol() {
    const panel = $('spaktLogPanel');
    if (panel) panel.open = true;
}

function clearProtocol() {
    if ($('spaktLog')) $('spaktLog').textContent = '';
    openProtocol();
}

function mountSetupEntraPermHosts() {
    const cfg = loadEffectiveAktPermissionsConfig();
    const admin = $('spaktEntraPermAdmin');
    if (admin && admin.dataset.aktPermMounted !== '1') {
        admin.innerHTML = htmlAktSetupEntraPermGrid(cfg, 'admin');
        admin.dataset.aktPermMounted = '1';
    }
    const ein = $('spaktEntraPermEinreicher');
    if (ein && ein.dataset.aktPermMounted !== '1') {
        ein.innerHTML = htmlAktSetupEntraPermGrid(cfg, 'einreicher');
        ein.dataset.aktPermMounted = '1';
    }
}

function fillPermissionsForm(cfg) {
    mountSetupEntraPermHosts();
    remountAktSetupPermPickers(() =>
        persistAktPickersToStorage($('spaktSkipPerms') && $('spaktSkipPerms').checked)
    );
    fillAktPermissionsPickers(cfg || loadEffectiveAktPermissionsConfig());
    const c = normalizeAktPermissionsConfig(cfg || loadEffectiveAktPermissionsConfig());
    if ($('spaktSkipPerms')) $('spaktSkipPerms').checked = c.skipPerms;
}

function wizardContextFromDom() {
    const urlEl = $('spaktSiteUrl');
    const perms = readPermissionsFromForm();
    const hasAdmin = !!(
        perms.groupAdmin ||
        perms.groupAdminId ||
        (perms.adminUsers && perms.adminUsers.length)
    );
    const hasEinreicher = !!(
        perms.groupLehrer ||
        perms.groupLehrerId ||
        (perms.lehrerUsers && perms.lehrerUsers.length)
    );
    return {
        siteUrl: String((urlEl && urlEl.value) || '').trim(),
        skipPerms: !!perms.skipPerms,
        hasAdmin,
        hasEinreicher,
        adminLabel: String(perms.groupAdmin || '').trim(),
        einreicherLabel: String(perms.groupLehrer || '').trim(),
        adminUsersCount: (perms.adminUsers || []).length,
        einreicherUsersCount: (perms.lehrerUsers || []).length
    };
}

function refreshWizardSummary() {
    const box = $('spaktWizardSummary');
    if (!box) return;
    const sum = buildWizardSummary(wizardContextFromDom());
    box.innerHTML =
        '<li><strong>Website:</strong> ' +
        sum.site +
        '</li>' +
        '<li><strong>Genehmigung:</strong> ' +
        sum.admin +
        '</li>' +
        '<li><strong>Einreicher:</strong> ' +
        sum.einreicher +
        '</li>' +
        '<li><strong>Listen:</strong> ' +
        sum.listsText +
        '</li>' +
        '<li><strong>Rechte:</strong> ' +
        (sum.skipPerms ? 'diesmal überspringen' : 'mit anwenden') +
        '</li>';
}

function showWizardStep(n) {
    const step = clampWizardStep(n);
    saveWizardStep(step);
    for (let i = 1; i <= SPAKT_WIZARD_STEP_COUNT; i++) {
        const panel = $('spaktWizardStep' + i);
        if (panel) panel.hidden = i !== step;
    }
    document.querySelectorAll('#spaktWizardGlance [data-spakt-wizard-step]').forEach((btn) => {
        const s = parseInt(btn.getAttribute('data-spakt-wizard-step') || '0', 10);
        const active = s === step;
        btn.classList.toggle('is-active', active);
        btn.setAttribute('aria-selected', active ? 'true' : 'false');
    });
    const hint = $('spaktPhaseHint');
    if (hint) {
        hint.textContent =
            'Schritt ' + step + ' von ' + SPAKT_WIZARD_STEP_COUNT + ': ' + SPAKT_WIZARD_STEP_LABELS[step - 1];
    }
    const back = $('spaktWizardBack');
    const next = $('spaktWizardNext');
    if (back) back.disabled = step <= 1;
    if (next) {
        next.hidden = step >= SPAKT_WIZARD_STEP_COUNT;
        if (step < SPAKT_WIZARD_STEP_COUNT) {
            const label = step === SPAKT_WIZARD_STEP_COUNT - 1 ? 'Zur Übersicht' : 'Weiter';
            next.innerHTML = label + ' <i class="bi bi-arrow-right"></i>';
        }
    }
    if (step === SPAKT_WIZARD_STEP_COUNT) refreshWizardSummary();
    if (step === 2 || step === 3) {
        fillPermissionsForm(loadEffectiveAktPermissionsConfig());
    }
}

function tryAdvanceWizard(fromStep) {
    persistAktPickersToStorage($('spaktSkipPerms') && $('spaktSkipPerms').checked);
    const err = validateWizardStep(fromStep, wizardContextFromDom());
    if (err) {
        toast(err);
        return false;
    }
    return true;
}

function wireWizard() {
    document.querySelectorAll('#spaktWizardGlance [data-spakt-wizard-step]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const target = parseInt(btn.getAttribute('data-spakt-wizard-step') || '1', 10);
            const cur = loadWizardStep();
            if (target > cur) {
                for (let s = cur; s < target; s++) {
                    if (!tryAdvanceWizard(s)) {
                        showWizardStep(s);
                        return;
                    }
                }
            }
            showWizardStep(target);
        });
    });
    const back = $('spaktWizardBack');
    const next = $('spaktWizardNext');
    if (back) {
        back.addEventListener('click', () => showWizardStep(loadWizardStep() - 1));
    }
    if (next) {
        next.addEventListener('click', () => {
            const cur = loadWizardStep();
            if (!tryAdvanceWizard(cur)) return;
            showWizardStep(cur + 1);
        });
    }
    showWizardStep(loadWizardStep());
}

document.addEventListener('DOMContentLoaded', () => {
    const urlEl = $('spaktSiteUrl');
    if (urlEl && !String(urlEl.value || '').trim()) {
        try {
            const saved = loadSavedSiteUrl();
            if (saved) urlEl.value = saved;
        } catch {
            /* ignore */
        }
    }
    if (urlEl && !String(urlEl.value || '').trim()) {
        try {
            const setup =
                window.ms365AppDataV2 && window.ms365AppDataV2.getSetup
                    ? window.ms365AppDataV2.getSetup()
                    : null;
            if (setup && setup.intranetSiteUrl) urlEl.value = String(setup.intranetSiteUrl).trim();
        } catch {
            /* ignore */
        }
    }
    if (urlEl) {
        const current = String(urlEl.value || '').trim();
        if (current) rememberAktSiteUrl(current);
        urlEl.addEventListener('change', () => {
            const webUrl = String(urlEl.value || '').trim();
            if (webUrl) rememberAktSiteUrl(webUrl);
        });
    }

    fillPermissionsForm(loadEffectiveAktPermissionsConfig());
    wireWizard();

    const skip = $('spaktSkipPerms');
    if (skip) {
        skip.addEventListener('change', () => {
            persistAktPickersToStorage(skip.checked);
            refreshWizardSummary();
        });
    }

    const run = $('spaktBtnRun');
    if (run) {
        run.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            rememberAktSiteUrl(webUrl);
            clearProtocol();
            const perms = readPermissionsFromForm();
            saveAktPermissionsConfig(perms);
            if (!perms.skipPerms && !aktEntraGroupsConfigured(perms)) {
                if (
                    !window.confirm(
                        'Keine Genehmigung- oder Einreicher-Gruppe/Person gewählt.\n\nPaket trotzdem ohne Listenrechte anlegen? (Site-Standardrechte bleiben.)'
                    )
                ) {
                    return;
                }
            }
            run.disabled = true;
            createSchulaktivitaetenLists(webUrl, log, perms)
                .then(() => {
                    rememberAktSiteUrl(webUrl);
                    toast(
                        'Paket angelegt bzw. geprüft (Listen' +
                            (perms.skipPerms ? '' : ' + Rechte') +
                            '). Site-URL für den Planer gespeichert – für Lehrer-Browser zusätzlich Admin → Schule „Tool-Site-URL“ setzen.'
                    );
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                })
                .finally(() => {
                    run.disabled = false;
                });
        });
    }

    const seedOnly = $('spaktBtnSeed');
    if (seedOnly) {
        seedOnly.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            if (!webUrl) {
                toast('Website-URL fehlt.');
                return;
            }
            rememberAktSiteUrl(webUrl);
            clearProtocol();
            seedOnly.disabled = true;
            seedSchulaktivitaetenStammdatenLists(webUrl, log)
                .then((r) => {
                    if (r && r.skipped) {
                        toast('Keine lokalen Stammdaten gefunden – zuerst Klassen/Lehrer pflegen.');
                    } else {
                        toast(
                            'Stammdaten geschrieben (' +
                                (r && r.teachers != null ? r.teachers : '?') +
                                ' Lehrkräfte, ' +
                                (r && r.classes != null ? r.classes : '?') +
                                ' Klassen).'
                        );
                    }
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                })
                .finally(() => {
                    seedOnly.disabled = false;
                });
        });
    }

    const permsOnly = $('spaktBtnPerms');
    if (permsOnly) {
        permsOnly.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            if (!webUrl) {
                toast('Website-URL fehlt.');
                return;
            }
            rememberAktSiteUrl(webUrl);
            clearProtocol();
            const perms = readPermissionsFromForm();
            if (perms.skipPerms) {
                toast('„Rechte überspringen“ ist aktiv – Haken entfernen.');
                return;
            }
            if (!aktEntraGroupsConfigured(perms)) {
                toast('Bitte Genehmigung und/oder Einreicher (Gruppe oder Person) wählen.');
                return;
            }
            saveAktPermissionsConfig(perms);
            permsOnly.disabled = true;
            applySchulaktivitaetenPackagePermissions(webUrl, perms, log)
                .then(async () => {
                    toast('Berechtigungen angewendet.');
                    try {
                        const report = await healthSchulaktivitaetenLists(webUrl);
                        const empty = (report || []).filter(
                            (r) =>
                                (r.key === 'klassen' || r.key === 'lehrer') &&
                                (r.itemCount === 0 || !r.ok)
                        );
                        if (empty.length) {
                            log(
                                'Hinweis: Stammdaten-Listen ohne nutzbare Zeilen – „Nur Stammdaten befüllen“ ausführen, sonst bleibt die Klassenauswahl im Planer leer.'
                            );
                            empty.forEach((r) => log('  · ' + r.title + ': ' + r.detail));
                            toast(
                                'Rechte ok – aber Klassen/Lehrer noch leer. Bitte „Nur Stammdaten befüllen“.'
                            );
                        }
                    } catch {
                        /* ignore */
                    }
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                })
                .finally(() => {
                    permsOnly.disabled = false;
                });
        });
    }

    const probe = $('spaktBtnProbe');
    if (probe) {
        probe.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            if (!webUrl) {
                toast('Website-URL fehlt.');
                return;
            }
            rememberAktSiteUrl(webUrl);
            probeSchulaktivitaetenSite(webUrl)
                .then((site) => toast('Site OK – URL für Planer gespeichert: ' + (site.displayName || site.id)))
                .catch((e) => toast('Fehler: ' + (e && e.message ? e.message : e)));
        });
    }

    const health = $('spaktBtnHealth');
    if (health) {
        health.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            if (webUrl) rememberAktSiteUrl(webUrl);
            const box = $('spaktHealth');
            clearProtocol();
            healthSchulaktivitaetenLists(webUrl)
                .then((report) => {
                    if (box) {
                        box.hidden = false;
                        box.innerHTML = report
                            .map(
                                (r) =>
                                    `<div>${r.ok ? '✓' : '✗'} <strong>${r.title}</strong> – ${r.detail}</div>`
                            )
                            .join('');
                    }
                    (report || []).forEach((r) =>
                        log((r.ok ? '✓ ' : '✗ ') + r.title + ' – ' + r.detail)
                    );
                })
                .catch((e) => toast('Fehler: ' + (e && e.message ? e.message : e)));
        });
    }
});
