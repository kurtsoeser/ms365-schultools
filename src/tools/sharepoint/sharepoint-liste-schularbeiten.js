/**
 * SharePoint: Schularbeiten-Listen anlegen (Regelwerk, Terminfenster, Schularbeiten).
 */
import {
    LIST_TITLES,
    LIST_KEYS,
    LIST_DESCRIPTIONS,
    REGELWERK_COLUMNS,
    TERMINFENSTER_COLUMNS,
    SCHULARBEITEN_COLUMNS,
    FACHMETA_COLUMNS,
    REQUIRED_COLUMNS_BY_KEY,
    DEFAULT_REGELWERK_FIELDS,
    toGraphColumnBody,
    titlesForListKey
} from '../schularbeiten-planer/schularbeiten-planer-schema.js';
import {
    ensureCanonicalListTitle,
    resolvePlanerList
} from '../schularbeiten-planer/schularbeiten-planer-lists.js';
import { findListByDisplayName } from '../schularbeiten-planer/schularbeiten-planer-graph.js';
import { currentSchoolYearFromDate } from '../schularbeiten-planer/schularbeiten-planer-schuljahr.js';
import {
    applySchularbeitenPackagePermissions,
    loadPermissionsConfig,
    savePermissionsConfig,
    normalizePermissionsConfig
} from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import { entraGroupsConfigured } from '../schularbeiten-planer/schularbeiten-planer-entra-role.js';
import { publishPlannerPermissionsToSite } from '../schularbeiten-planer/schularbeiten-planer-remote-config.js';
import {
    SETUP_GROUP_FIELDS,
    readPermissionsFromPickers,
    fillPermissionsPickers,
    wirePermissionGroupPickers,
    persistPickersToStorage,
    htmlSchularbeitenEntraPermGrid,
    initSchularbeitenSetupExtraUsers
} from '../schularbeiten-planer/schularbeiten-permissions-ui.js';

const PLANER_SITE_STORAGE_KEY = 'ms365-sa-site-url';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/Group.Read.All'
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
    const el = $('spsaLog');
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
    return findListByDisplayName(token, siteId, listTitle);
}

async function addMissingColumns(siteId, listId, token, defs, write) {
    const colsPath = G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns?$top=200';
    const colsData = await G().graphJson('GET', colsPath, token, undefined, 'v1.0');
    const existing = new Set(
        ((colsData && colsData.value) || [])
            .map((c) => String((c && c.name) || '').trim())
            .filter(Boolean)
    );
    const base = G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
    let added = 0;
    for (let i = 0; i < defs.length; i++) {
        const def = defs[i];
        if (existing.has(def.name)) continue;
        await G().graphJson('POST', base, token, toGraphColumnBody(def), 'v1.0');
        added++;
        write('  + Spalte ' + def.name);
        await G().sleep(120);
    }
    return added;
}

async function ensureList(token, siteId, title, columnDefs, write, description) {
    let list = await findListByTitle(token, siteId, title);
    if (!list || !list.id) {
        write('Erstelle Liste „' + title + '" …');
        const body = {
            displayName: title,
            list: { template: 'genericList' }
        };
        if (description) body.description = String(description);
        const created = await G().graphJson(
            'POST',
            G().graphPathSite(siteId) + '/lists',
            token,
            body,
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
    const added = await addMissingColumns(siteId, list.id, token, columnDefs, write);
    write(added ? added + ' Spalte(n) ergänzt.' : 'Spalten vollständig.');
    return list;
}

async function seedRegelwerkIfEmpty(token, siteId, listId, write) {
    const itemsPath =
        G().graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(listId) +
        '/items?$expand=fields&$top=5';
    const data = await G().graphJson('GET', itemsPath, token, undefined, 'v1.0');
    const items = (data && data.value) || [];
    if (items.length) {
        write('Regelwerk enthält bereits ' + items.length + ' Eintrag/Einträge – Seed übersprungen.');
        return false;
    }
    write('Lege Standard-Regelwerk an …');
    const seedFields = {
        ...DEFAULT_REGELWERK_FIELDS,
        Schuljahr: currentSchoolYearFromDate()
    };
    await G().graphJson(
        'POST',
        G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/items',
        token,
        { fields: seedFields },
        'v1.0'
    );
    write('Seed: „' + DEFAULT_REGELWERK_FIELDS.Title + '"');
    return true;
}

/**
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 * @param {{ applyPermissions?: boolean, skipPerms?: boolean, groupAdmin?: string, groupLehrer?: string, groupSchueler?: string }} [opts]
 */
async function createSchularbeitenLists(webUrl, logFn, opts) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');
    const o = normalizePermissionsConfig({ ...loadPermissionsConfig(), ...(opts || {}) });
    if (opts && opts.applyPermissions === false) o.skipPerms = true;
    savePermissionsConfig(o);

    const token = await ensureToken();
    write('Löse Website auf …');
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
    write('Site: ' + (site.displayName || siteId));

    for (let i = 0; i < LIST_KEYS.length; i++) {
        const key = LIST_KEYS[i];
        await ensureCanonicalListTitle(G(), token, siteId, key, findListByDisplayName, write);
    }

    const regelwerk = await ensureList(
        token,
        siteId,
        LIST_TITLES.regelwerk,
        REGELWERK_COLUMNS,
        write,
        LIST_DESCRIPTIONS.regelwerk
    );
    await seedRegelwerkIfEmpty(token, siteId, regelwerk.id, write);

    const terminfenster = await ensureList(
        token,
        siteId,
        LIST_TITLES.terminfenster,
        TERMINFENSTER_COLUMNS,
        write,
        LIST_DESCRIPTIONS.terminfenster
    );

    const schularbeiten = await ensureList(
        token,
        siteId,
        LIST_TITLES.schularbeiten,
        SCHULARBEITEN_COLUMNS,
        write,
        LIST_DESCRIPTIONS.schularbeiten
    );

    const fachMeta = await ensureList(
        token,
        siteId,
        LIST_TITLES.fachMeta,
        FACHMETA_COLUMNS,
        write,
        LIST_DESCRIPTIONS.fachMeta
    );

    if (!o.skipPerms) {
        try {
            await applySchularbeitenPackagePermissions(url, o, write);
        } catch (e) {
            write('Hinweis Berechtigungen: ' + (e && e.message ? e.message : e));
        }
    } else {
        write('Berechtigungen übersprungen (Haken gesetzt).');
    }

    if (!o.skipPerms && entraGroupsConfigured(o)) {
        try {
            await publishPlannerPermissionsToSite(url, schularbeiten.id, o);
            write(
                'Planer-Gruppen auf der Site gespeichert (' +
                    'Site Assets/ms365/schularbeiten-planer-groups.json) – Lehrkräfte laden diese beim Start.'
            );
        } catch (e) {
            write(
                'Hinweis Planer-Konfiguration auf Site: ' + (e && e.message ? e.message : String(e))
            );
        }
    }

    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
        window.ms365ActionLog.append({
            tool: 'sharepoint',
            action: 'create-schularbeiten-lists',
            target: url,
            summary: 'Schularbeiten-Paket (SAP-Regelwerk, SAP-Terminfenster, SAP-Schularbeiten, SAP-FachMeta)'
        });
    }

    write(
        'Fertig. Stammdaten (Klassen/Lehrer/Fächer) bleiben in den Schultools – SAP-FachMeta optional für Farbe/Kontingent. Mehrere Schuljahre über Spalte „Schuljahr“.'
    );
    return {
        siteId,
        lists: {
            regelwerk,
            terminfenster,
            schularbeiten,
            fachMeta
        }
    };
}

/**
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 */
async function probeListsHealth(webUrl, logFn) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

    const token = await ensureToken();
    write('Löse Website auf …');
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    const report = [];
    let allOk = true;

    for (let i = 0; i < LIST_KEYS.length; i++) {
        const key = LIST_KEYS[i];
        const required = REQUIRED_COLUMNS_BY_KEY[key];
        const canonical = LIST_TITLES[key];
        const resolved = await resolvePlanerList(token, siteId, key, findListByDisplayName);
        const list = resolved && resolved.list;
        const title = (resolved && resolved.displayName) || canonical;
        if (!list || !list.id) {
            write('Fehlt: Liste „' + canonical + '" (auch Legacy: ' + titlesForListKey(key).join(', ') + ')');
            report.push({ title: canonical, ok: false, missingColumns: required.slice(), missingList: true });
            allOk = false;
            continue;
        }
        if (resolved.isLegacyTitle) {
            write('Hinweis: „' + title + '" noch ohne SAP-Präfix – Setup erneut ausführen zum Umbenennen.');
        }
        const colsPath =
            G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(list.id) + '/columns?$top=200';
        const colsData = await G().graphJson('GET', colsPath, token, undefined, 'v1.0');
        const colNames = ((colsData && colsData.value) || [])
            .map((c) => String((c && c.name) || '').trim())
            .filter(Boolean);
        const missing = required.filter((n) => colNames.indexOf(n) === -1);
        if (missing.length) {
            write('„' + title + '": fehlende Spalten: ' + missing.join(', '));
            allOk = false;
        } else {
            write('„' + title + '": Spalten OK' + (list.webUrl ? ' · ' + list.webUrl : ''));
        }
        report.push({
            title,
            ok: missing.length === 0,
            missingColumns: missing,
            webUrl: list.webUrl || '',
            listId: list.id
        });
    }

    const panel = $('spsaHealth');
    if (panel) {
        panel.hidden = false;
        panel.textContent = allOk
            ? 'Status: OK – SAP-Listen mit erwarteten Spalten (inkl. Schuljahr).'
            : 'Status: prüfen – Details im Protokoll.';
        panel.classList.toggle('ok', allOk);
        panel.classList.toggle('warn', !allOk);
    }

    return { ok: allOk, report };
}

function readPermissionsFromForm() {
    return normalizePermissionsConfig({
        ...readPermissionsFromPickers(SETUP_GROUP_FIELDS),
        skipPerms: !!($('spsaSkipPerms') && $('spsaSkipPerms').checked)
    });
}

function mountSetupEntraPermGrid() {
    const host = $('spsaEntraPermGrid');
    if (!host || host.dataset.saPermMounted === '1') return;
    host.innerHTML = htmlSchularbeitenEntraPermGrid(loadPermissionsConfig(), { mode: 'setup' });
    host.dataset.saPermMounted = '1';
}

function fillPermissionsForm(cfg) {
    mountSetupEntraPermGrid();
    fillPermissionsPickers(cfg, SETUP_GROUP_FIELDS);
    const c = normalizePermissionsConfig(cfg);
    if ($('spsaSkipPerms')) $('spsaSkipPerms').checked = c.skipPerms;
}

async function runCreate() {
    const logEl = $('spsaLog');
    if (logEl) logEl.textContent = '';
    const webUrl = String(($('spsaSiteUrl') && $('spsaSiteUrl').value) || '').trim();
    const perms = readPermissionsFromForm();
    savePermissionsConfig(perms);
    const result = await createSchularbeitenLists(webUrl, log, perms);
    toast('Schularbeiten-Listen angelegt bzw. ergänzt.');
    return result;
}

async function runApplyPermissionsOnly() {
    const logEl = $('spsaLog');
    if (logEl) logEl.textContent = '';
    const webUrl = String(($('spsaSiteUrl') && $('spsaSiteUrl').value) || '').trim();
    if (!webUrl) throw new Error('Website-URL fehlt.');
    const perms = readPermissionsFromForm();
    if (perms.skipPerms) throw new Error('„Rechte überspringen“ ist aktiv – Haken entfernen.');
    savePermissionsConfig(perms);
    await applySchularbeitenPackagePermissions(webUrl, perms, log);
    toast('Berechtigungen angewendet.');
}

window.ms365SpoSchularbeiten = {
    createLists: createSchularbeitenLists,
    createList: createSchularbeitenLists,
    probeListsHealth: probeListsHealth,
    applyPackagePermissions: applySchularbeitenPackagePermissions,
    loadPermissionsConfig: loadPermissionsConfig,
    savePermissionsConfig: savePermissionsConfig,
    LIST_TITLES
};

function persistSiteUrlForPlaner() {
    const url = String(($('spsaSiteUrl') && $('spsaSiteUrl').value) || '').trim();
    if (!url) return;
    try {
        localStorage.setItem(PLANER_SITE_STORAGE_KEY, url);
    } catch {
        /* ignore */
    }
}

function wirePlanerLinks() {
    document.querySelectorAll('.spsa-go-planer').forEach((el) => {
        el.addEventListener('click', () => persistSiteUrlForPlaner());
    });
}

function wireUi() {
    wirePlanerLinks();

    const runBtn = $('spsaBtnRun');
    if (runBtn) {
        runBtn.addEventListener('click', () => {
            if (
                !window.confirm(
                    'Schularbeiten-Paket auf der Website anlegen bzw. fehlende Spalten ergänzen?\n\nListen: SAP-Regelwerk, SAP-Terminfenster, SAP-Schularbeiten, SAP-FachMeta (Legacy-Namen werden umbenannt).'
                )
            ) {
                return;
            }
            runCreate().catch((e) => {
                log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                toast('Fehler: ' + (e && e.message ? e.message : e));
            });
        });
    }

    const probeBtn = $('spsaBtnProbe');
    if (probeBtn) {
        probeBtn.addEventListener('click', () => {
            if ($('spsaLog')) $('spsaLog').textContent = '';
            const webUrl = String(($('spsaSiteUrl') && $('spsaSiteUrl').value) || '').trim();
            if (!webUrl) {
                toast('Website-URL fehlt.');
                return;
            }
            ensureToken()
                .then((token) => G().resolveSiteFromWebUrl(token, webUrl))
                .then((site) => {
                    log('Site gefunden: ' + (site.displayName || '') + '\nid: ' + (site.id || ''));
                    if (site.webUrl) log('webUrl: ' + site.webUrl);
                    toast('Website erkannt.');
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    const healthBtn = $('spsaBtnHealth');
    if (healthBtn) {
        healthBtn.addEventListener('click', () => {
            if ($('spsaLog')) $('spsaLog').textContent = '';
            const webUrl = String(($('spsaSiteUrl') && $('spsaSiteUrl').value) || '').trim();
            probeListsHealth(webUrl)
                .then((summary) => {
                    toast(summary && summary.ok ? 'Listen-Check OK' : 'Listen-Check: bitte Protokoll prüfen');
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    fillPermissionsForm(loadPermissionsConfig());

    wirePermissionGroupPickers(SETUP_GROUP_FIELDS, () => persistPickersToStorage(SETUP_GROUP_FIELDS));
    initSchularbeitenSetupExtraUsers(() => persistPickersToStorage(SETUP_GROUP_FIELDS));
    const skipEl = $('spsaSkipPerms');
    if (skipEl) {
        skipEl.addEventListener('change', () =>
            persistPickersToStorage(SETUP_GROUP_FIELDS, skipEl.checked)
        );
    }

    const permsBtn = $('spsaBtnPerms');
    if (permsBtn) {
        permsBtn.addEventListener('click', () => {
            if ($('spsaLog')) $('spsaLog').textContent = '';
            runApplyPermissionsOnly().catch((e) => {
                log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                toast('Fehler: ' + (e && e.message ? e.message : e));
            });
        });
    }

    try {
        const setup =
            window.ms365AppDataV2 && window.ms365AppDataV2.getSetup
                ? window.ms365AppDataV2.getSetup()
                : null;
        const saved = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
        if (saved && $('spsaSiteUrl') && !$('spsaSiteUrl').value) $('spsaSiteUrl').value = saved;
        const adminSite =
            setup && setup.schularbeitenSiteUrl ? String(setup.schularbeitenSiteUrl).trim() : '';
        if (adminSite && $('spsaSiteUrl') && !$('spsaSiteUrl').value) $('spsaSiteUrl').value = adminSite;
    } catch {
        /* ignore */
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wireUi);
} else {
    wireUi();
}
