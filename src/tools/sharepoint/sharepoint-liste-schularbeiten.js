/**
 * SharePoint: Schularbeiten-Listen anlegen (Regelwerk, Terminfenster, Schularbeiten).
 */
import {
    LIST_TITLES,
    REGELWERK_COLUMNS,
    TERMINFENSTER_COLUMNS,
    SCHULARBEITEN_COLUMNS,
    FACHMETA_COLUMNS,
    REQUIRED_COLUMNS,
    DEFAULT_REGELWERK_FIELDS,
    toGraphColumnBody
} from '../schularbeiten-planer/schularbeiten-planer-schema.js';

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
    const title = String(listTitle || '').trim();
    const path =
        G().graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await G().graphJson('GET', path, token, undefined, 'v1.0');
    const list = (data && data.value) || [];
    return list[0] || null;
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

async function ensureList(token, siteId, title, columnDefs, write) {
    let list = await findListByTitle(token, siteId, title);
    if (!list || !list.id) {
        write('Erstelle Liste „' + title + '" …');
        const created = await G().graphJson(
            'POST',
            G().graphPathSite(siteId) + '/lists',
            token,
            {
                displayName: title,
                list: { template: 'genericList' }
            },
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

/**
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 */
async function createSchularbeitenLists(webUrl, logFn) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

    const token = await ensureToken();
    write('Löse Website auf …');
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
    write('Site: ' + (site.displayName || siteId));

    const regelwerk = await ensureList(token, siteId, LIST_TITLES.regelwerk, REGELWERK_COLUMNS, write);
    await seedRegelwerkIfEmpty(token, siteId, regelwerk.id, write);

    const terminfenster = await ensureList(
        token,
        siteId,
        LIST_TITLES.terminfenster,
        TERMINFENSTER_COLUMNS,
        write
    );

    const schularbeiten = await ensureList(
        token,
        siteId,
        LIST_TITLES.schularbeiten,
        SCHULARBEITEN_COLUMNS,
        write
    );

    const fachMeta = await ensureList(token, siteId, LIST_TITLES.fachMeta, FACHMETA_COLUMNS, write);

    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
        window.ms365ActionLog.append({
            tool: 'sharepoint',
            action: 'create-schularbeiten-lists',
            target: url,
            summary: 'Schularbeiten-Paket (Regelwerk, Terminfenster, Schularbeiten, SA-FachMeta)'
        });
    }

    write('Fertig. Stammdaten (Klassen/Lehrer/Fächer) bleiben in den Schultools – optional SA-FachMeta für Farbe/Kontingent.');
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
    const titles = [
        [LIST_TITLES.regelwerk, REQUIRED_COLUMNS.Regelwerk],
        [LIST_TITLES.terminfenster, REQUIRED_COLUMNS.Terminfenster],
        [LIST_TITLES.schularbeiten, REQUIRED_COLUMNS.Schularbeiten],
        [LIST_TITLES.fachMeta, REQUIRED_COLUMNS['SA-FachMeta']]
    ];

    for (let i = 0; i < titles.length; i++) {
        const title = titles[i][0];
        const required = titles[i][1];
        const list = await findListByTitle(token, siteId, title);
        if (!list || !list.id) {
            write('Fehlt: Liste „' + title + '"');
            report.push({ title, ok: false, missingColumns: required.slice(), missingList: true });
            allOk = false;
            continue;
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
            ? 'Status: OK – Listen mit erwarteten Spalten (inkl. SA-FachMeta).'
            : 'Status: prüfen – Details im Protokoll.';
        panel.classList.toggle('ok', allOk);
        panel.classList.toggle('warn', !allOk);
    }

    return { ok: allOk, report };
}

async function runCreate() {
    const logEl = $('spsaLog');
    if (logEl) logEl.textContent = '';
    const webUrl = String(($('spsaSiteUrl') && $('spsaSiteUrl').value) || '').trim();
    const result = await createSchularbeitenLists(webUrl);
    toast('Schularbeiten-Listen angelegt bzw. ergänzt.');
    return result;
}

window.ms365SpoSchularbeiten = {
    createLists: createSchularbeitenLists,
    createList: createSchularbeitenLists,
    probeListsHealth: probeListsHealth,
    LIST_TITLES
};

function wireUi() {
    const runBtn = $('spsaBtnRun');
    if (runBtn) {
        runBtn.addEventListener('click', () => {
            if (
                !window.confirm(
                    'Schularbeiten-Paket auf der Website anlegen bzw. fehlende Spalten ergänzen?\n\nListen: Regelwerk, Terminfenster, Schularbeiten, SA-FachMeta'
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

    try {
        const setup =
            window.ms365AppDataV2 && window.ms365AppDataV2.getSetup
                ? window.ms365AppDataV2.getSetup()
                : null;
        const saved = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
        if (saved && $('spsaSiteUrl') && !$('spsaSiteUrl').value) $('spsaSiteUrl').value = saved;
    } catch {
        /* ignore */
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wireUi);
} else {
    wireUi();
}
