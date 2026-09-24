/**
 * SharePoint: Projektwochen-Listen anlegen (PW-Aktionen, PW-Angebote).
 */
import {
    LIST_TITLES,
    AKTIONEN_COLUMNS,
    ANGEBOTE_COLUMNS,
    REQUIRED_COLUMNS,
    DEFAULT_AKTION_TEMPLATE,
    toGraphColumnBody,
    nextProjectWeekRange
} from '../projektwochen/projektwochen-schema.js';
import { persistSiteUrl, loadSavedSiteUrl } from '../projektwochen/projektwochen-state.js';

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
    const el = $('sppwLog');
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
        // Legacy: frühere Schema-Version nutzte „Tag“ statt „Wochentag“
        if (def.name === 'Wochentag' && existing.has('Tag')) {
            write('  · Spalte Tag vorhanden (Legacy) – Wochentag übersprungen.');
            continue;
        }
        try {
            await G().graphJson('POST', base, token, toGraphColumnBody(def), 'v1.0');
            added++;
            existing.add(def.name);
            write('  + Spalte ' + def.name);
        } catch (e) {
            const msg = e && e.message ? e.message : String(e);
            write('  ! Spalte „' + def.name + '" fehlgeschlagen: ' + msg);
            throw new Error('Spalte „' + def.name + '" konnte nicht angelegt werden: ' + msg);
        }
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

async function seedAktionIfEmpty(token, siteId, listId, write) {
    const itemsPath =
        G().graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(listId) +
        '/items?$expand=fields&$top=5';
    const data = await G().graphJson('GET', itemsPath, token, undefined, 'v1.0');
    const items = (data && data.value) || [];
    if (items.length) {
        write('PW-Aktionen enthält bereits ' + items.length + ' Eintrag/Einträge – Seed übersprungen.');
        return false;
    }
    const range = nextProjectWeekRange();
    write('Lege Demo-Projektwoche an (' + range.startIso + ' – ' + range.endIso + ') …');
    await G().graphJson(
        'POST',
        G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/items',
        token,
        {
            fields: {
                ...DEFAULT_AKTION_TEMPLATE,
                Startdatum: range.startIso,
                Enddatum: range.endIso,
                BuchungAbDefault: range.buchungAbIso
            }
        },
        'v1.0'
    );
    write('Seed: „' + DEFAULT_AKTION_TEMPLATE.Title + '" (' + DEFAULT_AKTION_TEMPLATE.AktionId + ')');
    return true;
}

/**
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 */
export async function createProjektwochenLists(webUrl, logFn) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

    const token = await ensureToken();
    write('Löse Website auf …');
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
    write('Site: ' + (site.displayName || siteId));

    const aktionen = await ensureList(token, siteId, LIST_TITLES.aktionen, AKTIONEN_COLUMNS, write);
    await seedAktionIfEmpty(token, siteId, aktionen.id, write);

    const angebote = await ensureList(token, siteId, LIST_TITLES.angebote, ANGEBOTE_COLUMNS, write);

    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
        window.ms365ActionLog.append({
            tool: 'sharepoint',
            action: 'create-projektwochen-lists',
            target: url,
            summary: 'Projektwochen-Paket (PW-Aktionen, PW-Angebote)'
        });
    }

    write('Fertig. Stammdaten (Klassen/Lehrer) bleiben in den Schultools. Bookings-Sync folgt in Phase 3.');
    return {
        siteId,
        lists: {
            aktionen,
            angebote
        }
    };
}

/**
 * @param {string} webUrl
 * @param {(msg: string) => void} [logFn]
 */
export async function probeListsHealth(webUrl, logFn) {
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
        [LIST_TITLES.aktionen, REQUIRED_COLUMNS['PW-Aktionen']],
        [LIST_TITLES.angebote, REQUIRED_COLUMNS['PW-Angebote']]
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
        const missing = required.filter((n) => {
            if (colNames.indexOf(n) !== -1) return false;
            // Legacy-Spalte Tag zählt als Wochentag
            if (n === 'Wochentag' && colNames.indexOf('Tag') !== -1) return false;
            return true;
        });
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

    const panel = $('sppwHealth');
    if (panel) {
        panel.hidden = false;
        panel.textContent = allOk
            ? 'Status: OK – PW-Aktionen und PW-Angebote mit erwarteten Spalten.'
            : 'Status: prüfen – Details im Protokoll.';
        panel.classList.toggle('ok', allOk);
        panel.classList.toggle('warn', !allOk);
    }

    return { ok: allOk, report };
}

async function runCreate() {
    const logEl = $('sppwLog');
    if (logEl) logEl.textContent = '';
    const webUrl = String(($('sppwSiteUrl') && $('sppwSiteUrl').value) || '').trim();
    persistSiteUrl(webUrl);
    const result = await createProjektwochenLists(webUrl);
    toast('Projektwochen-Listen angelegt bzw. ergänzt.');
    return result;
}

window.ms365SpoProjektwochen = {
    createLists: createProjektwochenLists,
    createList: createProjektwochenLists,
    probeListsHealth: probeListsHealth,
    LIST_TITLES
};

function wireUi() {
    const runBtn = $('sppwBtnRun');
    if (runBtn) {
        runBtn.addEventListener('click', () => {
            if (
                !window.confirm(
                    'Projektwochen-Paket auf der Website anlegen bzw. fehlende Spalten ergänzen?\n\nListen: PW-Aktionen, PW-Angebote'
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

    const probeBtn = $('sppwBtnProbe');
    if (probeBtn) {
        probeBtn.addEventListener('click', () => {
            if ($('sppwLog')) $('sppwLog').textContent = '';
            const webUrl = String(($('sppwSiteUrl') && $('sppwSiteUrl').value) || '').trim();
            if (!webUrl) {
                toast('Website-URL fehlt.');
                return;
            }
            persistSiteUrl(webUrl);
            ensureToken()
                .then((token) => G().resolveSiteFromWebUrl(token, webUrl))
                .then((site) => {
                    log('Site gefunden: ' + (site.displayName || '') + '\nid: ' + (site.id || ''));
                    if (site.webUrl) log('webUrl: ' + site.webUrl);
                    toast('Website erkannt – URL gespeichert.');
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    const healthBtn = $('sppwBtnHealth');
    if (healthBtn) {
        healthBtn.addEventListener('click', () => {
            if ($('sppwLog')) $('sppwLog').textContent = '';
            const webUrl = String(($('sppwSiteUrl') && $('sppwSiteUrl').value) || '').trim();
            persistSiteUrl(webUrl);
            probeListsHealth(webUrl)
                .then((summary) => {
                    toast(summary && summary.ok ? 'Listen-Check OK – URL gespeichert' : 'Listen-Check: bitte Protokoll prüfen');
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    const urlEl = $('sppwSiteUrl');
    if (urlEl) {
        urlEl.addEventListener('change', () => {
            const webUrl = String(urlEl.value || '').trim();
            if (webUrl) persistSiteUrl(webUrl);
        });
        // Vorfüllung: PW-Key → Intranet-Setup
        try {
            const saved = loadSavedSiteUrl();
            if (saved && !String(urlEl.value || '').trim()) urlEl.value = saved;
        } catch {
            /* ignore */
        }
        // Bereits eingetragene URL (z. B. aus vorherigem Versuch) sofort merken
        const current = String(urlEl.value || '').trim();
        if (current) persistSiteUrl(current);
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        if (document.getElementById('sppwBtnRun') || document.getElementById('sppwSiteUrl')) wireUi();
    });
} else if (document.getElementById('sppwBtnRun') || document.getElementById('sppwSiteUrl')) {
    wireUi();
}

window.ms365ProjektwochenLists = {
    createProjektwochenLists,
    probeListsHealth
};
