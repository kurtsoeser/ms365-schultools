/**
 * SharePoint: Ausgabenverwaltung-Listen anlegen.
 */
import {
    GEGENSTAENDE_COLUMNS,
    LIST_TITLES,
    REQUIRED_COLUMNS,
    VORGAENGE_COLUMNS,
    toGraphColumnBody
} from '../ausgaben/ausgaben-schema.js';
import {
    loadAusgabenSpSettings,
    saveAusgabenSpSettings
} from '../ausgaben/ausgaben-settings.js';

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
    const el = $('spagLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + msg;
    el.scrollTop = el.scrollHeight;
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function loadSettings() {
    return loadAusgabenSpSettings();
}

function saveSettings(partial) {
    return saveAusgabenSpSettings(partial);
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

export async function createAusgabenLists(webUrl, logFn) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

    const token = await ensureToken();
    write('Löse Website auf …');
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    write('Site: ' + (site.displayName || siteId));

    const settings = loadSettings();
    const titles = {
        gegenstaende: settings.listGegenstaende || LIST_TITLES.gegenstaende,
        vorgaenge: settings.listVorgaenge || LIST_TITLES.vorgaenge
    };

    const gegenstaende = await ensureList(
        token,
        siteId,
        titles.gegenstaende,
        GEGENSTAENDE_COLUMNS,
        write
    );
    const vorgaenge = await ensureList(token, siteId, titles.vorgaenge, VORGAENGE_COLUMNS, write);

    saveSettings({ siteUrl: url });

    return {
        siteId,
        lists: {
            gegenstaende: { id: gegenstaende.id, webUrl: gegenstaende.webUrl || '' },
            vorgaenge: { id: vorgaenge.id, webUrl: vorgaenge.webUrl || '' }
        }
    };
}

document.addEventListener('DOMContentLoaded', () => {
    const urlEl = $('spagSiteUrl');
    const settings = loadSettings();
    if (urlEl && !String(urlEl.value || '').trim()) urlEl.value = settings.siteUrl || '';
    const gEl = $('spagListGegenstaende');
    const vEl = $('spagListVorgaenge');
    if (gEl && !String(gEl.value || '').trim()) {
        gEl.value = settings.listGegenstaende || LIST_TITLES.gegenstaende;
    }
    if (vEl && !String(vEl.value || '').trim()) {
        vEl.value = settings.listVorgaenge || LIST_TITLES.vorgaenge;
    }

    $('spagBtnSaveSettings')?.addEventListener('click', () => {
        saveSettings({
            siteUrl: String((urlEl && urlEl.value) || '').trim(),
            listGegenstaende: String((gEl && gEl.value) || '').trim(),
            listVorgaenge: String((vEl && vEl.value) || '').trim()
        });
        toast('Einstellungen gespeichert.');
    });

    const run = $('spagBtnRun');
    if (run) {
        run.addEventListener('click', () => {
            saveSettings({
                siteUrl: String((urlEl && urlEl.value) || '').trim(),
                listGegenstaende: String((gEl && gEl.value) || '').trim(),
                listVorgaenge: String((vEl && vEl.value) || '').trim()
            });
            if ($('spagLog')) $('spagLog').textContent = '';
            run.disabled = true;
            createAusgabenLists(String((urlEl && urlEl.value) || '').trim(), log)
                .then(() => toast('Ausgaben-Paket angelegt bzw. geprüft.'))
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                })
                .finally(() => {
                    run.disabled = false;
                });
        });
    }

    $('spagBtnProbe')?.addEventListener('click', () => {
        ensureToken()
            .then((token) => G().resolveSiteFromWebUrl(token, String((urlEl && urlEl.value) || '').trim()))
            .then((site) => toast('Site OK: ' + (site.displayName || site.id)))
            .catch((e) => toast('Fehler: ' + (e && e.message ? e.message : e)));
    });

    void REQUIRED_COLUMNS;
});
