/**
 * SharePoint: Wahlfach-Zuteilung-Listen anlegen.
 */
import {
    LIST_TITLES,
    RUNDEN_COLUMNS,
    ANGEBOTE_COLUMNS,
    ANMELDUNGEN_COLUMNS,
    ZUTEILUNGEN_COLUMNS,
    REQUIRED_COLUMNS,
    toGraphColumnBody
} from '../wahlfach-zuteilung/wahlfach-zuteilung-schema.js';
import { loadSettings, saveSettings } from '../wahlfach-zuteilung/wahlfach-zuteilung-settings.js';

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
    const el = $('spwzLog');
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

export async function createWahlfachLists(webUrl, logFn) {
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
        runden: settings.listRunden || LIST_TITLES.runden,
        angebote: settings.listAngebote || LIST_TITLES.angebote,
        anmeldungen: settings.listAnmeldungen || LIST_TITLES.anmeldungen,
        zuteilungen: settings.listZuteilungen || LIST_TITLES.zuteilungen
    };

    const runden = await ensureList(token, siteId, titles.runden, RUNDEN_COLUMNS, write);
    const angebote = await ensureList(token, siteId, titles.angebote, ANGEBOTE_COLUMNS, write);
    const anmeldungen = await ensureList(token, siteId, titles.anmeldungen, ANMELDUNGEN_COLUMNS, write);
    const zuteilungen = await ensureList(token, siteId, titles.zuteilungen, ZUTEILUNGEN_COLUMNS, write);

    saveSettings({ siteUrl: url });

    return {
        siteId,
        lists: {
            runden: { id: runden.id, webUrl: runden.webUrl || '' },
            angebote: { id: angebote.id, webUrl: angebote.webUrl || '' },
            anmeldungen: { id: anmeldungen.id, webUrl: anmeldungen.webUrl || '' },
            zuteilungen: { id: zuteilungen.id, webUrl: zuteilungen.webUrl || '' }
        }
    };
}

export async function probeWahlfachSite(webUrl) {
    const token = await ensureToken();
    return G().resolveSiteFromWebUrl(token, String(webUrl || '').trim());
}

export async function healthWahlfachLists(webUrl) {
    const token = await ensureToken();
    const site = await G().resolveSiteFromWebUrl(token, String(webUrl || '').trim());
    const siteId = site && site.id ? String(site.id) : '';
    const settings = loadSettings();
    const titles = [
        settings.listRunden || LIST_TITLES.runden,
        settings.listAngebote || LIST_TITLES.angebote,
        settings.listAnmeldungen || LIST_TITLES.anmeldungen,
        settings.listZuteilungen || LIST_TITLES.zuteilungen
    ];
    const report = [];
    for (const title of titles) {
        const list = await findListByTitle(token, siteId, title);
        if (!list) {
            report.push({ title, ok: false, detail: 'fehlt' });
            continue;
        }
        const colsPath =
            G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(list.id) + '/columns?$top=200';
        const colsData = await G().graphJson('GET', colsPath, token, undefined, 'v1.0');
        const names = new Set(((colsData && colsData.value) || []).map((c) => c.name));
        const required = REQUIRED_COLUMNS[title] || REQUIRED_COLUMNS[LIST_TITLES.runden] || [];
        // REQUIRED_COLUMNS keys are default titles; match flexibly
        const req =
            REQUIRED_COLUMNS[title] ||
            (title.includes('Runde')
                ? REQUIRED_COLUMNS[LIST_TITLES.runden]
                : title.includes('Angebot')
                  ? REQUIRED_COLUMNS[LIST_TITLES.angebote]
                  : title.includes('Anmeldung')
                    ? REQUIRED_COLUMNS[LIST_TITLES.anmeldungen]
                    : REQUIRED_COLUMNS[LIST_TITLES.zuteilungen]) ||
            [];
        const missing = req.filter((n) => !names.has(n));
        report.push({
            title,
            ok: missing.length === 0,
            detail: missing.length ? 'fehlt: ' + missing.join(', ') : 'ok'
        });
        void required;
    }
    return report;
}

document.addEventListener('DOMContentLoaded', () => {
    const urlEl = $('spwzSiteUrl');
    const settings = loadSettings();
    /* Prefill aus Settings / school-tool-sites kind `wahlfach` */
    if (urlEl && !String(urlEl.value || '').trim()) {
        urlEl.value = String(settings.siteUrl || '').trim();
    }
    ['spwzListRunden', 'spwzListAngebote', 'spwzListAnmeldungen', 'spwzListZuteilungen'].forEach(
        (id, i) => {
            const el = $(id);
            if (!el) return;
            const keys = ['listRunden', 'listAngebote', 'listAnmeldungen', 'listZuteilungen'];
            if (!String(el.value || '').trim()) el.value = settings[keys[i]] || '';
        }
    );

    const saveBtn = $('spwzBtnSaveSettings');
    if (saveBtn) {
        saveBtn.addEventListener('click', () => {
            saveSettings({
                siteUrl: String((urlEl && urlEl.value) || '').trim(),
                listRunden: String(($('spwzListRunden') && $('spwzListRunden').value) || '').trim(),
                listAngebote: String(($('spwzListAngebote') && $('spwzListAngebote').value) || '').trim(),
                listAnmeldungen: String(
                    ($('spwzListAnmeldungen') && $('spwzListAnmeldungen').value) || ''
                ).trim(),
                listZuteilungen: String(
                    ($('spwzListZuteilungen') && $('spwzListZuteilungen').value) || ''
                ).trim()
            });
            toast('Einstellungen gespeichert.');
        });
    }

    const run = $('spwzBtnRun');
    if (run) {
        run.addEventListener('click', () => {
            saveSettings({
                siteUrl: String((urlEl && urlEl.value) || '').trim(),
                listRunden: String(($('spwzListRunden') && $('spwzListRunden').value) || '').trim(),
                listAngebote: String(($('spwzListAngebote') && $('spwzListAngebote').value) || '').trim(),
                listAnmeldungen: String(
                    ($('spwzListAnmeldungen') && $('spwzListAnmeldungen').value) || ''
                ).trim(),
                listZuteilungen: String(
                    ($('spwzListZuteilungen') && $('spwzListZuteilungen').value) || ''
                ).trim()
            });
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            if ($('spwzLog')) $('spwzLog').textContent = '';
            run.disabled = true;
            createWahlfachLists(webUrl, log)
                .then(() => toast('Wahlfach-Listen angelegt bzw. geprüft.'))
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                })
                .finally(() => {
                    run.disabled = false;
                });
        });
    }

    const probe = $('spwzBtnProbe');
    if (probe) {
        probe.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            probeWahlfachSite(webUrl)
                .then((site) => toast('Site OK: ' + (site.displayName || site.id)))
                .catch((e) => toast('Fehler: ' + (e && e.message ? e.message : e)));
        });
    }

    const health = $('spwzBtnHealth');
    if (health) {
        health.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            const box = $('spwzHealth');
            healthWahlfachLists(webUrl)
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
                })
                .catch((e) => toast('Fehler: ' + (e && e.message ? e.message : e)));
        });
    }
});
