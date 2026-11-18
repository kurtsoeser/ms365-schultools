/**
 * SharePoint: Elternbrief-Listen anlegen.
 */
import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';
import {
    BRIEFE_COLUMNS,
    LESE_COLUMNS,
    LIST_TITLES,
    REQUIRED_COLUMNS,
    toGraphColumnBody
} from '../elternbrief/elternbrief-schema.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

const SETTINGS_KEY = 'ms365-elternbrief-sp-settings-v1';

function G() {
    const api = window.ms365SpoGraph;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

function $(id) {
    return document.getElementById(id);
}

function log(msg) {
    const el = $('spebLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + msg;
    el.scrollTop = el.scrollHeight;
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function loadSettings() {
    try {
        const raw = localStorage.getItem(SETTINGS_KEY);
        return raw ? JSON.parse(raw) : {};
    } catch {
        return {};
    }
}

function saveSettings(partial) {
    const next = { ...loadSettings(), ...partial };
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(next));
    return next;
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
            try {
                await G().graphJson('POST', base, token, toGraphColumnBody(def), 'v1.0');
            } catch (err) {
                const detail = err && err.message ? err.message : String(err);
                throw new Error('Spalte „' + def.name + '“: ' + detail);
            }
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

export async function createElternbriefLists(webUrl, logFn) {
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
        briefe: settings.listBriefe || LIST_TITLES.briefe,
        lese: settings.listLese || LIST_TITLES.lesebestaetigungen
    };

    const briefe = await ensureList(token, siteId, titles.briefe, BRIEFE_COLUMNS, write);
    const lese = await ensureList(token, siteId, titles.lese, LESE_COLUMNS, write);

    saveSettings({ siteUrl: url });

    return {
        siteId,
        lists: {
            briefe: { id: briefe.id, webUrl: briefe.webUrl || '' },
            lesebestaetigungen: { id: lese.id, webUrl: lese.webUrl || '' }
        }
    };
}

export async function probeElternbriefSite(webUrl) {
    const token = await ensureToken();
    return G().resolveSiteFromWebUrl(token, String(webUrl || '').trim());
}

export async function healthElternbriefLists(webUrl) {
    const token = await ensureToken();
    const site = await G().resolveSiteFromWebUrl(token, String(webUrl || '').trim());
    const siteId = site && site.id ? String(site.id) : '';
    const settings = loadSettings();
    const titles = [
        settings.listBriefe || LIST_TITLES.briefe,
        settings.listLese || LIST_TITLES.lesebestaetigungen
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
        const req =
            REQUIRED_COLUMNS[title] ||
            (title.toLowerCase().includes('lese')
                ? REQUIRED_COLUMNS[LIST_TITLES.lesebestaetigungen]
                : REQUIRED_COLUMNS[LIST_TITLES.briefe]) ||
            [];
        const missing = req.filter((n) => !names.has(n));
        report.push({
            title,
            ok: missing.length === 0,
            detail: missing.length ? 'fehlt: ' + missing.join(', ') : 'ok'
        });
    }
    return report;
}

document.addEventListener('DOMContentLoaded', () => {
    const urlEl = $('spebSiteUrl');
    const settings = loadSettings();
    if (urlEl && !String(urlEl.value || '').trim()) {
        const fromShared = String(resolveSchoolToolSiteUrl('elternbrief') || '').trim();
        urlEl.value = settings.siteUrl || fromShared || '';
    }
    const briefeEl = $('spebListBriefe');
    const leseEl = $('spebListLese');
    if (briefeEl && !String(briefeEl.value || '').trim()) {
        briefeEl.value = settings.listBriefe || LIST_TITLES.briefe;
    }
    if (leseEl && !String(leseEl.value || '').trim()) {
        leseEl.value = settings.listLese || LIST_TITLES.lesebestaetigungen;
    }

    const saveBtn = $('spebBtnSaveSettings');
    if (saveBtn) {
        saveBtn.addEventListener('click', () => {
            saveSettings({
                siteUrl: String((urlEl && urlEl.value) || '').trim(),
                listBriefe: String((briefeEl && briefeEl.value) || '').trim(),
                listLese: String((leseEl && leseEl.value) || '').trim()
            });
            toast('Einstellungen gespeichert.');
        });
    }

    const run = $('spebBtnRun');
    if (run) {
        run.addEventListener('click', () => {
            saveSettings({
                siteUrl: String((urlEl && urlEl.value) || '').trim(),
                listBriefe: String((briefeEl && briefeEl.value) || '').trim(),
                listLese: String((leseEl && leseEl.value) || '').trim()
            });
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            if ($('spebLog')) $('spebLog').textContent = '';
            run.disabled = true;
            createElternbriefLists(webUrl, log)
                .then(() => toast('Elternbrief-Listen angelegt bzw. geprüft.'))
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                })
                .finally(() => {
                    run.disabled = false;
                });
        });
    }

    const probe = $('spebBtnProbe');
    if (probe) {
        probe.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            probeElternbriefSite(webUrl)
                .then((site) => toast('Site OK: ' + (site.displayName || site.id)))
                .catch((e) => toast('Fehler: ' + (e && e.message ? e.message : e)));
        });
    }

    const health = $('spebBtnHealth');
    if (health) {
        health.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            const box = $('spebHealth');
            healthElternbriefLists(webUrl)
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
