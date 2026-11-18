/**
 * SharePoint: Einsammler-Listen anlegen.
 */
import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';
import {
    ANTWORTEN_COLUMNS,
    KOEPFE_COLUMNS,
    LIST_TITLES,
    REQUIRED_COLUMNS,
    toGraphColumnBody
} from '../einsammler/einsammler-schema.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

const SETTINGS_KEY = 'ms365-einsammler-sp-settings-v1';

function G() {
    const api = window.ms365SpoGraph;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

function $(id) {
    return document.getElementById(id);
}

function log(msg) {
    const el = $('speinLog');
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

export async function createEinsammlerLists(webUrl, logFn) {
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
        koepfe: settings.listKoepfe || LIST_TITLES.koepfe,
        antworten: settings.listAntworten || LIST_TITLES.antworten
    };

    const koepfe = await ensureList(token, siteId, titles.koepfe, KOEPFE_COLUMNS, write);
    const antworten = await ensureList(token, siteId, titles.antworten, ANTWORTEN_COLUMNS, write);

    saveSettings({ siteUrl: url });

    return {
        siteId,
        lists: {
            koepfe: { id: koepfe.id, webUrl: koepfe.webUrl || '' },
            antworten: { id: antworten.id, webUrl: antworten.webUrl || '' }
        }
    };
}

export async function probeEinsammlerSite(webUrl) {
    const token = await ensureToken();
    return G().resolveSiteFromWebUrl(token, String(webUrl || '').trim());
}

export async function healthEinsammlerLists(webUrl) {
    const token = await ensureToken();
    const site = await G().resolveSiteFromWebUrl(token, String(webUrl || '').trim());
    const siteId = site && site.id ? String(site.id) : '';
    const settings = loadSettings();
    const titles = [
        settings.listKoepfe || LIST_TITLES.koepfe,
        settings.listAntworten || LIST_TITLES.antworten
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
            (title.toLowerCase().includes('antwort')
                ? REQUIRED_COLUMNS[LIST_TITLES.antworten]
                : REQUIRED_COLUMNS[LIST_TITLES.koepfe]) ||
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
    const urlEl = $('speinSiteUrl');
    const settings = loadSettings();
    if (urlEl && !String(urlEl.value || '').trim()) {
        const fromShared = String(resolveSchoolToolSiteUrl('einsammler') || '').trim();
        urlEl.value = settings.siteUrl || fromShared || '';
    }
    const koepfeEl = $('speinListKoepfe');
    const antEl = $('speinListAntworten');
    if (koepfeEl && !String(koepfeEl.value || '').trim()) {
        koepfeEl.value = settings.listKoepfe || LIST_TITLES.koepfe;
    }
    if (antEl && !String(antEl.value || '').trim()) {
        antEl.value = settings.listAntworten || LIST_TITLES.antworten;
    }

    const saveBtn = $('speinBtnSaveSettings');
    if (saveBtn) {
        saveBtn.addEventListener('click', () => {
            saveSettings({
                siteUrl: String((urlEl && urlEl.value) || '').trim(),
                listKoepfe: String((koepfeEl && koepfeEl.value) || '').trim(),
                listAntworten: String((antEl && antEl.value) || '').trim()
            });
            toast('Einstellungen gespeichert.');
        });
    }

    const run = $('speinBtnRun');
    if (run) {
        run.addEventListener('click', () => {
            saveSettings({
                siteUrl: String((urlEl && urlEl.value) || '').trim(),
                listKoepfe: String((koepfeEl && koepfeEl.value) || '').trim(),
                listAntworten: String((antEl && antEl.value) || '').trim()
            });
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            if ($('speinLog')) $('speinLog').textContent = '';
            run.disabled = true;
            createEinsammlerLists(webUrl, log)
                .then(() => toast('Einsammler-Listen angelegt bzw. geprüft.'))
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                })
                .finally(() => {
                    run.disabled = false;
                });
        });
    }

    const probe = $('speinBtnProbe');
    if (probe) {
        probe.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            probeEinsammlerSite(webUrl)
                .then((site) => toast('Site OK: ' + (site.displayName || site.id)))
                .catch((e) => toast('Fehler: ' + (e && e.message ? e.message : e)));
        });
    }

    const health = $('speinBtnHealth');
    if (health) {
        health.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            const box = $('speinHealth');
            healthEinsammlerLists(webUrl)
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
