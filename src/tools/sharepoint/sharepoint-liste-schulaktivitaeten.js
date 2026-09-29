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
    const added = await addMissingColumns(siteId, list.id, token, columnDefs, write);
    write(added ? added + ' Spalte(n) ergänzt.' : 'Spalten vollständig.');
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

export async function createSchulaktivitaetenLists(webUrl, logFn) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

    const token = await ensureToken();
    write('Löse Website auf …');
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    write('Site: ' + (site.displayName || siteId));

    const regelwerk = await ensureList(token, siteId, LIST_TITLES.regelwerk, REGELWERK_COLUMNS, write);
    await seedRegelwerkIfEmpty(token, siteId, regelwerk.id, write);
    const aktivitaeten = await ensureList(
        token,
        siteId,
        LIST_TITLES.aktivitaeten,
        AKTIVITAETEN_COLUMNS,
        write
    );

    return {
        siteId,
        lists: {
            regelwerk: { id: regelwerk.id, webUrl: regelwerk.webUrl || '' },
            aktivitaeten: { id: aktivitaeten.id, webUrl: aktivitaeten.webUrl || '' }
        }
    };
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
    for (const [key, title] of Object.entries(LIST_TITLES)) {
        const list = await findListByTitle(token, siteId, title);
        if (!list) {
            report.push({ title, ok: false, detail: 'fehlt' });
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

document.addEventListener('DOMContentLoaded', () => {
    const urlEl = $('spaktSiteUrl');
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

    const run = $('spaktBtnRun');
    if (run) {
        run.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            if ($('spaktLog')) $('spaktLog').textContent = '';
            run.disabled = true;
            createSchulaktivitaetenLists(webUrl, log)
                .then(() => toast('Schulaktivitäten-Paket angelegt bzw. geprüft.'))
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                })
                .finally(() => {
                    run.disabled = false;
                });
        });
    }

    const probe = $('spaktBtnProbe');
    if (probe) {
        probe.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            probeSchulaktivitaetenSite(webUrl)
                .then((site) => toast('Site OK: ' + (site.displayName || site.id)))
                .catch((e) => toast('Fehler: ' + (e && e.message ? e.message : e)));
        });
    }

    const health = $('spaktBtnHealth');
    if (health) {
        health.addEventListener('click', () => {
            const webUrl = String((urlEl && urlEl.value) || '').trim();
            const box = $('spaktHealth');
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
                })
                .catch((e) => toast('Fehler: ' + (e && e.message ? e.message : e)));
        });
    }
});
