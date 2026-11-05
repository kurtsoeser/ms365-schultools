/**
 * UI: Bestehende Kursteams aus Microsoft 365 importieren → Unterrichtsbelegung (App-Daten).
 */
import { getGraphToken } from './kursteam-graph.js';
import {
    buildBelegungRowsFromGraphGroups,
    mergeBelegungWithGraphImport,
    parseKursteamDisplayName
} from '../../shared/kursteam-graph-import-logic.js';
import { summarizeBelegung } from '../../shared/unterrichtsbelegung-logic.js';
import {
    buildTeacherLookupFromCodeMap,
    enrichBelegungRows
} from '../../shared/unterrichtsbelegung-teacher-enrich-logic.js';
import {
    buildKursteamAbgleichReport,
    readKursteamStateFromBrowserStorage,
    resolvePlannedRowsForAbgleich
} from '../../shared/kursteam-belegung-abgleich-logic.js';
import { renderKursteamM365ImportStatusHtml } from '../../shared/kursteam-m365-import-status-ui.js';

const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

/** @type {Array<{ id: string, displayName: string, mailNickname: string, row: object, selected: boolean }>} */
let importCandidates = [];

/** @type {{ source: string, label: string, report: object }|null} */
let lastAbgleich = null;

function escapeHtml(text) {
    const d = document.createElement('div');
    d.textContent = text;
    return d.innerHTML;
}

function toast(msg) {
    if (typeof ns.showToast === 'function') ns.showToast(msg);
    else if (typeof window.ms365ShowToast === 'function') window.ms365ShowToast(msg);
    else window.alert(msg);
}

function getYearPrefixFromUi() {
    const el = document.getElementById('kursteamImportYearPrefix');
    if (el && el.value) return String(el.value).trim();
    const yp = document.getElementById('yearPrefix');
    if (yp && yp.value) return String(yp.value).trim();
    return 'SJ26-27';
}

function buildTeacherByCodeMap() {
    const map = new Map();
    try {
        const load =
            typeof window.ms365TenantSettingsLoad === 'function'
                ? window.ms365TenantSettingsLoad
                : null;
        const data = load ? load() : null;
        const teachers = data && Array.isArray(data.teachers) ? data.teachers : [];
        teachers.forEach((t) => {
            const code = String((t && t.code) || '')
                .trim()
                .toUpperCase();
            if (!code) return;
            const email = String((t && t.email) || '')
                .trim()
                .toLowerCase();
            map.set(code, { email, name: String((t && t.name) || '').trim() });
        });
    } catch {
        /* ignore */
    }
    return map;
}

async function graphJson(method, path, token, extraHeaders) {
    const url = path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path;
    const headers = Object.assign({ Authorization: 'Bearer ' + token }, extraHeaders || {});
    const res = await fetch(url, { method, headers });
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = text;
        }
    }
    if (!res.ok) {
        const msg =
            typeof data === 'object' && data && data.error
                ? JSON.stringify(data.error)
                : text || String(res.status);
        throw new Error(method + ' ' + path + ': ' + msg);
    }
    return data || {};
}

async function listTeamGroupsFromGraph(token) {
    const collected = [];
    let nextPath =
        "/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,mailNickname,resourceProvisioningOptions&$top=999";
    try {
        while (nextPath) {
            const useEv = nextPath.indexOf('http') !== 0;
            const data = await graphJson(
                'GET',
                nextPath,
                token,
                useEv ? { ConsistencyLevel: 'eventual' } : undefined
            );
            (data.value || []).forEach((g) => {
                if (g && g.mailNickname) collected.push(g);
            });
            nextPath = data['@odata.nextLink'] || null;
        }
        return collected;
    } catch {
        return listTeamGroupsFromGraphFallback(token);
    }
}

async function fetchOwnerEmailsForGroups(groups, token) {
    const map = new Map();
    const list = (Array.isArray(groups) ? groups : []).filter((g) => g && g.id);
    const batchSize = 10;
    for (let i = 0; i < list.length; i += batchSize) {
        const slice = list.slice(i, i + batchSize);
        await Promise.all(
            slice.map(async (g) => {
                try {
                    const data = await graphJson(
                        'GET',
                        '/groups/' + encodeURIComponent(g.id) + '/owners?$select=mail,userPrincipalName&$top=5',
                        token
                    );
                    const o = (data.value || []).find((x) => x && (x.mail || x.userPrincipalName));
                    if (o) {
                        map.set(
                            g.id,
                            String(o.mail || o.userPrincipalName || '')
                                .trim()
                                .toLowerCase()
                        );
                    }
                } catch {
                    /* ignore single group */
                }
            })
        );
    }
    return map;
}

async function listTeamGroupsFromGraphFallback(token) {
    const collected = [];
    let nextPath = '/groups?$select=id,displayName,mailNickname,resourceProvisioningOptions&$top=999';
    while (nextPath) {
        const data = await graphJson('GET', nextPath, token);
        (data.value || []).forEach((g) => {
            const opts = g && g.resourceProvisioningOptions ? g.resourceProvisioningOptions : [];
            if (opts.indexOf('Team') !== -1 && g.mailNickname) collected.push(g);
        });
        nextPath = data['@odata.nextLink'] || null;
    }
    return collected;
}

function getM365RowsForAbgleich() {
    if (importCandidates.length) {
        return importCandidates.map((c) => c.row);
    }
    const api = window.ms365AppDataV2;
    if (!api || typeof api.getUnterrichtsbelegung !== 'function') return [];
    try {
        const snap = api.getUnterrichtsbelegung();
        return (snap && Array.isArray(snap.rows) ? snap.rows : []).filter((r) => r && String(r.graphGroupId || '').trim());
    } catch {
        return [];
    }
}

function runAbgleich() {
    const kursteamState = readKursteamStateFromBrowserStorage();
    let belegung = null;
    const api = window.ms365AppDataV2;
    if (api && typeof api.getUnterrichtsbelegung === 'function') {
        try {
            belegung = api.getUnterrichtsbelegung();
        } catch {
            belegung = null;
        }
    }
    const plannedMeta = resolvePlannedRowsForAbgleich({ kursteamState, belegungSnapshot: belegung });
    const m365Rows = getM365RowsForAbgleich();
    const report = buildKursteamAbgleichReport(plannedMeta.rows, m365Rows);
    lastAbgleich = { source: plannedMeta.source, label: plannedMeta.label, report };
    return lastAbgleich;
}

function renderImportStatusSection() {
    const host =
        document.getElementById('kursteamGraphImportStatusHost') ||
        document.getElementById('kursteamGraphAbgleichHost');
    if (!host) return;

    const m365Rows = getM365RowsForAbgleich();
    if (!m365Rows.length) {
        host.innerHTML = renderKursteamM365ImportStatusHtml(null, {
            emptyHint:
                'Zuerst „Teams aus M365 laden“ (oder bereits importierte Verknüpfungen in der App nutzen).'
        });
        return;
    }

    const data = lastAbgleich || runAbgleich();
    const rep = data && data.report ? data.report : null;
    if (!rep) {
        host.innerHTML = renderKursteamM365ImportStatusHtml(null, {
            emptyHint: 'Status konnte nicht berechnet werden.'
        });
        return;
    }

    host.innerHTML = renderKursteamM365ImportStatusHtml(rep, {
        planLabel: data.label || 'Katalog/Plan',
        noPlanHint:
            'Kein WebUntis-/Wizard-Stand im Browser – Stundenplan-Daten im Assistenten laden oder Teamliste erzeugen, dann Status aktualisieren.'
    });
}

function existingLinkedMails() {
    const set = new Set();
    const api = window.ms365AppDataV2;
    if (!api || typeof api.getUnterrichtsbelegung !== 'function') return set;
    try {
        const snap = api.getUnterrichtsbelegung();
        (snap && Array.isArray(snap.rows) ? snap.rows : []).forEach((r) => {
            const nick = String((r && r.gruppenmail) || '')
                .trim()
                .toLowerCase();
            if (nick && r.graphGroupId) set.add(nick);
        });
    } catch {
        /* ignore */
    }
    return set;
}

function renderImportTable() {
    const host = document.getElementById('kursteamGraphImportTableHost');
    const summary = document.getElementById('kursteamGraphImportSummary');
    if (!host) return;

    if (!importCandidates.length) {
        host.innerHTML = '<p class="muted" style="margin:0;">Noch keine Teams geladen.</p>';
        if (summary) summary.textContent = '';
        return;
    }

    const linked = existingLinkedMails();
    let parseWarn = 0;
    importCandidates.forEach((c) => {
        if (!c.row.parseOk) parseWarn++;
    });

    if (summary) {
        summary.textContent =
            importCandidates.length +
            ' Kursteam(s) erkannt' +
            (parseWarn ? ' · ' + parseWarn + ' mit unklarem Anzeigenamen' : '') +
            ' · ' +
            linked.size +
            ' bereits in der App verknüpft';
    }

    const rows = importCandidates
        .map((c, idx) => {
            const nick = String(c.row.gruppenmail || '').toLowerCase();
            const already = linked.has(nick);
            const warn = !c.row.parseOk;
            return (
                '<tr>' +
                '<td><input type="checkbox" data-kt-import-idx="' +
                idx +
                '"' +
                (c.selected ? ' checked' : '') +
                '></td>' +
                '<td>' +
                escapeHtml(c.row.klasse || '—') +
                '</td>' +
                '<td>' +
                escapeHtml(c.row.fach || '—') +
                '</td>' +
                '<td>' +
                escapeHtml(c.row.lehrerCode || '—') +
                '</td>' +
                '<td><code style="font-size:0.82em;">' +
                escapeHtml(c.mailNickname) +
                '</code></td>' +
                '<td>' +
                (already ? '<span class="muted">verknüpft</span>' : warn ? '<span style="color:#856404;">Name prüfen</span>' : 'neu') +
                '</td>' +
                '</tr>'
            );
        })
        .join('');

    host.innerHTML =
        '<div style="overflow:auto;max-height:360px;border:1px solid var(--border);border-radius:10px;">' +
        '<table class="data-table kt-graph-import-table" style="margin:0;font-size:0.88em;">' +
        '<thead><tr>' +
        '<th style="width:36px;"></th><th>Klasse</th><th>Fach</th><th>LK</th><th>Mail-Nickname</th><th>Status</th>' +
        '</tr></thead><tbody>' +
        rows +
        '</tbody></table></div>';

    host.querySelectorAll('input[data-kt-import-idx]').forEach((inp) => {
        inp.addEventListener('change', () => {
            const i = Number(inp.getAttribute('data-kt-import-idx'));
            if (importCandidates[i]) importCandidates[i].selected = inp.checked;
        });
    });
}

async function loadGraphKursteams() {
    const btn = document.getElementById('kursteamGraphImportLoad');
    if (btn) btn.disabled = true;
    importCandidates = [];
    renderImportTable();
    try {
        const yearPrefix = getYearPrefixFromUi();
        const token = await getGraphToken();
        const groups = await listTeamGroupsFromGraph(token);
        const teacherByCode = buildTeacherByCodeMap();
        const ownerByGroupId = await fetchOwnerEmailsForGroups(groups, token);
        const rawRows = buildBelegungRowsFromGraphGroups(groups, { yearPrefix, teacherByCode });
        const lookup = buildTeacherLookupFromCodeMap(teacherByCode);
        const rows = enrichBelegungRows(rawRows, lookup, ownerByGroupId);
        const withDisplayPipe = groups.filter((g) => String(g.displayName || '').includes(' | '));

        importCandidates = rows.map((row) => ({
            id: row.graphGroupId,
            displayName: row.teamName,
            mailNickname: row.gruppenmail,
            row,
            selected: true
        }));

        if (!importCandidates.length) {
            toast(
                'Keine Kursteams mit Präfix „' +
                    yearPrefix +
                    '“ gefunden (' +
                    groups.length +
                    ' Teams im Tenant, ' +
                    withDisplayPipe.length +
                    ' mit „ | “ im Namen). Präfix anpassen?'
            );
        } else {
            toast(importCandidates.length + ' Kursteam(s) für Import vorbereitet.');
        }
        renderImportTable();
        runAbgleich();
        renderImportStatusSection();
        const saveBtn = document.getElementById('kursteamGraphImportSave');
        if (saveBtn) saveBtn.disabled = importCandidates.length === 0;
        const selectedCount = importCandidates.filter((c) => c.selected).length;
        if (selectedCount > 0) {
            saveSelectedToApp({ mentionCatalog: true });
        }
    } catch (e) {
        toast('Graph-Import: ' + (e && e.message ? e.message : e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

function saveSelectedToApp(options) {
    const opts = options && typeof options === 'object' ? options : {};
    const api = window.ms365AppDataV2;
    if (!api || typeof api.getUnterrichtsbelegung !== 'function' || typeof api.setUnterrichtsbelegung !== 'function') {
        toast('App-Daten nicht verfügbar – Seite neu laden oder Stammdaten öffnen.');
        return false;
    }
    const selected = importCandidates.filter((c) => c.selected).map((c) => c.row);
    if (!selected.length) {
        toast('Bitte mindestens ein Team auswählen.');
        return false;
    }
    const yearPrefix = getYearPrefixFromUi();
    const existing = api.getUnterrichtsbelegung();
    const { snapshot, stats } = mergeBelegungWithGraphImport(existing, selected, {
        yearPrefix,
        source: existing && existing.rows && existing.rows.length ? 'kursteams+graph-import' : 'graph-import'
    });
    api.setUnterrichtsbelegung(snapshot);
    const s = summarizeBelegung(snapshot);
    toast(
        'Unterrichtsbelegung gespeichert: ' +
            stats.added +
            ' neu, ' +
            stats.updated +
            ' aktualisiert · ' +
            stats.linkedCount +
            ' mit M365-ID · ' +
            s.rows +
            ' Einträge gesamt.' +
            (opts.mentionCatalog ? ' Sichtbar im Unterrichtsteams-Katalog.' : '')
    );
    if (typeof ns.updateUnterrichtsbelegungHint === 'function') ns.updateUnterrichtsbelegungHint();
    renderImportTable();
    runAbgleich();
    renderImportStatusSection();
    return true;
}

function showImportPanel() {
    const panel = document.getElementById('kursteamGraphImportPanel');
    if (panel) panel.style.display = 'block';
    const yp = document.getElementById('kursteamImportYearPrefix');
    if (yp && !yp.value) {
        yp.value = getYearPrefixFromUi();
    }
    if (typeof ns.updateUnterrichtsbelegungHint === 'function') ns.updateUnterrichtsbelegungHint();
    renderImportStatusSection();
}

function wireImportUi() {
    const openBtn = document.getElementById('kursteamGraphImportOpen');
    const loadBtn = document.getElementById('kursteamGraphImportLoad');
    const saveBtn = document.getElementById('kursteamGraphImportSave');
    const statusBtn =
        document.getElementById('kursteamGraphImportStatusRefresh') ||
        document.getElementById('kursteamGraphAbgleichRefresh');
    const selAll = document.getElementById('kursteamGraphImportSelectAll');
    const selNone = document.getElementById('kursteamGraphImportSelectNone');

    if (openBtn) openBtn.addEventListener('click', showImportPanel);
    if (statusBtn) {
        statusBtn.addEventListener('click', () => {
            runAbgleich();
            renderImportStatusSection();
            toast('Status aktualisiert.');
        });
    }
    if (loadBtn) loadBtn.addEventListener('click', () => loadGraphKursteams());
    if (saveBtn) {
        saveBtn.disabled = true;
        saveBtn.addEventListener('click', () => saveSelectedToApp());
    }
    if (selAll) {
        selAll.addEventListener('click', () => {
            importCandidates.forEach((c) => {
                c.selected = true;
            });
            renderImportTable();
        });
    }
    if (selNone) {
        selNone.addEventListener('click', () => {
            importCandidates.forEach((c) => {
                c.selected = false;
            });
            renderImportTable();
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wireImportUi);
} else {
    wireImportUi();
}

ns.startKursteamGraphImport = showImportPanel;
window.startKursteamGraphImport = showImportPanel;

window.ms365KursteamGraphImport = {
    parseKursteamDisplayName,
    loadGraphKursteams,
    saveSelectedToApp,
    runAbgleich,
    buildKursteamAbgleichReport
};
