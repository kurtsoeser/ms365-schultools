/**
 * Namenskonvention-Audit UI
 * Pilot Phase A (Analyse 02): ausschließlich shared/graph-client.js – keine lokale PCA.
 */
import { analyzeUsersNaming, buildAdExportCsv } from '../../shared/naming-convention-audit.js';
import { escapeHtml } from '../../shared/utils/strings.js';
import { getGraphToken, graphJson, fetchAllPages } from '../../shared/graph-client.js';

function $(id) {
    return document.getElementById(id);
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

/** @type {ReturnType<typeof analyzeUsersNaming>|null} */
let lastResult = null;
/** @type {object[]} */
let loadedUsers = [];

const READ_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/Directory.Read.All'
];

const WRITE_SCOPES = [
    'https://graph.microsoft.com/User.ReadWrite.All',
    'https://graph.microsoft.com/Directory.AccessAsUser.All'
];

async function loadUsers() {
    const token = await getGraphToken(READ_SCOPES);
    const select =
        'id,displayName,givenName,surname,userPrincipalName,mail,accountEnabled,onPremisesSyncEnabled';
    const path = '/users?$select=' + encodeURIComponent(select) + '&$top=999';
    const page = await fetchAllPages(token, path, {
        maxItems: 8000,
        maxPages: 40,
        onProgress: function (info) {
            if ($('naStatus')) $('naStatus').textContent = 'Geladen: ' + info.loaded + ' …';
        }
    });
    if (page.truncated && $('naStatus')) {
        $('naStatus').textContent =
            'Geladen: ' + page.items.length + ' (Liste gekürzt – Truncation-Limit)';
    }
    loadedUsers = page.items.filter((u) => u && u.accountEnabled !== false);
    return loadedUsers;
}

function currentRules() {
    return {
        displayOrder: ($('naOrder') && $('naOrder').value) || 'given-sur',
        allowNumericUpn: !($('naNoNumericUpn') && $('naNoNumericUpn').checked)
    };
}

function render() {
    const filter = ($('naFilter') && $('naFilter').value) || 'problems';
    const result = analyzeUsersNaming(loadedUsers, currentRules());
    lastResult = result;
    if ($('naSummary')) {
        $('naSummary').textContent =
            result.summary.total +
            ' Benutzer · OK ' +
            result.summary.ok +
            ' · Cloud fixbar ' +
            result.summary.cloud_fix +
            ' · AD-Export ' +
            result.summary.ad_export;
    }
    const body = $('naBody');
    if (!body) return;
    body.replaceChildren();
    result.rows
        .filter(function (r) {
            if (filter === 'all') return true;
            if (filter === 'cloud') return r.severity === 'cloud_fix';
            if (filter === 'ad') return r.severity === 'ad_export';
            return r.severity !== 'ok';
        })
        .forEach(function (r) {
            const tr = document.createElement('tr');
            const issues = (r.issues || []).map((i) => i.message).join(' · ');
            tr.innerHTML =
                '<td>' +
                escapeHtml(r.displayName) +
                '</td><td>' +
                escapeHtml(r.userPrincipalName) +
                '</td><td>' +
                escapeHtml(r.expectedDisplay) +
                '</td><td>' +
                escapeHtml(r.severity) +
                '</td><td>' +
                escapeHtml(issues) +
                '</td><td></td>';
            const td = tr.lastElementChild;
            if (r.action === 'patch') {
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn btn-small';
                btn.textContent = 'Cloud fixen';
                btn.addEventListener('click', function () {
                    patchOne(r).catch((e) => toast(e.message || String(e)));
                });
                td.appendChild(btn);
            } else if (r.action === 'export') {
                td.textContent = 'nur AD';
            } else {
                td.textContent = '—';
            }
            body.appendChild(tr);
        });
}

async function patchOne(row) {
    const token = await getGraphToken(WRITE_SCOPES);
    const body = row.patch || {};
    if (!body.displayName) throw new Error('Kein Patch möglich');
    await graphJson('PATCH', '/users/' + encodeURIComponent(row.id), token, body);
    const u = loadedUsers.find((x) => x.id === row.id);
    if (u) {
        u.displayName = body.displayName;
        if (body.givenName) u.givenName = body.givenName;
        if (body.surname) u.surname = body.surname;
    }
    toast('Aktualisiert: ' + body.displayName);
    render();
}

async function patchAllCloud() {
    if (!lastResult) return;
    const list = lastResult.rows.filter((r) => r.action === 'patch');
    if (!list.length) {
        toast('Keine cloud-fixbaren Einträge.');
        return;
    }
    if (!window.confirm(list.length + ' Cloud-Benutzer korrigieren?')) return;
    for (let i = 0; i < list.length; i++) {
        await patchOne(list[i]);
    }
}

function downloadAdCsv() {
    if (!lastResult) {
        toast('Zuerst prüfen.');
        return;
    }
    const csv = buildAdExportCsv(lastResult.rows);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'namenskonvention-ad-export.csv';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 500);
}

function boot() {
    $('naBtnScan') &&
        $('naBtnScan').addEventListener('click', function () {
            loadUsers()
                .then(function () {
                    render();
                    toast('Scan fertig.');
                })
                .catch(function (e) {
                    toast(e.message || String(e));
                });
        });
    $('naFilter') && $('naFilter').addEventListener('change', render);
    $('naOrder') && $('naOrder').addEventListener('change', render);
    $('naNoNumericUpn') && $('naNoNumericUpn').addEventListener('change', render);
    $('naBtnFixCloud') &&
        $('naBtnFixCloud').addEventListener('click', () => patchAllCloud().catch((e) => toast(e.message || String(e))));
    $('naBtnExportAd') && $('naBtnExportAd').addEventListener('click', downloadAdCsv);
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
