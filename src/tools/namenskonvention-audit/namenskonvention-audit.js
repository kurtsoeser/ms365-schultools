/**
 * Namenskonvention-Audit UI
 */
import { analyzeUsersNaming, buildAdExportCsv } from '../../shared/naming-convention-audit.js';

function $(id) {
    return document.getElementById(id);
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function gug() {
    return window.ms365GraphUnifiedGroups;
}

/** @type {ReturnType<typeof analyzeUsersNaming>|null} */
let lastResult = null;
/** @type {object[]} */
let loadedUsers = [];

async function loadUsers() {
    const G = gug();
    if (!G) throw new Error('Graph-Modul nicht geladen.');
    const token = await G.getGraphToken([
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/Directory.Read.All'
    ]);
    const select =
        'id,displayName,givenName,surname,userPrincipalName,mail,accountEnabled,onPremisesSyncEnabled';
    let next = '/users?$select=' + encodeURIComponent(select) + '&$top=999';
    const raw = [];
    let page = 0;
    while (next && page < 40 && raw.length < 8000) {
        page++;
        const data = await G.graphJson('GET', next, token, undefined);
        if (Array.isArray(data.value)) raw.push.apply(raw, data.value);
        next = data['@odata.nextLink'] || null;
        if ($('naStatus')) $('naStatus').textContent = 'Geladen: ' + raw.length + ' …';
    }
    loadedUsers = raw.filter((u) => u && u.accountEnabled !== false);
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

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

async function patchOne(row) {
    const G = gug();
    const token = await G.getGraphToken([
        'https://graph.microsoft.com/User.ReadWrite.All',
        'https://graph.microsoft.com/Directory.AccessAsUser.All'
    ]);
    const body = row.patch || {};
    if (!body.displayName) throw new Error('Kein Patch möglich');
    await G.graphJson('PATCH', '/users/' + encodeURIComponent(row.id), token, body);
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
    $('naBtnFixCloud') && $('naBtnFixCloud').addEventListener('click', () => patchAllCloud().catch((e) => toast(e.message || String(e))));
    $('naBtnExportAd') && $('naBtnExportAd').addEventListener('click', downloadAdCsv);
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
