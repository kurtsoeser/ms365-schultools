/**
 * Namenskonvention-Audit UI
 * Pilot Phase A (Analyse 02): ausschließlich shared/graph-client.js – keine lokale PCA.
 */
import {
    analyzeUsersNaming,
    buildAdExportCsv,
    namingAuditLicenseGroupKey,
    namingAuditLicenseGroupLabel,
    namingAuditMatchesUserTypeFilter,
    namingAuditMembershipKind,
    namingAuditMembershipLabel
} from '../../shared/naming-convention-audit.js';
import { summarizeUserLicenses, userMatchesLicenseFilter } from '../../shared/graph-licenses.js';
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
/** @type {Map<string, {skuPartNumber?: string}>|null} */
let skuLookup = null;

const READ_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/Directory.Read.All',
    'https://graph.microsoft.com/Organization.Read.All'
];

const WRITE_SCOPES = [
    'https://graph.microsoft.com/User.ReadWrite.All',
    'https://graph.microsoft.com/Directory.AccessAsUser.All'
];

function buildSkuLookup(skus) {
    const map = new Map();
    (Array.isArray(skus) ? skus : []).forEach(function (s) {
        const id = String((s && s.skuId) || '').toLowerCase();
        if (!id) return;
        map.set(id, {
            skuId: id,
            skuPartNumber: String((s && s.skuPartNumber) || '')
        });
    });
    return map;
}

async function loadSubscribedSkus(token) {
    skuLookup = null;
    const G = window.ms365GraphUnifiedGroups;
    if (G && typeof G.fetchSubscribedSkus === 'function') {
        const res = await G.fetchSubscribedSkus(token);
        if (res && res.ok && typeof G.skuLookupFromSubscribed === 'function') {
            skuLookup = G.skuLookupFromSubscribed(res.skus);
            return;
        }
    }
    try {
        const data = await graphJson(
            'GET',
            '/subscribedSkus?$select=' +
                encodeURIComponent('skuId,skuPartNumber,prepaidUnits,consumedUnits,capabilityStatus'),
            token,
            undefined
        );
        skuLookup = buildSkuLookup(data.value);
    } catch {
        skuLookup = null;
    }
}

async function loadUsers() {
    const token = await getGraphToken(READ_SCOPES);
    const select =
        'id,displayName,givenName,surname,userPrincipalName,mail,accountEnabled,onPremisesSyncEnabled,userType,assignedLicenses';
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
    if ($('naStatus')) $('naStatus').textContent = 'Lizenzen des Mandanten lesen …';
    await loadSubscribedSkus(token);
    return loadedUsers;
}

function currentRules() {
    return {
        displayOrder: ($('naOrder') && $('naOrder').value) || 'given-sur',
        allowNumericUpn: !($('naNoNumericUpn') && $('naNoNumericUpn').checked)
    };
}

function userForRow(row) {
    if (!row || !row.id) return null;
    return loadedUsers.find(function (u) {
        return u && u.id === row.id;
    });
}

function passesSeverityFilter(r, filter) {
    if (filter === 'all') return true;
    if (filter === 'cloud') return r.severity === 'cloud_fix';
    if (filter === 'ad') return r.severity === 'ad_export';
    return r.severity !== 'ok';
}

function passesAudienceFilters(row) {
    const u = userForRow(row);
    const typeVal = ($('naFilterUserType') && $('naFilterUserType').value) || '';
    const licVal = ($('naFilterLicense') && $('naFilterLicense').value) || '';
    if (!namingAuditMatchesUserTypeFilter(u || {}, typeVal)) return false;
    if (licVal && !userMatchesLicenseFilter(u || {}, licVal, skuLookup)) return false;
    return true;
}

function groupSortKey(mode, row) {
    const u = userForRow(row);
    const sum = summarizeUserLicenses(u || {}, skuLookup);
    if (mode === 'kontotyp') {
        const k = namingAuditMembershipKind(u || {});
        if (k === 'member') return '0';
        if (k === 'guest') return '1';
        return '2';
    }
    if (mode === 'lizenz') {
        const g = namingAuditLicenseGroupKey(sum);
        if (g === 'faculty') return '0';
        if (g === 'student') return '1';
        if (g === 'mixed') return '2';
        if (g === 'other') return '3';
        return '4';
    }
    return '';
}

function groupLabel(mode, row) {
    const u = userForRow(row);
    const sum = summarizeUserLicenses(u || {}, skuLookup);
    if (mode === 'kontotyp') return namingAuditMembershipLabel(namingAuditMembershipKind(u || {}));
    if (mode === 'lizenz') return namingAuditLicenseGroupLabel(namingAuditLicenseGroupKey(sum));
    return '';
}

function appendDataRow(body, r) {
    const u = userForRow(r);
    const sum = summarizeUserLicenses(u || {}, skuLookup);
    const membership = namingAuditMembershipLabel(namingAuditMembershipKind(u || {}));
    const license = sum && sum.hasAny ? sum.primaryLabel : '—';
    const issues = (r.issues || []).map((i) => i.message).join(' · ');
    const tr = document.createElement('tr');
    tr.innerHTML =
        '<td>' +
        escapeHtml(r.displayName) +
        '</td><td>' +
        escapeHtml(r.userPrincipalName) +
        '</td><td>' +
        escapeHtml(membership) +
        '</td><td>' +
        escapeHtml(license) +
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
}

function render() {
    const filter = ($('naFilter') && $('naFilter').value) || 'problems';
    const groupBy = ($('naGroupBy') && $('naGroupBy').value) || 'none';
    const result = analyzeUsersNaming(loadedUsers, currentRules());
    lastResult = result;
    let rows = result.rows.filter(function (r) {
        return passesSeverityFilter(r, filter) && passesAudienceFilters(r);
    });
    if (groupBy !== 'none') {
        rows = rows.slice().sort(function (a, b) {
            const ga = groupSortKey(groupBy, a);
            const gb = groupSortKey(groupBy, b);
            if (ga !== gb) return ga < gb ? -1 : 1;
            return String(a.displayName || '').localeCompare(String(b.displayName || ''), 'de', {
                sensitivity: 'base'
            });
        });
    }
    const typeActive = $('naFilterUserType') && $('naFilterUserType').value;
    const licActive = $('naFilterLicense') && $('naFilterLicense').value;
    if ($('naSummary')) {
        let line =
            result.summary.total +
            ' Benutzer · OK ' +
            result.summary.ok +
            ' · Cloud fixbar ' +
            result.summary.cloud_fix +
            ' · AD-Export ' +
            result.summary.ad_export;
        if (filter !== 'all' || typeActive || licActive) {
            line += ' · Angezeigt ' + rows.length;
        }
        $('naSummary').textContent = line;
    }
    const body = $('naBody');
    if (!body) return;
    body.replaceChildren();
    let lastGroup = '';
    rows.forEach(function (r) {
        if (groupBy !== 'none') {
            const label = groupLabel(groupBy, r);
            if (label && label !== lastGroup) {
                lastGroup = label;
                const gr = document.createElement('tr');
                gr.className = 'na-group-row';
                const td = document.createElement('td');
                td.colSpan = 8;
                td.textContent = label;
                gr.appendChild(td);
                body.appendChild(gr);
            }
        }
        appendDataRow(body, r);
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

function bindFilter(id) {
    const el = $(id);
    if (el) el.addEventListener('change', render);
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
    bindFilter('naFilter');
    bindFilter('naFilterUserType');
    bindFilter('naFilterLicense');
    bindFilter('naGroupBy');
    $('naOrder') && $('naOrder').addEventListener('change', render);
    $('naNoNumericUpn') && $('naNoNumericUpn').addEventListener('change', render);
    $('naBtnFixCloud') &&
        $('naBtnFixCloud').addEventListener('click', () => patchAllCloud().catch((e) => toast(e.message || String(e))));
    $('naBtnExportAd') && $('naBtnExportAd').addEventListener('click', downloadAdCsv);
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
