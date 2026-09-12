/**
 * Stammdaten-Health Panel für Datenhygiene-Seite.
 */
import { diffStudentAttributes } from './stammdaten-health.js';

function $(id) {
    return document.getElementById(id);
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

async function runAttributeScan() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G) throw new Error('Graph nicht geladen.');
    const settings = typeof window.ms365TenantSettingsLoad === 'function' ? window.ms365TenantSettingsLoad() : null;
    const students = (settings && settings.students) || [];
    if (!students.length) throw new Error('Keine Schüler in den Stammdaten.');

    const status = $('shStatus');
    if (status) status.textContent = 'Lade Benutzer …';
    const token = await G.getGraphToken([
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/Directory.Read.All'
    ]);
    const select = 'id,displayName,mail,userPrincipalName,department,officeLocation,otherMails';
    let next = '/users?$select=' + encodeURIComponent(select) + '&$top=999';
    const users = [];
    let page = 0;
    while (next && page < 40 && users.length < 8000) {
        page++;
        const data = await G.graphJson('GET', next, token, undefined);
        if (Array.isArray(data.value)) users.push.apply(users, data.value);
        next = data['@odata.nextLink'] || null;
        if (status) status.textContent = 'Benutzer: ' + users.length;
    }

    const diff = diffStudentAttributes(students, users);
    if (status) {
        status.textContent =
            'Verglichen: ' +
            diff.summary.matched +
            ' · OK ' +
            diff.summary.ok +
            ' · Abweichungen ' +
            diff.summary.mismatch +
            ' · ohne Graph ' +
            diff.summary.missingInGraph;
    }
    const body = $('shBody');
    if (!body) return;
    body.replaceChildren();
    const problems = diff.rows.filter((r) => r.status !== 'ok');
    if (!problems.length) {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td colspan="5" class="muted">Keine Abweichungen gefunden.</td>';
        body.appendChild(tr);
        return;
    }
    problems.slice(0, 500).forEach(function (r) {
        const tr = document.createElement('tr');
        tr.innerHTML =
            '<td>' +
            escapeHtml(r.name) +
            '</td><td>' +
            escapeHtml(r.email) +
            '</td><td>' +
            escapeHtml(r.localClass) +
            '</td><td>' +
            escapeHtml(r.graphClass) +
            '</td><td>' +
            escapeHtml(r.message) +
            '</td>';
        body.appendChild(tr);
    });
}

export function mountStammdatenHealth() {
    const btn = $('shScan');
    if (!btn || btn.dataset.bound === '1') return;
    btn.dataset.bound = '1';
    btn.addEventListener('click', function () {
        runAttributeScan().catch(function (e) {
            toast(e.message || String(e));
        });
    });
}

if (typeof window !== 'undefined') {
    window.ms365StammdatenHealthUi = { mount: mountStammdatenHealth };
}
