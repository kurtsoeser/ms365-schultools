/**
 * Diplomarbeiten-Gruppen UI.
 */
import {
    buildDiplomPlan,
    filterDiplomGroups,
    isDiplomGroup
} from './diplomarbeiten-logic.js';

function $(id) {
    return document.getElementById(id);
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function log(msg) {
    const el = $('daLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(msg || '');
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function readForm() {
    return {
        year: ($('daYear') && $('daYear').value) || '',
        topic: ($('daTopic') && $('daTopic').value) || '',
        student: ($('daStudent') && $('daStudent').value) || '',
        mentor: ($('daMentor') && $('daMentor').value) || '',
        asTeam: !($('daGroupOnly') && $('daGroupOnly').checked)
    };
}

function refreshPreview() {
    const plan = buildDiplomPlan(readForm());
    const el = $('daPreview');
    if (!el) return;
    if (!plan.ok) {
        el.innerHTML = '<p class="muted">' + escapeHtml(plan.issues.join(' · ')) + '</p>';
        return;
    }
    el.innerHTML =
        '<p><strong>Anzeigename:</strong> ' +
        escapeHtml(plan.displayName) +
        '</p><p><strong>Alias:</strong> <code>' +
        escapeHtml(plan.mailNickname) +
        '</code></p><p class="muted">' +
        escapeHtml(plan.description) +
        '</p>';
}

async function createDiplom() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G) throw new Error('Graph-Modul nicht geladen.');
    const plan = buildDiplomPlan(readForm());
    if (!plan.ok) throw new Error(plan.issues.join(', '));
    if (!window.confirm('Gruppe anlegen?\n\n' + plan.displayName + '\n' + plan.mailNickname)) return;

    const token = await G.getGraphToken([
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.ReadWrite.All'
    ]);
    log('Lege Gruppe an …');
    const group = await G.createUnifiedGroup(token, plan.displayName, plan.mailNickname, plan.description);
    log('Gruppe: ' + (group.id || ''));
    if (plan.asTeam) {
        log('Provisioniere Team …');
        await G.provisionTeamForGroup(token, group.id);
        log('Team bereit.');
    }
    toast('Diplomarbeit-Gruppe angelegt.');
    await loadDiplomList();
}

async function loadDiplomList() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G) throw new Error('Graph-Modul nicht geladen.');
    const status = $('daListStatus');
    if (status) status.textContent = 'Lade Gruppen …';
    const token = await G.getGraphToken([
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.Read.All',
        'https://graph.microsoft.com/Directory.Read.All'
    ]);
    let next = '/groups?$select=id,displayName,mailNickname,mail,description&$top=999';
    const all = [];
    let page = 0;
    while (next && page < 30) {
        page++;
        const data = await G.graphJson('GET', next, token, undefined);
        if (Array.isArray(data.value)) all.push.apply(all, data.value);
        next = data['@odata.nextLink'] || null;
        if (status) status.textContent = 'Gruppen: ' + all.length;
    }
    const q = ($('daFilter') && $('daFilter').value) || '';
    const rows = filterDiplomGroups(all, q);
    const body = $('daBody');
    if (body) {
        body.replaceChildren();
        if (!rows.length) {
            body.innerHTML = '<tr><td colspan="3" class="muted">Keine Diplom-Gruppen gefunden (Präfix dipl- / „Diplomarbeit …“).</td></tr>';
        } else {
            rows.forEach(function (g) {
                const tr = document.createElement('tr');
                tr.innerHTML =
                    '<td>' +
                    escapeHtml(g.displayName) +
                    '</td><td><code>' +
                    escapeHtml(g.mailNickname || '') +
                    '</code></td><td>' +
                    escapeHtml(g.mail || '') +
                    '</td>';
                body.appendChild(tr);
            });
        }
    }
    if (status) {
        status.textContent =
            rows.length + ' Diplom-Gruppe(n) von ' + all.length + ' gesamt · Treffer-Regel: ' + (isDiplomGroup({ mailNickname: 'dipl-x' }) ? 'dipl-' : '');
    }
    toast(rows.length + ' Diplom-Gruppen');
}

function boot() {
    if ($('daYear') && !$('daYear').value) $('daYear').value = String(new Date().getFullYear() + 1);
    ['daYear', 'daTopic', 'daStudent', 'daMentor', 'daGroupOnly'].forEach(function (id) {
        const el = $(id);
        if (el) el.addEventListener('input', refreshPreview);
        if (el) el.addEventListener('change', refreshPreview);
    });
    refreshPreview();
    const create = $('daBtnCreate');
    if (create && create.dataset.bound !== '1') {
        create.dataset.bound = '1';
        create.addEventListener('click', function () {
            createDiplom().catch(function (e) {
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });
    }
    const list = $('daBtnList');
    if (list && list.dataset.bound !== '1') {
        list.dataset.bound = '1';
        list.addEventListener('click', function () {
            loadDiplomList().catch(function (e) {
                toast(e.message || String(e));
            });
        });
    }
    const filter = $('daFilter');
    if (filter && filter.dataset.bound !== '1') {
        filter.dataset.bound = '1';
        filter.addEventListener('input', function () {
            /* list already loaded – re-filter would need cache; trigger reload on Enter via button */
        });
    }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
