/**
 * Spielwiesen-Team Generator + Notebook-Checkliste.
 */
import { buildSpielwiesenPlan } from './spielwiesen-logic.js';

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

function readForm() {
    return {
        label: ($('spLabel') && $('spLabel').value) || '',
        year: ($('spYear') && $('spYear').value) || '',
        asDemo: !!( $('spAsDemo') && $('spAsDemo').checked )
    };
}

function refresh() {
    const plan = buildSpielwiesenPlan(readForm());
    const prev = $('spPreview');
    if (prev) {
        prev.innerHTML =
            '<p><strong>' +
            escapeHtml(plan.displayName) +
            '</strong></p><p><code>' +
            escapeHtml(plan.mailNickname) +
            '</code></p><p class="muted">' +
            escapeHtml(plan.description) +
            '</p>';
    }
    const list = $('spNotebookList');
    if (list) {
        list.replaceChildren();
        plan.notebookChecklist.forEach(function (step, i) {
            const li = document.createElement('li');
            li.innerHTML =
                '<label style="display:flex;gap:8px;align-items:flex-start;"><input type="checkbox" data-sp-nb="' +
                i +
                '"><span>' +
                escapeHtml(step) +
                '</span></label>';
            list.appendChild(li);
        });
        restoreNb();
    }
}

function nbKey() {
    return 'ms365-spielwiesen-notebook-v1';
}

function restoreNb() {
    let state = {};
    try {
        state = JSON.parse(localStorage.getItem(nbKey()) || '{}') || {};
    } catch {
        state = {};
    }
    document.querySelectorAll('[data-sp-nb]').forEach(function (el) {
        el.checked = !!state[el.getAttribute('data-sp-nb')];
        el.addEventListener('change', function () {
            const s = {};
            document.querySelectorAll('[data-sp-nb]').forEach(function (x) {
                s[x.getAttribute('data-sp-nb')] = !!x.checked;
            });
            try {
                localStorage.setItem(nbKey(), JSON.stringify(s));
            } catch {
                /* ignore */
            }
        });
    });
}

async function createTeam() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G) throw new Error('Graph nicht geladen.');
    const plan = buildSpielwiesenPlan(readForm());
    if (!plan.ok) throw new Error(plan.issues.join(', '));
    if (!window.confirm('Spielwiesen-Team anlegen?\n\n' + plan.displayName)) return;
    const token = await G.getGraphToken([
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.ReadWrite.All'
    ]);
    const group = await G.createUnifiedGroup(token, plan.displayName, plan.mailNickname, plan.description);
    await G.provisionTeamForGroup(token, group.id);
    toast('Spielwiesen-Team angelegt.');
    const log = $('spLog');
    if (log) log.textContent = 'OK: ' + group.id + '\n' + plan.displayName;
}

function boot() {
    if ($('spYear') && !$('spYear').value) $('spYear').value = String(new Date().getFullYear());
    ['spLabel', 'spYear', 'spAsDemo'].forEach(function (id) {
        const el = $(id);
        if (el) {
            el.addEventListener('input', refresh);
            el.addEventListener('change', refresh);
        }
    });
    refresh();
    const btn = $('spBtnCreate');
    if (btn && btn.dataset.bound !== '1') {
        btn.dataset.bound = '1';
        btn.addEventListener('click', function () {
            createTeam().catch(function (e) {
                toast(e.message || String(e));
            });
        });
    }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
