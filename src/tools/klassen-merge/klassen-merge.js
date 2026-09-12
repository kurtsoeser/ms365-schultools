/**
 * Klassen-Merge Wizard – Wiring.
 */
import { buildMergePlan, applyLocalMerge } from './klassen-merge-logic.js';

function $(id) {
    return document.getElementById(id);
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function loadSettings() {
    return typeof window.ms365TenantSettingsLoad === 'function' ? window.ms365TenantSettingsLoad() : null;
}

function saveSettings(next) {
    if (typeof window.ms365TenantSettingsSave === 'function') window.ms365TenantSettingsSave(next);
}

function loadClassTeams() {
    try {
        const c = window.ms365AppDataV2 && window.ms365AppDataV2.getContainer && window.ms365AppDataV2.getContainer();
        return (c && c.core && Array.isArray(c.core.classTeams) ? c.core.classTeams : []).slice();
    } catch {
        return [];
    }
}

function saveClassTeams(teams) {
    try {
        const api = window.ms365AppDataV2;
        if (!api || typeof api.getContainer !== 'function' || typeof api.setContainer !== 'function') return;
        const c = api.getContainer();
        if (!c.core) c.core = {};
        c.core.classTeams = Array.isArray(teams) ? teams : [];
        api.setContainer(c);
    } catch {
        /* ignore */
    }
}

function gug() {
    return window.ms365GraphUnifiedGroups;
}

function fillClassSelects() {
    const settings = loadSettings() || {};
    const classes = Array.isArray(settings.classes) ? settings.classes.slice() : [];
    classes.sort((a, b) => String(a.code || '').localeCompare(String(b.code || ''), 'de'));
    const surv = $('kmSurvivor');
    const src = $('kmSources');
    if (!surv || !src) return classes;
    surv.replaceChildren();
    src.replaceChildren();
    const opt0 = document.createElement('option');
    opt0.value = '';
    opt0.textContent = '(wählen)';
    surv.appendChild(opt0);
    classes.forEach(function (c) {
        const label = (c.code || '') + (c.name && c.name !== c.code ? ' – ' + c.name : '');
        const o1 = document.createElement('option');
        o1.value = String(c.code || '');
        o1.textContent = label;
        surv.appendChild(o1);
        const o2 = document.createElement('option');
        o2.value = String(c.code || '');
        o2.textContent = label;
        src.appendChild(o2);
    });
    return classes;
}

function selectedSourceCodes() {
    const src = $('kmSources');
    if (!src) return [];
    return Array.from(src.selectedOptions)
        .map((o) => String(o.value || '').trim())
        .filter(Boolean);
}

function currentPlan() {
    const settings = loadSettings() || {};
    const classes = Array.isArray(settings.classes) ? settings.classes : [];
    const byCode = new Map(classes.map((c) => [String(c.code || '').toUpperCase().replace(/\s+/g, ''), c]));
    const survivorCode = String(($('kmSurvivor') && $('kmSurvivor').value) || '')
        .toUpperCase()
        .replace(/\s+/g, '');
    const survivor = byCode.get(survivorCode);
    const sources = selectedSourceCodes()
        .filter((c) => c !== survivorCode)
        .map((c) => byCode.get(c))
        .filter(Boolean);
    return buildMergePlan({
        survivor,
        survivorOriginalCode: survivorCode,
        sources,
        newCode: ($('kmNewCode') && $('kmNewCode').value) || survivorCode,
        newName: ($('kmNewName') && $('kmNewName').value) || '',
        newDisplayName: ($('kmNewDisplay') && $('kmNewDisplay').value) || '',
        sourceAction: ($('kmSourceAction') && $('kmSourceAction').value) || 'archive',
        students: settings.students || [],
        classTeams: loadClassTeams()
    });
}

function renderPreview() {
    const box = $('kmPreview');
    const warn = $('kmWarnings');
    if (!box) return;
    const plan = currentPlan();
    if (!plan.ok) {
        box.textContent = plan.error || 'Plan unvollständig.';
        if (warn) warn.textContent = '';
        return;
    }
    box.replaceChildren();
    const ul = document.createElement('ul');
    plan.steps.forEach(function (s) {
        const li = document.createElement('li');
        li.textContent = s.label + (s.count != null ? ' (' + s.count + ')' : '');
        ul.appendChild(li);
    });
    box.appendChild(ul);
    const p = document.createElement('p');
    p.textContent = 'Schüler-E-Mails in der Merge-Gruppe: ' + plan.memberEmails.length;
    box.appendChild(p);
    if (warn) {
        warn.textContent = (plan.warnings || []).join(' ');
    }
}

async function runMerge() {
    const plan = currentPlan();
    if (!plan.ok) {
        toast(plan.error || 'Plan ungültig');
        return;
    }
    const ok = window.confirm(
        'Klassen zusammenführen zu „' +
            plan.newCode +
            '“?\n\nLokale Stammdaten werden geändert' +
            (plan.survivorTeam && plan.survivorTeam.graphGroupId ? ' und Microsoft 365 aktualisiert' : '') +
            '.'
    );
    if (!ok) return;

    const log = $('kmLog');
    if (log) log.textContent = '';
    function write(m) {
        if (!log) return;
        log.textContent += (log.textContent ? '\n' : '') + m;
    }

    const settings = loadSettings() || {};
    const applied = applyLocalMerge(plan, settings, loadClassTeams());
    if (applied.error) {
        toast(applied.error);
        return;
    }
    saveSettings(applied.settings);
    saveClassTeams(applied.classTeams);
    write('Stammdaten aktualisiert.');

    const G = gug();
    if (G && plan.survivorTeam && plan.survivorTeam.graphGroupId) {
        try {
            const token = await G.getGraphToken([
                'https://graph.microsoft.com/User.Read',
                'https://graph.microsoft.com/Group.ReadWrite.All',
                'https://graph.microsoft.com/Directory.Read.All'
            ]);
            if (plan.memberEmails.length && typeof G.syncEmailsToGroup === 'function') {
                write('Mitglieder abgleichen …');
                await G.syncEmailsToGroup(token, plan.survivorTeam.graphGroupId, plan.memberEmails, 'Merge', write);
            }
            if (plan.newDisplayName && typeof G.patchGroupDisplayName === 'function') {
                write('Anzeigename setzen …');
                await G.patchGroupDisplayName(token, plan.survivorTeam.graphGroupId, plan.newDisplayName);
            }
            for (let i = 0; i < plan.steps.length; i++) {
                const st = plan.steps[i];
                if (!st.groupId || !String(st.id).startsWith('graph-source-')) continue;
                if (st.action === 'keep') continue;
                if (st.action === 'archive') {
                    write('Archiviere Team ' + st.classCode + ' …');
                    try {
                        await G.graphJson(
                            'POST',
                            '/teams/' + encodeURIComponent(st.groupId) + '/archive',
                            token,
                            { shouldSetSpoSiteReadOnlyForMembers: false }
                        );
                        write('Archiviert: ' + st.classCode);
                    } catch (e) {
                        write('Archiv fehlgeschlagen (' + st.classCode + '): ' + (e.message || e) + ' – ggf. kein Team.');
                    }
                } else if (st.action === 'delete' && typeof G.deleteUnifiedGroup === 'function') {
                    write('Lösche Gruppe ' + st.classCode + ' …');
                    await G.deleteUnifiedGroup(token, st.groupId);
                }
            }
        } catch (e) {
            write('Graph-Fehler: ' + (e && e.message ? e.message : e));
            toast('Lokal gespeichert, Graph teilweise fehlgeschlagen.');
            fillClassSelects();
            return;
        }
    }

    write('Fertig.');
    toast('Zusammenführung abgeschlossen.');
    fillClassSelects();
    renderPreview();
}

function boot() {
    fillClassSelects();
    ['kmSurvivor', 'kmSources', 'kmNewCode', 'kmNewName', 'kmNewDisplay', 'kmSourceAction'].forEach(function (id) {
        const el = $(id);
        if (!el) return;
        el.addEventListener('change', renderPreview);
        el.addEventListener('input', renderPreview);
    });
    const surv = $('kmSurvivor');
    if (surv) {
        surv.addEventListener('change', function () {
            const v = surv.value;
            if ($('kmNewCode') && !$('kmNewCode').value) $('kmNewCode').value = v;
            if ($('kmNewName') && !$('kmNewName').value) {
                const opt = surv.selectedOptions[0];
                $('kmNewName').value = opt ? opt.textContent.split('–')[0].trim() : v;
            }
            renderPreview();
        });
    }
    const btn = $('kmBtnRun');
    if (btn) btn.addEventListener('click', () => runMerge().catch((e) => toast(String(e && e.message ? e.message : e))));
    const reload = $('kmBtnReload');
    if (reload) reload.addEventListener('click', () => { fillClassSelects(); renderPreview(); toast('Klassen neu geladen.'); });
    renderPreview();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
