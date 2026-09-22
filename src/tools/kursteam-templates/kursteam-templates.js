/**
 * Kursteam-Vorlagen – UI (Verwaltung + Anwenden).
 */
import {
    createEmptyTemplate,
    cloneTemplate,
    filterTemplates,
    renameTemplateChannel,
    addTemplateChannel,
    removeTemplateChannel,
    moveTemplateChannel,
    normalizeTemplate,
    diffChannels,
    summarizeDiff,
    buildExportPayload,
    parseImportPayload,
    mergeTemplates,
    subjectOptionsFromCore,
    isGeneralChannelName,
    buildTemplateTree,
    schulstufeLabel,
    semesterLabel,
    collectSchoolForms,
    rememberSchoolForm,
    renameSchoolFormInTemplates,
    normalizeSchoolForm,
    mergeCatalogView,
    templateContentKey,
    localTemplatesForCatalogMerge
} from './kursteam-templates-logic.js';
import { getSeedTemplates } from './kursteam-templates-seed.js';
import {
    loadState,
    saveTemplates,
    saveState,
    upsertTemplate,
    deleteTemplate,
    resetToSeedTemplates
} from './kursteam-templates-storage.js';
import { searchTeams, listChannels, applyDiff } from './kursteam-templates-graph.js';
import { fetchCentralCatalog } from './kursteam-templates-catalog.js';

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
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function log(msg) {
    const el = $('ktplLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(msg || '');
    el.scrollTop = el.scrollHeight;
}

const ui = {
    /** @type {import('./kursteam-templates-logic.js').ChannelTemplate[]} */
    templates: [],
    /** @type {import('./kursteam-templates-logic.js').ChannelTemplate[]} */
    localTemplates: [],
    /** @type {import('./kursteam-templates-logic.js').ChannelTemplate[]} */
    centralTemplates: [],
    /** @type {string[]} */
    schoolForms: [],
    /** @type {string[]} */
    localSchoolForms: [],
    /** @type {string[]} */
    centralSchoolForms: [],
    /** @type {''|'central'|'local'} */
    sourceFilter: '',
    catalog: {
        ok: false,
        missing: true,
        message: 'Zentrale wird geladen …',
        updatedAt: null,
        updatedBy: '',
        webUrl: '',
        library: '',
        path: ''
    },
    selectedId: '',
    schoolFormFilter: '',
    subjectFilter: '',
    schulstufeFilter: '',
    semesterFilter: '',
    applySchoolFormFilter: '',
    applySubjectFilter: '',
    applySchulstufeFilter: '',
    applySemesterFilter: '',
    /** @type {'schoolForm'|'subject'|'schulstufe'|'semester'} */
    groupBy: 'schoolForm',
    /** @type {Set<string>} */
    expanded: new Set(),
    treeBootstrapped: false,
    /** @type {{ id: string, displayName: string, mailNickname: string }|null} */
    pickedTeam: null,
    /** @type {import('./kursteam-templates-logic.js').DiffRow[]|null} */
    lastDiff: null,
    applyTemplateId: ''
};

function currentFilters() {
    return {
        schoolForm: ui.schoolFormFilter,
        subjectCode: ui.subjectFilter,
        schulstufe: ui.schulstufeFilter,
        semester: ui.semesterFilter
    };
}

function applyFilters() {
    return {
        schoolForm: ui.applySchoolFormFilter,
        subjectCode: ui.applySubjectFilter,
        schulstufe: ui.applySchulstufeFilter,
        semester: ui.applySemesterFilter
    };
}

function loadSubjects() {
    try {
        const api = window.ms365AppDataV2;
        if (api && typeof api.getContainer === 'function') {
            const c = api.getContainer();
            const subjects = (c && c.core && c.core.subjects) || [];
            return subjectOptionsFromCore(subjects);
        }
    } catch {
        /* ignore */
    }
    return [];
}

function fillSelect(sel, options, emptyLabel, current) {
    if (!sel) return;
    const cur = current != null ? current : sel.value;
    sel.innerHTML = '';
    const first = document.createElement('option');
    first.value = '';
    first.textContent = emptyLabel;
    sel.appendChild(first);
    for (const o of options) {
        const opt = document.createElement('option');
        if (typeof o === 'string') {
            opt.value = o;
            opt.textContent = o;
        } else {
            opt.value = o.code;
            opt.textContent = o.label;
        }
        sel.appendChild(opt);
    }
    if ([...sel.options].some((o) => o.value === cur)) sel.value = cur;
}

function fillMetaSelects() {
    const schoolForms = collectSchoolForms(ui.templates, ui.schoolForms);
    fillSelect($('ktplFilterSchoolForm'), schoolForms, 'Alle Schulformen', ui.schoolFormFilter);
    fillSelect($('ktplApplySchoolForm'), schoolForms, 'Alle Schulformen', ui.applySchoolFormFilter);
    // Editor: aktuelle Auswahl merken, dann Liste neu füllen
    const editSchool = $('ktplEditSchoolForm');
    const editCur = editSchool ? editSchool.value : '';
    fillSelect($('ktplEditSchoolForm'), schoolForms, '— keine Schulform —', editCur);

    const subjects = loadSubjects();
    const fromTpl = new Map(subjects.map((s) => [s.code, s]));
    for (const t of ui.templates) {
        if (t.subjectCode && !fromTpl.has(t.subjectCode)) {
            fromTpl.set(t.subjectCode, { code: t.subjectCode, label: t.subjectCode });
        }
    }
    const subjOpts = Array.from(fromTpl.values()).sort((a, b) => a.code.localeCompare(b.code, 'de'));
    fillSelect($('ktplFilterSubject'), subjOpts, 'Alle Fächer', ui.subjectFilter);
    fillSelect($('ktplEditSubject'), subjOpts, '— kein Fach —', null);
    fillSelect($('ktplApplySubjectFilter'), subjOpts, 'Alle Fächer', ui.applySubjectFilter);

    const stufen = Array.from(
        new Set(ui.templates.map((t) => t.schulstufe).filter(Boolean))
    ).sort((a, b) => String(a).localeCompare(String(b), 'de', { numeric: true }));
    const stufeOpts = stufen.map((s) => ({ code: s, label: schulstufeLabel(s) }));
    fillSelect($('ktplFilterSchulstufe'), stufeOpts, 'Alle Schulstufen', ui.schulstufeFilter);
    fillSelect($('ktplApplySchulstufeFilter'), stufeOpts, 'Alle Schulstufen', ui.applySchulstufeFilter);

    const semesters = Array.from(
        new Set(ui.templates.map((t) => t.semester).filter(Boolean))
    ).sort((a, b) => String(a).localeCompare(String(b), 'de'));
    const semOpts = semesters.map((s) => ({ code: s, label: semesterLabel(s) }));
    fillSelect($('ktplFilterSemester'), semOpts, 'Alle Semester', ui.semesterFilter);
    fillSelect($('ktplApplySemesterFilter'), semOpts, 'Alle Semester', ui.applySemesterFilter);
}

function refreshFromStorage() {
    const state = loadState();
    ui.localTemplates = state.templates;
    ui.localSchoolForms = state.schoolForms || [];
    let forms = ui.localSchoolForms.slice();
    for (const f of ui.centralSchoolForms) forms = rememberSchoolForm(forms, f);
    for (const t of ui.centralTemplates) forms = rememberSchoolForm(forms, t.schoolForm);
    ui.schoolForms = forms;
    const localForMerge = localTemplatesForCatalogMerge(
        ui.localTemplates,
        ui.centralTemplates,
        getSeedTemplates()
    );
    ui.templates = mergeCatalogView(ui.centralTemplates, localForMerge);
}

function templatesForLibrary() {
    if (ui.sourceFilter === 'central') {
        return ui.templates.filter((t) => t.origin === 'central');
    }
    if (ui.sourceFilter === 'local') {
        return ui.templates.filter((t) => t.origin === 'local' || t.origin === 'override');
    }
    return ui.templates;
}

function originLabel(origin) {
    if (origin === 'central') return 'Zentral';
    if (origin === 'override') return 'Angepasst';
    if (origin === 'local') return 'Lokal';
    return '';
}

function formatCatalogStamp(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return '';
    return d.toLocaleString('de-AT', { dateStyle: 'short', timeStyle: 'short' });
}

function renderCatalogStatus() {
    const el = $('ktplCatalogStatus');
    if (!el) return;
    const centralN = ui.templates.filter((t) => t.origin === 'central').length;
    const overrideN = ui.templates.filter((t) => t.origin === 'override').length;
    const localN = ui.templates.filter((t) => t.origin === 'local').length;
    const parts = ['Zentral ' + centralN];
    if (overrideN) parts.push(overrideN + ' angepasst');
    if (localN) parts.push(localN + ' nur lokal');
    const stamp = formatCatalogStamp(ui.catalog.updatedAt);
    let text = parts.join(' · ');
    if (ui.catalog.message && !ui.catalog.ok) text = ui.catalog.message;
    else if (ui.catalog.missing && ui.catalog.ok) {
        text = 'Zentrale noch leer. ' + text;
    }
    if (stamp && ui.catalog.ok && !ui.catalog.missing) text += ' · Stand ' + stamp;
    el.textContent = text;
    const link = $('ktplCatalogLink');
    if (link) {
        link.hidden = false;
        link.href = ui.catalog.webUrl || 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools';
    }
}

function selectedTemplate() {
    return ui.templates.find((t) => t.id === ui.selectedId) || null;
}

function ensureExpandedForSelection() {
    const tpl = selectedTemplate();
    if (!tpl) return;
    if (ui.groupBy === 'schulstufe') {
        ui.expanded.add(tpl.schulstufe ? 'st:' + tpl.schulstufe : 'st:_none');
    } else if (ui.groupBy === 'semester') {
        ui.expanded.add(tpl.semester ? 'sem:' + tpl.semester : 'sem:_none');
    } else if (ui.groupBy === 'schoolForm') {
        ui.expanded.add(tpl.schoolForm ? 'sf:' + tpl.schoolForm : 'sf:_none');
    } else {
        ui.expanded.add(tpl.subjectCode ? 'subj:' + tpl.subjectCode : 'subj:_none');
    }
}

function bootstrapTreeExpanded() {
    if (ui.treeBootstrapped) return;
    const tree = buildTemplateTree(templatesForLibrary(), ui.groupBy, currentFilters());
    for (const g of tree) ui.expanded.add(g.key);
    ui.treeBootstrapped = true;
}

function renderTemplateList() {
    const ul = $('ktplTemplateList');
    const countEl = $('ktplListCount');
    if (!ul) return;
    bootstrapTreeExpanded();
    ensureExpandedForSelection();
    const visible = templatesForLibrary();
    const filtered = filterTemplates(visible, currentFilters());
    if (countEl) countEl.textContent = String(filtered.length);
    if (!filtered.length) {
        ul.innerHTML =
            '<li class="ktpl-empty">Keine Vorlagen für diese Filter.<br>Mit „Neu“ anlegen, JSON importieren oder die Zentrale laden.</li>';
        renderCatalogStatus();
        return;
    }

    const tree = buildTemplateTree(visible, ui.groupBy, currentFilters());
    let html = '';
    for (const group of tree) {
        const open = ui.expanded.has(group.key);
        html +=
            `<li class="ktpl-tree-node ktpl-tree-node--group">` +
            `<button type="button" class="ktpl-tree-toggle" data-toggle="${escapeHtml(group.key)}" aria-expanded="${open ? 'true' : 'false'}">` +
            `<i class="bi ${open ? 'bi-caret-down-fill' : 'bi-caret-right-fill'}" aria-hidden="true"></i>` +
            `<span class="ktpl-tree-label">${escapeHtml(group.label)}</span>` +
            `<span class="ktpl-tree-count">${group.count}</span>` +
            `</button>`;
        if (open) {
            html += `<ul class="ktpl-tree-children">`;
            for (const leaf of group.children || []) {
                const t = leaf.template;
                if (!t) continue;
                const active = t.id === ui.selectedId ? ' is-active' : '';
                const badge = leaf.badge
                    ? `<span class="ktpl-pill">${escapeHtml(leaf.badge)}</span>`
                    : '';
                const origin = originLabel(t.origin);
                const originPill = origin
                    ? `<span class="ktpl-origin ktpl-origin--${escapeHtml(t.origin || 'local')}">${escapeHtml(origin)}</span>`
                    : '';
                html +=
                    `<li><button type="button" class="ktpl-list-btn ktpl-list-btn--leaf${active}" data-id="${escapeHtml(t.id)}">` +
                    `<span class="ktpl-leaf-top"><strong>${escapeHtml(t.name)}</strong><span class="ktpl-leaf-pills">${originPill}${badge}</span></span>` +
                    `<span class="ktpl-list-meta">${t.channels.length} Kanal${t.channels.length === 1 ? '' : 'e'}</span>` +
                    `</button></li>`;
            }
            html += `</ul>`;
        }
        html += `</li>`;
    }
    ul.innerHTML = html;
    renderCatalogStatus();
}

function renderEditor() {
    const empty = $('ktplEditorEmpty');
    const form = $('ktplEditorForm');
    const countEl = $('ktplChannelCount');
    const tpl = selectedTemplate();
    if (!tpl) {
        if (empty) empty.hidden = false;
        if (form) form.hidden = true;
        if (countEl) countEl.hidden = true;
        return;
    }
    if (empty) empty.hidden = true;
    if (form) form.hidden = false;
    if (countEl) {
        countEl.hidden = false;
        countEl.textContent = tpl.channels.length + ' Kanal' + (tpl.channels.length === 1 ? '' : 'e');
    }

    const nameEl = $('ktplEditName');
    const schoolEl = $('ktplEditSchoolForm');
    const subjEl = $('ktplEditSubject');
    const stufeEl = $('ktplEditSchulstufe');
    const semEl = $('ktplEditSemester');
    const descEl = $('ktplEditDesc');
    if (nameEl) nameEl.value = tpl.name;
    if (schoolEl) {
        if (tpl.schoolForm && ![...schoolEl.options].some((o) => o.value === tpl.schoolForm)) {
            const opt = document.createElement('option');
            opt.value = tpl.schoolForm;
            opt.textContent = tpl.schoolForm;
            schoolEl.appendChild(opt);
        }
        schoolEl.value = tpl.schoolForm || '';
    }
    if (subjEl) {
        if (tpl.subjectCode && ![...subjEl.options].some((o) => o.value === tpl.subjectCode)) {
            const opt = document.createElement('option');
            opt.value = tpl.subjectCode;
            opt.textContent = tpl.subjectCode;
            subjEl.appendChild(opt);
        }
        subjEl.value = tpl.subjectCode || '';
    }
    if (stufeEl) stufeEl.value = tpl.schulstufe || '';
    if (semEl) semEl.value = tpl.semester || '';
    if (descEl) descEl.value = tpl.description || '';

    const note = $('ktplOriginNote');
    if (note) {
        if (tpl.origin === 'central') {
            note.hidden = false;
            note.textContent =
                'Zentrale Vorlage für alle Schulen. Speichern legt nur eine lokale Anpassung in diesem Browser an. Den Katalog für alle Schulen pflegst du im Admin.';
        } else if (tpl.origin === 'override') {
            note.hidden = false;
            note.textContent =
                'Lokale Anpassung einer zentralen Vorlage. Andere Schulen sehen weiter die Zentrale, bis der Betreiber neu veröffentlicht.';
        } else if (tpl.origin === 'local') {
            note.hidden = false;
            note.textContent = 'Nur in diesem Browser. Andere Schulen sehen diese Vorlage nicht.';
        } else {
            note.hidden = true;
            note.textContent = '';
        }
    }

    const list = $('ktplChannelBody');
    if (!list) return;
    if (!tpl.channels.length) {
        list.innerHTML =
            '<li class="ktpl-empty">Noch keine Kanäle. Oben einen Namen eingeben und hinzufügen.</li>';
        return;
    }
    list.innerHTML = tpl.channels
        .map((c, i) => {
            return (
                `<li class="ktpl-ch-row" data-ch="${escapeHtml(c.id)}">` +
                `<span class="ktpl-ch-ord">${i + 1}</span>` +
                `<input type="text" class="ktpl-ch-name" value="${escapeHtml(c.displayName)}" aria-label="Kanalname">` +
                `<span class="ktpl-ch-actions">` +
                `<button type="button" class="btn btn-ghost" data-act="up" title="Nach oben"><i class="bi bi-arrow-up"></i></button>` +
                `<button type="button" class="btn btn-ghost" data-act="down" title="Nach unten"><i class="bi bi-arrow-down"></i></button>` +
                `<button type="button" class="btn btn-ghost" data-act="del" title="Löschen"><i class="bi bi-trash"></i></button>` +
                `</span></li>`
            );
        })
        .join('');
}

function persistEditorMeta() {
    const tpl = selectedTemplate();
    if (!tpl) return null;
    const nameEl = $('ktplEditName');
    const schoolEl = $('ktplEditSchoolForm');
    const subjEl = $('ktplEditSubject');
    const stufeEl = $('ktplEditSchulstufe');
    const semEl = $('ktplEditSemester');
    const descEl = $('ktplEditDesc');
    const next = normalizeTemplate({
        ...tpl,
        name: nameEl ? nameEl.value : tpl.name,
        schoolForm: schoolEl ? schoolEl.value : tpl.schoolForm,
        subjectCode: subjEl ? subjEl.value : tpl.subjectCode,
        schulstufe: stufeEl ? stufeEl.value : tpl.schulstufe,
        semester: semEl ? semEl.value : tpl.semester,
        module: undefined,
        modul: undefined,
        description: descEl ? descEl.value : tpl.description,
        updatedAt: new Date().toISOString()
    });
    if (tpl.origin === 'central' && templateContentKey(next) === templateContentKey(tpl)) {
        return tpl;
    }
    upsertTemplate(next);
    refreshFromStorage();
    ui.selectedId = next.id;
    ensureExpandedForSelection();
    return next;
}

function persistChannelNamesFromDom() {
    let tpl = selectedTemplate();
    if (!tpl) return null;
    const list = $('ktplChannelBody');
    if (!list) return tpl;
    const rows = list.querySelectorAll('[data-ch]');
    let next = tpl;
    rows.forEach((row) => {
        const id = row.getAttribute('data-ch');
        const input = row.querySelector('.ktpl-ch-name');
        if (!id || !input) return;
        const val = String(input.value || '').trim();
        if (!val || isGeneralChannelName(val)) return;
        try {
            next = renameTemplateChannel(next, id, val);
        } catch {
            /* skip invalid */
        }
    });
    if (tpl.origin === 'central' && templateContentKey(next) === templateContentKey(tpl)) {
        return tpl;
    }
    upsertTemplate(next);
    refreshFromStorage();
    ui.selectedId = next.id;
    return next;
}

function renderApplyTemplateSelect() {
    const sel = $('ktplApplyTemplate');
    if (!sel) return;
    const filtered = filterTemplates(ui.templates, applyFilters());
    const cur = ui.applyTemplateId || sel.value;
    sel.innerHTML = '<option value="">— Vorlage wählen —</option>';
    for (const t of filtered) {
        const opt = document.createElement('option');
        opt.value = t.id;
        const tag = originLabel(t.origin);
        opt.textContent = [
            tag,
            t.schoolForm,
            t.subjectCode,
            t.schulstufe ? schulstufeLabel(t.schulstufe) : '',
            t.semester ? semesterLabel(t.semester) : '',
            t.name + ' (' + t.channels.length + ')'
        ]
            .filter(Boolean)
            .join(' · ');
        sel.appendChild(opt);
    }
    if ([...sel.options].some((o) => o.value === cur)) sel.value = cur;
    ui.applyTemplateId = sel.value;
}

function statusLabel(status) {
    const map = {
        create: 'Anlegen',
        ok: 'OK',
        rename: 'Umbenennen',
        skip_general: 'Allgemein',
        extra: 'Extra'
    };
    return map[status] || status;
}

function renderDiff(rows) {
    const tbody = $('ktplDiffBody');
    const summaryEl = $('ktplDiffSummary');
    if (!tbody) return;
    ui.lastDiff = rows;
    if (!rows || !rows.length) {
        tbody.innerHTML = '<tr><td colspan="3" class="muted">Kein Vergleich.</td></tr>';
        if (summaryEl) summaryEl.textContent = 'Noch kein Vergleich.';
        return;
    }
    const sum = summarizeDiff(rows);
    if (summaryEl) {
        summaryEl.textContent =
            `Anlegen: ${sum.create} · OK: ${sum.ok} · Umbenennen: ${sum.rename} · Extra: ${sum.extra}` +
            (sum.skip_general ? ` · Allgemein: ${sum.skip_general}` : '');
    }
    tbody.innerHTML = rows
        .map((r) => {
            const want = (r.templateChannel && r.templateChannel.displayName) || '—';
            const have = (r.teamChannel && r.teamChannel.displayName) || '—';
            return (
                `<tr class="ktpl-diff-${escapeHtml(r.status)}">` +
                `<td><span class="ktpl-status ktpl-status--${escapeHtml(r.status)}">${escapeHtml(statusLabel(r.status))}</span></td>` +
                `<td>${escapeHtml(want)}</td>` +
                `<td>${escapeHtml(have)}<div class="muted" style="font-size:0.85em;">${escapeHtml(r.message)}</div></td>` +
                `</tr>`
            );
        })
        .join('');
}

function downloadJson(filename, obj) {
    const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

function bindTabs() {
    document.querySelectorAll('[data-ktpl-tab]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const tab = btn.getAttribute('data-ktpl-tab');
            document.querySelectorAll('[data-ktpl-tab]').forEach((b) => {
                b.classList.toggle('is-active', b === btn);
                b.setAttribute('aria-selected', b === btn ? 'true' : 'false');
            });
            document.querySelectorAll('[data-ktpl-panel]').forEach((p) => {
                p.hidden = p.getAttribute('data-ktpl-panel') !== tab;
            });
        });
    });
}

function bindManage() {
    function refreshFiltersAndList() {
        ui.treeBootstrapped = false;
        fillMetaSelects();
        renderTemplateList();
        renderApplyTemplateSelect();
    }

    $('ktplFilterSchoolForm')?.addEventListener('change', (e) => {
        ui.schoolFormFilter = e.target.value || '';
        refreshFiltersAndList();
    });
    $('ktplFilterSubject')?.addEventListener('change', (e) => {
        ui.subjectFilter = e.target.value || '';
        refreshFiltersAndList();
    });
    $('ktplFilterSchulstufe')?.addEventListener('change', (e) => {
        ui.schulstufeFilter = e.target.value || '';
        refreshFiltersAndList();
    });
    $('ktplFilterSemester')?.addEventListener('change', (e) => {
        ui.semesterFilter = e.target.value || '';
        refreshFiltersAndList();
    });

    document.querySelectorAll('[data-ktpl-groupby]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const raw = btn.getAttribute('data-ktpl-groupby') || 'subject';
            const mode =
                raw === 'schulstufe' || raw === 'module'
                    ? 'schulstufe'
                    : raw === 'semester'
                      ? 'semester'
                      : raw === 'schoolForm'
                        ? 'schoolForm'
                        : 'subject';
            if (ui.groupBy === mode) return;
            ui.groupBy = mode;
            ui.expanded = new Set();
            ui.treeBootstrapped = false;
            document.querySelectorAll('[data-ktpl-groupby]').forEach((b) => {
                const on = b.getAttribute('data-ktpl-groupby') === mode;
                b.classList.toggle('is-active', on);
                b.setAttribute('aria-pressed', on ? 'true' : 'false');
            });
            renderTemplateList();
        });
    });

    $('ktplBtnSchoolFormAdd')?.addEventListener('click', async () => {
        const ask =
            typeof window.ms365Prompt === 'function'
                ? await window.ms365Prompt('Neue Schulform (z. B. HAK, AHS, HLW):', '')
                : window.prompt('Neue Schulform (z. B. HAK, AHS, HLW):', '');
        if (ask == null) return;
        const name = normalizeSchoolForm(ask);
        if (!name) return toast('Bitte eine Schulform eintragen.');
        ui.localSchoolForms = rememberSchoolForm(ui.localSchoolForms, name);
        saveState(ui.localTemplates, ui.localSchoolForms);
        refreshFromStorage();
        fillMetaSelects();
        const sel = $('ktplEditSchoolForm');
        if (sel) sel.value = name;
        toast('Schulform „' + name + '“ hinzugefügt.');
    });

    $('ktplBtnSchoolFormRename')?.addEventListener('click', async () => {
        const sel = $('ktplEditSchoolForm');
        const from = normalizeSchoolForm(sel ? sel.value : '');
        if (!from) return toast('Zuerst eine Schulform im Dropdown wählen.');
        const ask =
            typeof window.ms365Prompt === 'function'
                ? await window.ms365Prompt('Neuer Name für Schulform „' + from + '“:', from)
                : window.prompt('Neuer Name für Schulform „' + from + '“:', from);
        if (ask == null) return;
        const to = normalizeSchoolForm(ask);
        if (!to) return toast('Neuer Name fehlt.');
        if (to === from) return;
        let nextTemplates = renameSchoolFormInTemplates(ui.localTemplates, from, to);
        const selected = selectedTemplate();
        if (
            selected &&
            selected.origin === 'central' &&
            normalizeSchoolForm(selected.schoolForm) === from
        ) {
            const forked = normalizeTemplate({
                ...selected,
                schoolForm: to,
                updatedAt: new Date().toISOString()
            });
            const idx = nextTemplates.findIndex((t) => t.id === forked.id);
            if (idx >= 0) nextTemplates[idx] = forked;
            else nextTemplates.push(forked);
        }
        let catalog = ui.localSchoolForms.slice();
        catalog = catalog.filter((s) => normalizeSchoolForm(s) !== from);
        catalog = rememberSchoolForm(catalog, to);
        saveState(nextTemplates, catalog);
        refreshFromStorage();
        if (ui.schoolFormFilter === from) ui.schoolFormFilter = to;
        if (ui.applySchoolFormFilter === from) ui.applySchoolFormFilter = to;
        fillMetaSelects();
        if (sel) sel.value = to;
        renderTemplateList();
        renderEditor();
        renderApplyTemplateSelect();
        toast('Schulform „' + from + '“ → „' + to + '“ in der lokalen Bibliothek.');
    });

    $('ktplTemplateList')?.addEventListener('click', (e) => {
        const toggle = e.target.closest('[data-toggle]');
        if (toggle) {
            const key = toggle.getAttribute('data-toggle') || '';
            if (ui.expanded.has(key)) ui.expanded.delete(key);
            else ui.expanded.add(key);
            renderTemplateList();
            return;
        }
        const btn = e.target.closest('[data-id]');
        if (!btn) return;
        persistEditorMeta();
        persistChannelNamesFromDom();
        ui.selectedId = btn.getAttribute('data-id') || '';
        ensureExpandedForSelection();
        renderTemplateList();
        renderEditor();
    });

    $('ktplBtnNew')?.addEventListener('click', () => {
        const tpl = createEmptyTemplate(
            'Neue Vorlage',
            ui.subjectFilter || '',
            '',
            ui.schulstufeFilter || '',
            ui.schoolFormFilter || '',
            ui.semesterFilter || ''
        );
        upsertTemplate(tpl);
        refreshFromStorage();
        ui.selectedId = tpl.id;
        fillMetaSelects();
        renderTemplateList();
        renderEditor();
        renderApplyTemplateSelect();
        toast('Vorlage angelegt.');
    });

    $('ktplBtnClone')?.addEventListener('click', () => {
        const tpl = selectedTemplate();
        if (!tpl) return toast('Zuerst eine Vorlage wählen.');
        persistEditorMeta();
        persistChannelNamesFromDom();
        const copy = cloneTemplate(selectedTemplate());
        upsertTemplate(copy);
        refreshFromStorage();
        ui.selectedId = copy.id;
        renderTemplateList();
        renderEditor();
        renderApplyTemplateSelect();
        toast('Kopie angelegt.');
    });

    $('ktplBtnDelete')?.addEventListener('click', async () => {
        const tpl = selectedTemplate();
        if (!tpl) return;
        if (tpl.origin === 'central') {
            toast('Zentrale Vorlagen bleiben für alle Schulen. Im Admin pflegen, oder hier eine Kopie anlegen.');
            return;
        }
        const msg =
            tpl.origin === 'override'
                ? 'Lokale Anpassung von „' + tpl.name + '“ entfernen? Die zentrale Vorlage bleibt sichtbar.'
                : 'Vorlage „' + tpl.name + '“ wirklich löschen?';
        const ok =
            typeof window.ms365Confirm === 'function' ? await window.ms365Confirm(msg) : window.confirm(msg);
        if (!ok) return;
        deleteTemplate(tpl.id);
        refreshFromStorage();
        ui.selectedId = ui.templates[0] ? ui.templates[0].id : '';
        renderTemplateList();
        renderEditor();
        renderApplyTemplateSelect();
        renderCatalogStatus();
    });

    function doSave() {
        const before = selectedTemplate();
        const wasCentral = before && before.origin === 'central';
        persistEditorMeta();
        persistChannelNamesFromDom();
        const after = selectedTemplate();
        renderTemplateList();
        renderEditor();
        renderApplyTemplateSelect();
        if (wasCentral && after && after.origin === 'override') toast('Als lokale Anpassung gespeichert.');
        else if (wasCentral && after && after.origin === 'central') toast('Keine Änderung gegenüber der Zentrale.');
        else toast('Gespeichert.');
    }

    $('ktplBtnSave')?.addEventListener('click', doSave);

    $('ktplNewChannelName')?.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            $('ktplBtnAddChannel')?.click();
        }
    });

    $('ktplBtnAddChannel')?.addEventListener('click', () => {
        let tpl = persistEditorMeta();
        tpl = persistChannelNamesFromDom() || tpl;
        if (!tpl) return;
        const input = $('ktplNewChannelName');
        const name = input ? String(input.value || '').trim() : '';
        try {
            tpl = addTemplateChannel(tpl, name);
            upsertTemplate(tpl);
            refreshFromStorage();
            ui.selectedId = tpl.id;
            if (input) input.value = '';
            renderEditor();
            renderTemplateList();
            renderApplyTemplateSelect();
        } catch (err) {
            toast((err && err.message) || String(err));
        }
    });

    $('ktplChannelBody')?.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-act]');
        if (!btn) return;
        const row = btn.closest('[data-ch]');
        if (!row) return;
        const id = row.getAttribute('data-ch');
        const act = btn.getAttribute('data-act');
        let tpl = persistChannelNamesFromDom() || selectedTemplate();
        if (!tpl || !id) return;
        try {
            if (act === 'del') tpl = removeTemplateChannel(tpl, id);
            else if (act === 'up' || act === 'down') tpl = moveTemplateChannel(tpl, id, act);
            upsertTemplate(tpl);
            refreshFromStorage();
            ui.selectedId = tpl.id;
            renderEditor();
            renderTemplateList();
            renderApplyTemplateSelect();
        } catch (err) {
            toast((err && err.message) || String(err));
        }
    });

    $('ktplBtnExport')?.addEventListener('click', () => {
        persistEditorMeta();
        persistChannelNamesFromDom();
        refreshFromStorage();
        const payload = buildExportPayload(ui.templates);
        downloadJson('kursteam-vorlagen.json', payload);
        toast('Export gestartet.');
    });

    $('ktplBtnExportOne')?.addEventListener('click', () => {
        const tpl = persistEditorMeta();
        persistChannelNamesFromDom();
        if (!tpl) return toast('Keine Vorlage gewählt.');
        const payload = buildExportPayload([selectedTemplate()]);
        const safe = (tpl.name || 'vorlage').replace(/[^\w\-]+/g, '_').slice(0, 40);
        downloadJson('kursteam-vorlage-' + safe + '.json', payload);
    });

    $('ktplImportFile')?.addEventListener('change', async (e) => {
        const file = e.target.files && e.target.files[0];
        e.target.value = '';
        if (!file) return;
        try {
            const text = await file.text();
            const parsed = parseImportPayload(text);
            const merged = mergeTemplates(ui.templates, parsed.templates);
            saveTemplates(merged);
            // Schulformen aus Import merken
            let forms = ui.schoolForms.slice();
            for (const t of parsed.templates) forms = rememberSchoolForm(forms, t.schoolForm);
            saveState(merged, forms);
            refreshFromStorage();
            if (parsed.templates[0]) ui.selectedId = parsed.templates[0].id;
            fillMetaSelects();
            renderTemplateList();
            renderEditor();
            renderApplyTemplateSelect();
            toast(
                parsed.templates.length +
                    ' Vorlage(n) importiert' +
                    (parsed.warnings.length ? ' (' + parsed.warnings.join('; ') + ')' : '') +
                    '.'
            );
        } catch (err) {
            toast((err && err.message) || String(err));
        }
    });

    $('ktplBtnReset')?.addEventListener('click', async () => {
        const n = ui.localTemplates.length;
        const msg =
            n > 0
                ? 'Alle ' +
                  n +
                  ' lokalen Vorlagen löschen und die mitgelieferten Standard-Vorlagen (HAKB MAM) neu laden? Der zentrale Katalog bleibt unverändert.'
                : 'Mitgelieferte Standard-Vorlagen (HAKB MAM) lokal neu laden? Der zentrale Katalog bleibt unverändert.';
        const ok =
            typeof window.ms365Confirm === 'function'
                ? await window.ms365Confirm(msg)
                : window.confirm(msg);
        if (!ok) return;

        resetToSeedTemplates();
        refreshFromStorage();
        ui.selectedId = ui.templates[0] ? ui.templates[0].id : '';
        ui.expanded = new Set();
        ui.treeBootstrapped = false;
        ui.applyTemplateId = '';
        fillMetaSelects();
        renderTemplateList();
        renderEditor();
        renderApplyTemplateSelect();
        renderDiff([]);
        toast('Zurückgesetzt – ' + ui.templates.length + ' Standard-Vorlage(n) geladen.');
    });
}

function applyCatalogData(data) {
    ui.centralTemplates = Array.isArray(data.templates) ? data.templates : [];
    ui.centralSchoolForms = Array.isArray(data.schoolForms) ? data.schoolForms : [];
    ui.catalog = {
        ok: data.ok !== false,
        missing: !!data.missing,
        message: data.message || '',
        updatedAt: data.updatedAt || null,
        updatedBy: data.updatedBy || '',
        webUrl: data.webUrl || data.siteWebUrl || '',
        library: data.library || '',
        path: data.path || ''
    };
    const editing = document.activeElement && document.activeElement.closest('#ktplEditorForm');
    refreshFromStorage();
    fillMetaSelects();
    renderTemplateList();
    renderApplyTemplateSelect();
    renderCatalogStatus();
    if (!editing) renderEditor();
}

let catalogBusy = false;

async function reloadCentral() {
    if (catalogBusy) return;
    catalogBusy = true;
    try {
        const data = await fetchCentralCatalog();
        if (!data.ok) {
            ui.catalog.ok = false;
            ui.catalog.missing = true;
            ui.catalog.message = data.message || 'Zentrale nicht geladen.';
            renderCatalogStatus();
            return;
        }
        applyCatalogData(data);
    } catch (err) {
        const msg = (err && err.message) || String(err);
        ui.catalog.ok = false;
        ui.catalog.missing = true;
        ui.catalog.message = /404|not found|nicht gefunden/i.test(msg)
            ? 'Der zentrale Katalog ist auf dieser License-API noch nicht vorhanden. Bis zum Deploy gilt die lokale Bibliothek.'
            : msg;
        renderCatalogStatus();
    } finally {
        catalogBusy = false;
    }
}

function bindCatalog() {
    document.querySelectorAll('[data-ktpl-source]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const raw = btn.getAttribute('data-ktpl-source') || '';
            ui.sourceFilter = raw === 'central' || raw === 'local' ? raw : '';
            ui.treeBootstrapped = false;
            document.querySelectorAll('[data-ktpl-source]').forEach((b) => {
                const on = (b.getAttribute('data-ktpl-source') || '') === ui.sourceFilter;
                b.classList.toggle('is-active', on);
                b.setAttribute('aria-pressed', on ? 'true' : 'false');
            });
            renderTemplateList();
        });
    });

    $('ktplBtnCatalogRefresh')?.addEventListener('click', () => {
        reloadCentral();
    });
}

function bindApply() {
    $('ktplApplySchoolForm')?.addEventListener('change', (e) => {
        ui.applySchoolFormFilter = e.target.value || '';
        renderApplyTemplateSelect();
    });
    $('ktplApplySubjectFilter')?.addEventListener('change', (e) => {
        ui.applySubjectFilter = e.target.value || '';
        renderApplyTemplateSelect();
    });
    $('ktplApplySchulstufeFilter')?.addEventListener('change', (e) => {
        ui.applySchulstufeFilter = e.target.value || '';
        renderApplyTemplateSelect();
    });
    $('ktplApplySemesterFilter')?.addEventListener('change', (e) => {
        ui.applySemesterFilter = e.target.value || '';
        renderApplyTemplateSelect();
    });

    $('ktplApplyTemplate')?.addEventListener('change', (e) => {
        ui.applyTemplateId = e.target.value || '';
    });

    $('ktplBtnSearchTeam')?.addEventListener('click', async () => {
        const q = $('ktplTeamQuery')?.value || '';
        const ul = $('ktplTeamHits');
        try {
            const hits = await searchTeams(q);
            if (!ul) return;
            if (!hits.length) {
                ul.hidden = false;
                ul.innerHTML = '<li class="muted" style="padding:8px;">Keine Treffer.</li>';
                return;
            }
            ul.hidden = false;
            ul.innerHTML = hits
                .map(
                    (h) =>
                        `<li><button type="button" data-team-id="${escapeHtml(h.id)}" data-team-name="${escapeHtml(h.displayName)}" data-team-nick="${escapeHtml(h.mailNickname || '')}">` +
                        `<strong>${escapeHtml(h.displayName)}</strong>` +
                        `<span class="muted">${escapeHtml(h.mailNickname || h.mail || h.id)}</span></button></li>`
                )
                .join('');
        } catch (err) {
            toast((err && err.message) || String(err));
        }
    });

    $('ktplTeamHits')?.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-team-id]');
        if (!btn) return;
        ui.pickedTeam = {
            id: btn.getAttribute('data-team-id') || '',
            displayName: btn.getAttribute('data-team-name') || '',
            mailNickname: btn.getAttribute('data-team-nick') || ''
        };
        const picked = $('ktplTeamPicked');
        if (picked) {
            picked.innerHTML =
                '<strong>' +
                escapeHtml(ui.pickedTeam.displayName) +
                '</strong>' +
                (ui.pickedTeam.mailNickname
                    ? '<br><span class="muted">' + escapeHtml(ui.pickedTeam.mailNickname) + '</span>'
                    : '');
            picked.classList.remove('muted');
        }
        const ul = $('ktplTeamHits');
        if (ul) ul.hidden = true;
        ui.lastDiff = null;
        renderDiff([]);
    });

    $('ktplBtnDiff')?.addEventListener('click', async () => {
        if (!ui.pickedTeam) return toast('Zuerst ein Team wählen.');
        const tplId = $('ktplApplyTemplate')?.value || '';
        const tpl = ui.templates.find((t) => t.id === tplId);
        if (!tpl) return toast('Vorlage wählen.');
        try {
            log('Lade Kanäle von „' + ui.pickedTeam.displayName + '“…');
            const channels = await listChannels(ui.pickedTeam.id);
            log('Gefunden: ' + channels.length + ' Kanal/Kanäle.');
            const rows = diffChannels(tpl, channels);
            renderDiff(rows);
            toast('Vergleich fertig.');
        } catch (err) {
            toast((err && err.message) || String(err));
            log('Fehler: ' + ((err && err.message) || err));
        }
    });

    $('ktplBtnApply')?.addEventListener('click', async () => {
        if (!ui.pickedTeam) return toast('Zuerst ein Team wählen.');
        if (!ui.lastDiff || !ui.lastDiff.length) return toast('Zuerst vergleichen.');
        const doRename = !($('ktplSkipRename') && $('ktplSkipRename').checked);
        const need =
            ui.lastDiff.filter((r) => r.status === 'create').length +
            (doRename ? ui.lastDiff.filter((r) => r.status === 'rename').length : 0);
        if (!need) return toast('Nichts zu tun (alles OK oder nur Extra-Kanäle).');
        const ok =
            typeof window.ms365Confirm === 'function'
                ? await window.ms365Confirm(
                      need +
                          ' Änderung(en) auf „' +
                          ui.pickedTeam.displayName +
                          '“ anwenden?'
                  )
                : window.confirm(need + ' Änderung(en) anwenden?');
        if (!ok) return;
        try {
            log('Apply starten…');
            const results = await applyDiff(ui.lastDiff, ui.pickedTeam.id, {
                doRename,
                onProgress: (m) => log(m)
            });
            const fail = results.filter((r) => !r.ok);
            const okN = results.filter((r) => r.ok).length;
            log('Fertig: ' + okN + ' ok, ' + fail.length + ' Fehler.');
            fail.forEach((f) => log('  ! ' + f.name + ': ' + (f.error || '')));
            toast(fail.length ? 'Mit Fehlern beendet – siehe Protokoll.' : 'Vorlage angewendet.');
            // Diff neu laden
            const tplId = $('ktplApplyTemplate')?.value || '';
            const tpl = ui.templates.find((t) => t.id === tplId);
            if (tpl) {
                const channels = await listChannels(ui.pickedTeam.id);
                renderDiff(diffChannels(tpl, channels));
            }
        } catch (err) {
            toast((err && err.message) || String(err));
            log('Fehler: ' + ((err && err.message) || err));
        }
    });
}

function init() {
    refreshFromStorage();
    fillMetaSelects();
    if (!ui.selectedId && ui.templates[0]) ui.selectedId = ui.templates[0].id;
    document.querySelectorAll('[data-ktpl-groupby]').forEach((b) => {
        const on = b.getAttribute('data-ktpl-groupby') === ui.groupBy;
        b.classList.toggle('is-active', on);
        b.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
    bindTabs();
    bindManage();
    bindCatalog();
    bindApply();
    renderTemplateList();
    renderEditor();
    renderApplyTemplateSelect();
    renderCatalogStatus();
    renderDiff([]);
    let wasLoggedIn =
        typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
    window.addEventListener('ms365-auth-state-changed', (ev) => {
        renderCatalogStatus();
        const loggedIn = !!(ev.detail && ev.detail.loggedIn);
        if (loggedIn && !wasLoggedIn) reloadCentral();
        wasLoggedIn = loggedIn;
    });
    window.addEventListener('ms365-auth-widget-ready', () => renderCatalogStatus());
    reloadCentral();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
