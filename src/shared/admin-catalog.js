/**
 * Admin: zentrale Kursteam-Vorlagen pflegen und nach SharePoint schreiben.
 * Schulen lesen denselben Katalog im Werkzeug, ohne ihn zu veröffentlichen.
 */
import {
    createEmptyTemplate,
    cloneTemplate,
    addTemplateChannel,
    removeTemplateChannel,
    moveTemplateChannel,
    renameTemplateChannel,
    normalizeTemplate,
    normalizeTemplateList,
    collectSchoolForms,
    rememberSchoolForm,
    buildExportPayload,
    parseImportPayload,
    mergeTemplates,
    isGeneralChannelName,
    templateContentKey
} from '../tools/kursteam-templates/kursteam-templates-logic.js';
import { STORAGE_KEY } from '../tools/kursteam-templates/kursteam-templates-storage.js';
import { fetchCentralCatalog, publishCentralCatalog } from '../tools/kursteam-templates/kursteam-templates-catalog.js';

const DRAFT_KEY = 'ms365-kursteam-catalog-draft-v1';
const SITE = 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools';

function $(id) {
    return document.getElementById(id);
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

const ui = {
    templates: [],
    schoolForms: [],
    selectedId: '',
    dirty: false,
    query: '',
    remoteUpdatedAt: null,
    webUrl: ''
};

function readDraft() {
    try {
        const raw = JSON.parse(localStorage.getItem(DRAFT_KEY) || 'null');
        if (!raw || typeof raw !== 'object' || !Array.isArray(raw.templates)) return null;
        return {
            templates: normalizeTemplateList(raw.templates),
            schoolForms: Array.isArray(raw.schoolForms) ? raw.schoolForms : [],
            dirty: !!raw.dirty
        };
    } catch {
        return null;
    }
}

function writeDraft() {
    const payload = {
        templates: ui.templates,
        schoolForms: ui.schoolForms,
        dirty: ui.dirty
    };
    localStorage.setItem(DRAFT_KEY, JSON.stringify(payload));
}

function setBanner(text) {
    const el = $('adminCatBanner');
    if (!el) return;
    if (!text) {
        el.hidden = true;
        el.textContent = '';
        return;
    }
    el.hidden = false;
    el.textContent = text;
}

function stamp(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return '';
    return d.toLocaleString('de-AT', { dateStyle: 'short', timeStyle: 'short' });
}

function renderStatus() {
    const el = $('adminCatStatus');
    if (!el) return;
    const n = ui.templates.length;
    const when = stamp(ui.remoteUpdatedAt);
    let text = n + ' Vorlage' + (n === 1 ? '' : 'n');
    if (ui.dirty) text += ' · Entwurf';
    else if (when) text += ' · Stand ' + when;
    el.textContent = text;
    const link = $('adminCatLink');
    if (link) link.href = ui.webUrl || SITE;
}

function setupHint(message) {
    const msg = String(message || '');
    if (!/access denied|accessdenied|forbidden|schreibrecht/i.test(msg)) return msg;
    return (
        msg +
        '\n\nDie Bibliothek einmal selbst anlegen, danach erneut „In Zentrale schreiben“:\n' +
        '1. ' +
        SITE +
        ' öffnen\n' +
        '2. Zahnrad → Websiteinhalte → Neu → Dokumentbibliothek\n' +
        '3. Name genau: MS365-Katalog'
    );
}

function selected() {
    return ui.templates.find((t) => t.id === ui.selectedId) || null;
}

function formsNow() {
    let forms = ui.schoolForms.slice();
    for (const t of ui.templates) forms = rememberSchoolForm(forms, t.schoolForm);
    return forms;
}

function fillSchoolSelect(current) {
    const sel = $('adminCatSchool');
    if (!sel) return;
    const forms = collectSchoolForms(ui.templates, formsNow());
    const cur = current != null ? current : sel.value;
    sel.innerHTML = '<option value="">— keine Schulform —</option>';
    for (const name of forms) {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        sel.appendChild(opt);
    }
    if ([...sel.options].some((o) => o.value === cur)) sel.value = cur;
}

function renderList() {
    const ul = $('adminCatList');
    if (!ul) return;
    const q = ui.query.trim().toLowerCase();
    const rows = ui.templates.filter((t) => {
        if (!q) return true;
        return (t.name + ' ' + t.schoolForm + ' ' + t.subjectCode + ' ' + t.schulstufe)
            .toLowerCase()
            .includes(q);
    });
    if (!rows.length) {
        ul.innerHTML = '<li class="admin-catalog__empty">Keine Vorlagen.</li>';
        return;
    }
    ul.innerHTML = rows
        .map((t) => {
            const active = t.id === ui.selectedId ? ' is-active' : '';
            const meta = [t.schoolForm, t.subjectCode, t.schulstufe, t.semester].filter(Boolean).join(' · ');
            return (
                '<li><button type="button" class="' +
                active.trim() +
                '" data-id="' +
                escapeHtml(t.id) +
                '"><strong>' +
                escapeHtml(t.name) +
                '</strong><small>' +
                escapeHtml(meta || 'ohne Zuordnung') +
                ' · ' +
                t.channels.length +
                ' Kanäle</small></button></li>'
            );
        })
        .join('');
}

function renderChannels(tpl) {
    const ul = $('adminCatChannels');
    if (!ul) return;
    if (!tpl.channels.length) {
        ul.innerHTML = '<li class="admin-catalog__empty">Noch keine Kanäle.</li>';
        return;
    }
    ul.innerHTML = tpl.channels
        .map(
            (c, i) =>
                '<li class="admin-catalog__ch" data-ch="' +
                escapeHtml(c.id) +
                '"><span>' +
                (i + 1) +
                '</span><input type="text" value="' +
                escapeHtml(c.displayName) +
                '" aria-label="Kanalname"><span>' +
                '<button type="button" class="btn btn-sm alt" data-act="up" title="Nach oben"><i class="bi bi-arrow-up"></i></button> ' +
                '<button type="button" class="btn btn-sm alt" data-act="down" title="Nach unten"><i class="bi bi-arrow-down"></i></button> ' +
                '<button type="button" class="btn btn-sm alt" data-act="del" title="Löschen"><i class="bi bi-trash"></i></button>' +
                '</span></li>'
        )
        .join('');
}

function renderEditor() {
    const empty = $('adminCatEmpty');
    const form = $('adminCatForm');
    const tpl = selected();
    if (!tpl) {
        if (empty) empty.hidden = false;
        if (form) form.hidden = true;
        return;
    }
    if (empty) empty.hidden = true;
    if (form) form.hidden = false;
    const name = $('adminCatName');
    const subj = $('adminCatSubject');
    const stufe = $('adminCatStufe');
    const sem = $('adminCatSemester');
    const desc = $('adminCatDesc');
    if (name) name.value = tpl.name;
    fillSchoolSelect(tpl.schoolForm || '');
    if (subj) subj.value = tpl.subjectCode || '';
    if (stufe) stufe.value = tpl.schulstufe || '';
    if (sem) sem.value = tpl.semester || '';
    if (desc) desc.value = tpl.description || '';
    renderChannels(tpl);
}

function render() {
    renderStatus();
    renderList();
    renderEditor();
}

function markDirty() {
    ui.dirty = true;
    writeDraft();
    renderStatus();
}

function replaceWorking(templates, schoolForms, dirty) {
    ui.templates = normalizeTemplateList(templates);
    ui.schoolForms = Array.isArray(schoolForms) ? schoolForms.slice() : [];
    for (const t of ui.templates) ui.schoolForms = rememberSchoolForm(ui.schoolForms, t.schoolForm);
    ui.dirty = !!dirty;
    if (!ui.templates.some((t) => t.id === ui.selectedId)) {
        ui.selectedId = ui.templates[0] ? ui.templates[0].id : '';
    }
    writeDraft();
    render();
}

function readEditorInto(tpl) {
    if (!tpl) return tpl;
    return normalizeTemplate({
        ...tpl,
        name: $('adminCatName') ? $('adminCatName').value : tpl.name,
        schoolForm: $('adminCatSchool') ? $('adminCatSchool').value : tpl.schoolForm,
        subjectCode: $('adminCatSubject') ? $('adminCatSubject').value : tpl.subjectCode,
        schulstufe: $('adminCatStufe') ? $('adminCatStufe').value : tpl.schulstufe,
        semester: $('adminCatSemester') ? $('adminCatSemester').value : tpl.semester,
        description: $('adminCatDesc') ? $('adminCatDesc').value : tpl.description,
        updatedAt: tpl.updatedAt
    });
}

function commitEditor() {
    const tpl = selected();
    if (!tpl) return null;
    const next = readEditorInto(tpl);
    const list = $('adminCatChannels');
    let withNames = next;
    if (list) {
        list.querySelectorAll('[data-ch]').forEach((row) => {
            const id = row.getAttribute('data-ch');
            const input = row.querySelector('input');
            if (!id || !input) return;
            const val = String(input.value || '').trim();
            if (!val || isGeneralChannelName(val)) return;
            try {
                withNames = renameTemplateChannel(withNames, id, val);
            } catch {
                /* ungültigen Namen überspringen */
            }
        });
    }
    const idx = ui.templates.findIndex((t) => t.id === withNames.id);
    if (idx >= 0) ui.templates[idx] = withNames;
    ui.schoolForms = rememberSchoolForm(ui.schoolForms, withNames.schoolForm);
    ui.selectedId = withNames.id;
    if (templateContentKey(withNames) !== templateContentKey(tpl)) markDirty();
    return withNames;
}

function readToolLibrary() {
    try {
        const raw = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
        if (!raw || typeof raw !== 'object' || !Array.isArray(raw.templates)) {
            return { templates: [], schoolForms: [] };
        }
        return {
            templates: normalizeTemplateList(raw.templates),
            schoolForms: Array.isArray(raw.schoolForms) ? raw.schoolForms : []
        };
    } catch {
        return { templates: [], schoolForms: [] };
    }
}

function applyRemote(data, discardDraft) {
    ui.remoteUpdatedAt = data.updatedAt || null;
    ui.webUrl = data.webUrl || data.siteWebUrl || '';
    const draft = readDraft();
    if (!discardDraft && draft && draft.dirty && draft.templates.length) {
        ui.templates = draft.templates;
        ui.schoolForms = draft.schoolForms;
        ui.dirty = true;
        if (!ui.templates.some((t) => t.id === ui.selectedId)) {
            ui.selectedId = ui.templates[0] ? ui.templates[0].id : '';
        }
        setBanner('In diesem Browser liegt ein Entwurf, der noch nicht in der Zentrale ist.');
        render();
        return;
    }
    setBanner(data.missing ? 'Die Zentrale ist noch leer. Vorlagen anlegen und dann „In Zentrale schreiben“.' : '');
    replaceWorking(data.templates || [], data.schoolForms || [], false);
}

async function reloadFromCentral(mode) {
    const discard = mode === 'ask';
    if (discard && ui.dirty) {
        const ok = window.confirm('Entwurf verwerfen und den Stand aus der Zentrale laden?');
        if (!ok) return;
    }
    const status = $('adminCatStatus');
    if (status) status.textContent = 'Lade …';
    try {
        const data = await fetchCentralCatalog();
        if (!data.ok) {
            setBanner(setupHint(data.message || 'Zentrale nicht geladen.'));
            const draft = readDraft();
            if (draft) {
                ui.templates = draft.templates;
                ui.schoolForms = draft.schoolForms;
                ui.dirty = draft.dirty;
                render();
            } else {
                renderStatus();
            }
            return;
        }
        applyRemote(data, discard);
    } catch (err) {
        setBanner(setupHint((err && err.message) || String(err)));
        renderStatus();
    }
}

function bind() {
    $('adminCatSearch')?.addEventListener('input', (e) => {
        ui.query = e.target.value || '';
        renderList();
    });
    $('adminCatList')?.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-id]');
        if (!btn) return;
        commitEditor();
        ui.selectedId = btn.getAttribute('data-id') || '';
        render();
    });
    $('adminCatNew')?.addEventListener('click', () => {
        commitEditor();
        const tpl = createEmptyTemplate('Neue Vorlage');
        ui.templates.push(tpl);
        ui.selectedId = tpl.id;
        markDirty();
        render();
    });
    $('adminCatClone')?.addEventListener('click', () => {
        const current = commitEditor();
        if (!current) return;
        const copy = cloneTemplate(current);
        ui.templates.push(copy);
        ui.selectedId = copy.id;
        markDirty();
        render();
    });
    $('adminCatDelete')?.addEventListener('click', () => {
        const tpl = selected();
        if (!tpl) return;
        if (!window.confirm('„' + tpl.name + '“ aus dem Entwurf löschen? In der Zentrale ist sie erst weg, wenn du danach schreibst.')) {
            return;
        }
        ui.templates = ui.templates.filter((t) => t.id !== tpl.id);
        ui.selectedId = ui.templates[0] ? ui.templates[0].id : '';
        markDirty();
        render();
    });

    function bindField(id) {
        $(id)?.addEventListener('input', () => {
            const tpl = selected();
            if (!tpl) return;
            const next = readEditorInto(tpl);
            const idx = ui.templates.findIndex((t) => t.id === next.id);
            if (idx >= 0) ui.templates[idx] = next;
            markDirty();
            renderList();
        });
    }
    ['adminCatName', 'adminCatSubject', 'adminCatStufe', 'adminCatDesc'].forEach(bindField);
    $('adminCatSchool')?.addEventListener('change', () => {
        const tpl = selected();
        if (!tpl) return;
        const next = readEditorInto(tpl);
        const idx = ui.templates.findIndex((t) => t.id === next.id);
        if (idx >= 0) ui.templates[idx] = next;
        ui.schoolForms = rememberSchoolForm(ui.schoolForms, next.schoolForm);
        markDirty();
        renderList();
    });
    $('adminCatSemester')?.addEventListener('change', () => {
        const tpl = selected();
        if (!tpl) return;
        const next = readEditorInto(tpl);
        const idx = ui.templates.findIndex((t) => t.id === next.id);
        if (idx >= 0) ui.templates[idx] = next;
        markDirty();
        renderList();
    });

    $('adminCatChannelAdd')?.addEventListener('click', () => {
        let tpl = commitEditor();
        if (!tpl) return;
        const input = $('adminCatChannelName');
        const name = input ? String(input.value || '').trim() : '';
        try {
            tpl = addTemplateChannel(tpl, name);
            const idx = ui.templates.findIndex((t) => t.id === tpl.id);
            if (idx >= 0) ui.templates[idx] = tpl;
            if (input) input.value = '';
            markDirty();
            renderChannels(tpl);
            renderList();
        } catch (err) {
            setBanner((err && err.message) || String(err));
        }
    });
    $('adminCatChannelName')?.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            $('adminCatChannelAdd')?.click();
        }
    });
    $('adminCatChannels')?.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-act]');
        if (!btn) return;
        const row = btn.closest('[data-ch]');
        if (!row) return;
        let tpl = commitEditor();
        if (!tpl) return;
        const id = row.getAttribute('data-ch');
        const act = btn.getAttribute('data-act');
        try {
            if (act === 'del') tpl = removeTemplateChannel(tpl, id);
            else tpl = moveTemplateChannel(tpl, id, act === 'up' ? 'up' : 'down');
            const idx = ui.templates.findIndex((t) => t.id === tpl.id);
            if (idx >= 0) ui.templates[idx] = tpl;
            markDirty();
            renderChannels(tpl);
            renderList();
        } catch (err) {
            setBanner((err && err.message) || String(err));
        }
    });
    $('adminCatChannels')?.addEventListener('input', (e) => {
        if (!e.target.matches('input')) return;
        commitEditor();
    });

    $('adminCatReload')?.addEventListener('click', () => reloadFromCentral('ask'));
    $('adminCatTakeLocal')?.addEventListener('click', () => {
        commitEditor();
        const local = readToolLibrary();
        if (!local.templates.length) {
            setBanner('In diesem Browser liegen keine Vorlagen aus dem Werkzeug.');
            return;
        }
        if (
            ui.templates.length &&
            !window.confirm(
                local.templates.length + ' Vorlagen aus dem Werkzeug in diesen Entwurf übernehmen? Gleiche IDs werden ersetzt.'
            )
        ) {
            return;
        }
        const merged = mergeTemplates(ui.templates, local.templates);
        replaceWorking(merged, local.schoolForms, true);
        setBanner('Übernommen. Noch nicht in der Zentrale – dafür „In Zentrale schreiben“.');
    });
    $('adminCatImport')?.addEventListener('change', async (e) => {
        const file = e.target.files && e.target.files[0];
        e.target.value = '';
        if (!file) return;
        try {
            commitEditor();
            const parsed = parseImportPayload(await file.text());
            const merged = mergeTemplates(ui.templates, parsed.templates);
            let forms = ui.schoolForms.slice();
            for (const t of parsed.templates) forms = rememberSchoolForm(forms, t.schoolForm);
            replaceWorking(merged, forms, true);
            setBanner(parsed.templates.length + ' Vorlage(n) in den Entwurf importiert.');
        } catch (err) {
            setBanner((err && err.message) || String(err));
        }
    });
    $('adminCatPublish')?.addEventListener('click', async () => {
        commitEditor();
        if (!ui.templates.length) {
            setBanner('Der Entwurf ist leer.');
            return;
        }
        if (
            !window.confirm(
                ui.templates.length +
                    ' Vorlagen in die zentrale Bibliothek schreiben? Alle Schulen sehen danach diesen Stand.'
            )
        ) {
            return;
        }
        const status = $('adminCatStatus');
        if (status) status.textContent = 'Schreibe …';
        try {
            const payload = buildExportPayload(ui.templates);
            payload.schoolForms = formsNow();
            const data = await publishCentralCatalog(payload);
            ui.remoteUpdatedAt = data.updatedAt || null;
            ui.webUrl = data.webUrl || data.siteWebUrl || '';
            replaceWorking(data.templates || ui.templates, data.schoolForms || [], false);
            setBanner('Zentrale aktualisiert. Schulen sehen diesen Stand nach „Zentrale“ im Werkzeug.');
        } catch (err) {
            setBanner(setupHint((err && err.message) || String(err)));
            renderStatus();
        }
    });
}

function init() {
    if (!$('adminPanelTemplates')) return;
    bind();
    const draft = readDraft();
    if (draft && draft.templates.length) {
        ui.templates = draft.templates;
        ui.schoolForms = draft.schoolForms;
        ui.dirty = draft.dirty;
        ui.selectedId = ui.templates[0] ? ui.templates[0].id : '';
        render();
    }
    if (typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn()) {
        reloadFromCentral('keep-draft');
    } else {
        setBanner('Mit dem Betreiber-Konto anmelden, dann lädt die Zentrale.');
        window.addEventListener('ms365-auth-state-changed', function onAuth(ev) {
            if (ev.detail && ev.detail.loggedIn) {
                window.removeEventListener('ms365-auth-state-changed', onAuth);
                reloadFromCentral('keep-draft');
            }
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
