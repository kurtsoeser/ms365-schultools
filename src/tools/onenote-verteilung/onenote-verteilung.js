import { escapeHtml } from '../../shared/utils/strings.js';
/**
 * OneNote-Inhalte verteilen – 4-Schritt-Assistent
 * Material → Teams → Wohin → Kopieren
 */
import {
    searchTeams,
    searchSites,
    listOnenoteNotebooks,
    listCentralTemplateNotebooks,
    loadOnenoteNotebookTree,
    copySectionToGroupSectionGroup,
    listSectionPages,
    getPagePreview,
    getPageContent,
    pickClassNotebook,
    pickSectionGroup,
    publishCentralNotebookSnapshot,
    listPublishedSnapshotNotebooks,
    CENTRAL_TEMPLATE_NOTEBOOK_NAME
} from './onenote-verteilung-graph.js';
import {
    initOnenoteSteps,
    bindWizardNav,
    bindStep1,
    bindStep2,
    bindStep3,
    bindStep4
} from './onenote-verteilung-steps.js';

function $(id) {
    return document.getElementById(id);
}

function toast(m) {
    const el = $('toast');
    if (el) {
        el.textContent = String(m || '');
        el.classList.add('show');
        clearTimeout(toast._t);
        toast._t = setTimeout(() => el.classList.remove('show'), 4200);
        return;
    }
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else if (typeof window.ms365ShowToast === 'function') window.ms365ShowToast(m);
    else window.alert(m);
}


function log(msg) {
    const el = $('onvLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(msg || '');
    el.scrollTop = el.scrollHeight;
}

const ONENOTE_TAB_COLORS = [
    '#5c2d91',
    '#c239b3',
    '#e3008c',
    '#d13438',
    '#ca5010',
    '#8e562e',
    '#986f0b',
    '#498205',
    '#00ad56',
    '#00b7c3',
    '#0078d4',
    '#4f6bed',
    '#8764b8',
    '#69797e'
];

import { ui } from './onenote-verteilung-state.js';

function sourceScope() {
    if (ui.srcMode === 'central') {
        if (ui.srcVia === 'catalog-api' || ui.srcVia === 'catalog-snapshot') {
            return { kind: 'catalog' };
        }
        if (ui.srcSite && ui.srcSite.id) return { kind: 'site', id: ui.srcSite.id };
    }
    if (ui.srcMode === 'team' && ui.srcGroup && ui.srcGroup.id) {
        return { kind: 'group', id: ui.srcGroup.id };
    }
    if (ui.srcMode === 'site' && ui.srcSite && ui.srcSite.id) {
        return { kind: 'site', id: ui.srcSite.id };
    }
    return { kind: 'me' };
}

function srcModeLabel(mode) {
    if (mode === 'me') return 'Dein OneDrive';
    if (mode === 'team') return 'Team / Kursnotizbuch – zuerst Team suchen und wählen';
    if (mode === 'site') return 'SharePoint-Site – suchen oder Site-URL einfügen';
    return 'Zentrale Vorlagen (MS365-Katalog / notebooks)';
}

function updateSrcPickVisibility() {
    const pickTeam = $('onvSrcPickTeam');
    const pickSite = $('onvSrcPickSite');
    if (pickTeam) pickTeam.hidden = ui.srcMode !== 'team';
    if (pickSite) pickSite.hidden = ui.srcMode !== 'site';
    const btn = $('onvBtnLoadSrc');
    if (btn) {
        const icon = '<i class="bi bi-cloud-arrow-down"></i>';
        if (ui.srcMode === 'central') btn.innerHTML = icon + 'Vorlagen laden';
        else if (ui.srcMode === 'me') btn.innerHTML = icon + 'OneDrive laden';
        else if (ui.srcMode === 'team') btn.innerHTML = icon + 'Team-Notizbücher laden';
        else btn.innerHTML = icon + 'Site-Notizbücher laden';
    }
    updatePublishBtn();
    renderSrcTeamSelected();
    renderSrcSiteSelected();
}

function canPublishSnapshots() {
    // Veröffentlichen braucht Site-Zugriff (kurtrocks), Notebooks dürfen aus dem Snapshot kommen
    return (
        ui.srcMode === 'central' &&
        !!(ui.srcSite && ui.srcSite.id) &&
        Array.isArray(ui.onSrcNotebooks) &&
        ui.onSrcNotebooks.length > 0
    );
}

function updatePublishBtn() {
    const btn = $('onvBtnPublishSnapshot');
    const hint = $('onvPublishHint');
    if (!btn) return;
    const show = canPublishSnapshots();
    btn.hidden = !show;
    if (hint) hint.hidden = !show;
    if (!show) {
        btn.textContent = 'Veröffentlichen';
        return;
    }
    const n = ui.publishPickIds.size;
    const label = n > 0 ? 'Veröffentlichen (' + n + ')' : 'Veröffentlichen';
    btn.innerHTML = '<i class="bi bi-cloud-upload"></i>' + label;
    if (hint) {
        const pub = ui.publishedNotebookIds.size;
        hint.textContent =
            pub > 0
                ? pub +
                  ' Buch/Bücher bereits für Schulen veröffentlicht · Haken setzen, dann Veröffentlichen'
                : 'Haken setzen bei den Büchern für Schulen, dann Veröffentlichen';
    }
}

async function refreshPublishedMarks() {
    ui.publishedNotebookIds = new Set();
    ui.publishedAtById = new Map();
    if (!canPublishSnapshots()) {
        updatePublishBtn();
        return;
    }
    try {
        const published = await listPublishedSnapshotNotebooks();
        ui.publishedNotebookIds = new Set(published.map((p) => p.id));
        published.forEach((p) => {
            if (p.id && p.publishedAt) ui.publishedAtById.set(p.id, p.publishedAt);
        });
    } catch {
        ui.publishedNotebookIds = new Set();
        ui.publishedAtById = new Map();
    }
    const sel = $('onvSrcNotebook');
    renderNotebookShelf(ui.onSrcNotebooks, sel && sel.value);
    updatePublishBtn();
}

function renderSrcTeamSelected() {
    const el = $('onvSrcTeamSelected');
    if (!el) return;
    if (ui.srcGroup && ui.srcGroup.id) {
        el.innerHTML =
            'Gewählt: <strong>' +
            escapeHtml(ui.srcGroup.displayName || ui.srcGroup.id) +
            '</strong>' +
            (ui.srcGroup.mailNickname
                ? ' <span class="muted">(' + escapeHtml(ui.srcGroup.mailNickname) + ')</span>'
                : '');
    } else {
        el.textContent = 'Noch kein Team gewählt.';
    }
}

function renderSrcSiteSelected() {
    const el = $('onvSrcSiteSelected');
    if (!el) return;
    if (ui.srcSite && ui.srcSite.id && ui.srcMode === 'site') {
        el.innerHTML =
            'Gewählt: <strong>' +
            escapeHtml(ui.srcSite.displayName || ui.srcSite.id) +
            '</strong>' +
            (ui.srcSite.webUrl
                ? '<br><span class="muted" style="font-size:0.78rem;word-break:break-all;">' +
                  escapeHtml(ui.srcSite.webUrl) +
                  '</span>'
                : '');
    } else if (ui.srcMode === 'site') {
        el.textContent = 'Noch keine Site gewählt.';
    }
}

function onenoteTabColor(seed) {
    const s = String(seed || '');
    let h = 0;
    for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
    return ONENOTE_TAB_COLORS[h % ONENOTE_TAB_COLORS.length];
}

function findSectionName(tree, sectionId) {
    if (!tree) return '';
    for (const s of tree.sections || []) {
        if (s.id === sectionId) return s.displayName;
    }
    for (const g of tree.groups || []) {
        for (const s of g.sections || []) {
            if (s.id === sectionId) return s.displayName;
        }
    }
    return '';
}

function countTreeSections(tree) {
    if (!tree) return { sections: 0, groups: 0 };
    let n = (tree.sections || []).length;
    for (const g of tree.groups || []) n += (g.sections || []).length;
    return { sections: n, groups: (tree.groups || []).length };
}

const NOTEBOOK_COVER_COLORS = [
    { spine: '#5c2d91', face: '#8764b8' },
    { spine: '#0078d4', face: '#2b88d8' },
    { spine: '#00ad56', face: '#26c281' },
    { spine: '#ca5010', face: '#e87a2e' },
    { spine: '#d13438', face: '#e3575a' },
    { spine: '#038387', face: '#00b7c3' },
    { spine: '#8764b8', face: '#a18bce' },
    { spine: '#4f6bed', face: '#6b83f2' },
    { spine: '#986f0b', face: '#c19c1c' },
    { spine: '#69797e', face: '#859399' }
];

function notebookCoverColors(seed) {
    const s = String(seed || '');
    let h = 0;
    for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
    return NOTEBOOK_COVER_COLORS[h % NOTEBOOK_COVER_COLORS.length];
}

function fillOnSelect(sel, items, emptyLabel, preferredId) {
    if (!sel) return;
    const cur = preferredId != null ? preferredId : sel.value;
    sel.innerHTML = '';
    const first = document.createElement('option');
    first.value = '';
    first.textContent = emptyLabel;
    sel.appendChild(first);
    for (const item of items || []) {
        const opt = document.createElement('option');
        opt.value = item.id;
        opt.textContent = item.displayName || item.id;
        sel.appendChild(opt);
    }
    if (cur && [...sel.options].some((o) => o.value === cur)) sel.value = cur;
    renderNotebookShelf(items || [], sel.value);
}

function formatNotebookWhen(iso) {
    const raw = String(iso || '').trim();
    if (!raw) return '';
    const d = new Date(raw);
    if (Number.isNaN(d.getTime())) return '';
    try {
        return new Intl.DateTimeFormat('de-AT', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        }).format(d);
    } catch {
        return d.toLocaleString('de-AT');
    }
}

function renderNotebookShelf(items, selectedId) {
    const shelf = $('onvNotebookShelf');
    if (!shelf) return;
    const list = Array.isArray(items) ? items : [];
    if (!list.length) {
        shelf.innerHTML =
            '<p class="onv-nb-shelf__empty" id="onvNotebookEmpty">Noch keine Notizbücher – „Vorlagen laden“ tippen.</p>';
        return;
    }
    const prefer = CENTRAL_TEMPLATE_NOTEBOOK_NAME.toLowerCase();
    const showPublishUi = canPublishSnapshots();
    shelf.innerHTML = list
        .map((nb) => {
            const colors = notebookCoverColors(nb.id || nb.displayName);
            const name = nb.displayName || nb.id;
            const isPref = name.toLowerCase() === prefer;
            const selected = selectedId && selectedId === nb.id;
            const published = ui.publishedNotebookIds.has(nb.id);
            const picked = ui.publishPickIds.has(nb.id);
            const when = formatNotebookWhen(nb.lastModifiedDateTime);
            const who = String(nb.lastModifiedByName || '').trim();
            const pubWhen = formatNotebookWhen(ui.publishedAtById.get(nb.id));
            const metaParts = [];
            if (when) {
                metaParts.push(
                    '<span class="onv-nb-meta__when" title="Zuletzt in OneNote bearbeitet">' +
                        '<span class="onv-nb-meta__lbl">Bearbeitet</span> ' +
                        escapeHtml(when) +
                        '</span>'
                );
            }
            if (who) {
                metaParts.push(
                    '<span class="onv-nb-meta__who" title="' +
                        escapeHtml(who) +
                        '">' +
                        escapeHtml(who) +
                        '</span>'
                );
            } else if (when) {
                metaParts.push('<span class="onv-nb-meta__who">Bearbeiter unbekannt</span>');
            }
            if (published) {
                metaParts.push(
                    '<span class="onv-nb-meta__pub" title="Zuletzt als Schul-Snapshot veröffentlicht">' +
                        '<span class="onv-nb-meta__lbl">Veröffentlicht</span>' +
                        (pubWhen ? ' ' + escapeHtml(pubWhen) : '') +
                        '</span>'
                );
            } else if (showPublishUi) {
                metaParts.push(
                    '<span class="onv-nb-meta__pub onv-nb-meta__pub--missing">Nicht veröffentlicht</span>'
                );
            }
            const metaHtml = metaParts.length
                ? metaParts.join('')
                : '<span class="onv-nb-meta__empty">Keine Änderungsinfo</span>';
            const tipParts = [name];
            if (when) tipParts.push('Bearbeitet: ' + when);
            if (who) tipParts.push('Von: ' + who);
            if (published) {
                tipParts.push(
                    pubWhen ? 'Veröffentlicht: ' + pubWhen : 'Veröffentlicht'
                );
            }
            const pickHtml = showPublishUi
                ? '<label class="onv-nb-pick" title="Für Snapshot auswählen">' +
                  '<input type="checkbox" data-nb-publish-pick="' +
                  escapeHtml(nb.id) +
                  '"' +
                  (picked ? ' checked' : '') +
                  '>' +
                  '<span>Snapshot</span></label>'
                : '';
            return (
                '<div class="onv-nb-item' +
                (selected ? ' is-selected' : '') +
                (published ? ' is-published' : '') +
                '">' +
                pickHtml +
                '<button type="button" class="onv-nb-cover' +
                (selected ? ' is-selected' : '') +
                (isPref ? ' is-preferred' : '') +
                (published ? ' is-published' : '') +
                '" role="option" aria-selected="' +
                (selected ? 'true' : 'false') +
                '" data-nb-id="' +
                escapeHtml(nb.id) +
                '" style="--nb-spine:' +
                colors.spine +
                ';--nb-face:' +
                colors.face +
                '" title="' +
                escapeHtml(tipParts.join(' · ')) +
                '">' +
                '<span class="onv-nb-cover__spine" aria-hidden="true"></span>' +
                '<span class="onv-nb-cover__face">' +
                '<span class="onv-nb-cover__badge" aria-hidden="true">N</span>' +
                '<span class="onv-nb-cover__title">' +
                escapeHtml(name) +
                '</span>' +
                (published
                    ? '<span class="onv-nb-cover__hint">Veröffentlicht</span>'
                    : isPref
                      ? '<span class="onv-nb-cover__hint">Vorlage</span>'
                      : '<span class="onv-nb-cover__hint">Notizbuch</span>') +
                '</span></button>' +
                '<div class="onv-nb-meta">' +
                metaHtml +
                '</div></div>'
            );
        })
        .join('');
}

function selectNotebook(id, opts) {
    const sel = $('onvSrcNotebook');
    const nid = String(id || '').trim();
    const prev = sel ? String(sel.value || '') : '';
    if (sel) {
        if (nid && [...sel.options].some((o) => o.value === nid)) sel.value = nid;
        else if (!nid) sel.value = '';
    }
    renderNotebookShelf(ui.onSrcNotebooks, nid);
    const skipSame =
        opts && opts.force
            ? false
            : nid &&
              nid === prev &&
              ui.onSrcTree &&
              ui.loadedSrcNotebookId === nid;
    if (skipSame) return;
    if (!opts || opts.dispatch !== false) {
        if (sel) sel.dispatchEvent(new Event('change', { bubbles: true }));
    }
}

const PREVIEW_IFRAME_CSS = `
html,body{margin:0;padding:0;background:#fff;color:#252423;
font:14px/1.45 "Segoe UI",Calibri,Arial,sans-serif;}
body{padding:12px 14px 18px;}
img,object,video{max-width:100%;height:auto;}
table{border-collapse:collapse;width:100%;margin:10px 0;font-size:13px;}
th,td{border:1px solid #c8c6c4;padding:6px 8px;vertical-align:top;}
th{background:#f3f2f1;font-weight:700;}
p{margin:0 0 8px;}
h1,h2,h3,h4{margin:12px 0 6px;line-height:1.25;}
ul,ol{margin:0 0 10px;padding-left:1.4em;}
a{color:#5c2d91;}
`;

function buildPreviewDocument(html) {
    const raw = String(html || '');
    const hasHtml = /<html[\s>]/i.test(raw);
    if (hasHtml) {
        // Inject table-friendly CSS into existing document
        if (/<\/head>/i.test(raw)) {
            return raw.replace(/<\/head>/i, '<style>' + PREVIEW_IFRAME_CSS + '</style></head>');
        }
        return raw;
    }
    return (
        '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
        PREVIEW_IFRAME_CSS +
        '</style></head><body>' +
        raw +
        '</body></html>'
    );
}

function teamsWithNotebook() {
    return ui.selectedTeams.filter((t) => t.notebookId && !t.warn);
}

function canEnterStep(n) {
    if (n <= 1) return true;
    if (n === 2) return ui.onSrcChecked.size >= 1;
    if (n === 3) return ui.onSrcChecked.size >= 1 && teamsWithNotebook().length >= 1;
    if (n === 4) return ui.onSrcChecked.size >= 1 && ui.onTargets.length >= 1;
    return false;
}

function updateNavButtons() {
    const next1 = $('onvBtnNext1');
    const next2 = $('onvBtnNext2');
    const next3 = $('onvBtnNext3');
    const copy = $('onvBtnCopy');
    if (next1) next1.disabled = ui.onSrcChecked.size < 1;
    if (next2) next2.disabled = teamsWithNotebook().length < 1;
    if (next3) next3.disabled = ui.onTargets.length < 1;
    if (copy) {
        copy.disabled = !(ui.onSrcChecked.size && ui.onTargets.length);
    }

    for (let i = 1; i <= 4; i++) {
        const ind = $('onvStepInd' + i);
        if (!ind) continue;
        ind.classList.toggle('is-active', ui.step === i);
        ind.classList.toggle('is-done', i < ui.step);
        ind.disabled = i > ui.step && !canEnterStep(i);
        if (i <= ui.step) ind.disabled = false;
    }
}

function showStep(n) {
    const target = Math.max(1, Math.min(4, Number(n) || 1));
    if (target > ui.step && !canEnterStep(target)) {
        toast('Bitte zuerst die vorherigen Schritte ausfüllen.');
        return;
    }
    ui.step = target;
    for (let i = 1; i <= 4; i++) {
        const panel = $('onvStep' + i);
        if (!panel) continue;
        const on = i === ui.step;
        panel.classList.toggle('is-visible', on);
        panel.hidden = !on;
    }
    updateNavButtons();
    if (ui.step === 3 && !ui.onTargets.length && teamsWithNotebook().length) {
        buildTargets().catch(() => {});
    }
    if (ui.step === 4) renderSummary();
}

function setSrcMode(mode) {
    const m = String(mode || 'central');
    ui.srcMode =
        m === 'me' ? 'me' : m === 'team' ? 'team' : m === 'site' ? 'site' : 'central';
    document.querySelectorAll('[data-onv-src-mode]').forEach((btn) => {
        const on = btn.getAttribute('data-onv-src-mode') === ui.srcMode;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
    const label = $('onvSrcModeLabel');
    if (label) label.textContent = srcModeLabel(ui.srcMode);
    ui.onSrcNotebooks = [];
    ui.onSrcTree = null;
    ui.onSrcChecked = new Set();
    ui.onSrcExpanded = new Set();
    ui.srcVia = '';
    ui.loadedSrcNotebookId = '';
    ui.srcTreeLoadGen++;
    ui.publishedNotebookIds = new Set();
    ui.publishedAtById = new Map();
    ui.publishPickIds = new Set();
    ui.previewSectionId = '';
    ui.previewPageId = '';
    if (ui.srcMode !== 'site') ui.srcSite = null;
    if (ui.srcMode !== 'team') ui.srcGroup = null;
    if (ui.srcMode === 'central') {
        ui.srcSite = null;
        ui.srcGroup = null;
    }
    fillOnSelect($('onvSrcNotebook'), [], '— laden —');
    const title = $('onvSrcTitle');
    if (title) title.textContent = 'Quell-Notizbuch';
    const hitsT = $('onvSrcTeamHits');
    if (hitsT) {
        hitsT.innerHTML = '';
        hitsT.hidden = true;
    }
    const hitsS = $('onvSrcSiteHits');
    if (hitsS) {
        hitsS.innerHTML = '';
        hitsS.hidden = true;
    }
    clearPreview();
    renderSrcTree();
    updateSrcPickVisibility();
    updateNavButtons();
}

function clearPreview() {
    const pages = $('onvPreviewPages');
    const body = $('onvPreviewBody');
    const title = $('onvPreviewTitle');
    const meta = $('onvPreviewMeta');
    const pagesTitle = $('onvPagesTitle');
    const pagesMeta = $('onvPagesMeta');
    ui.previewPages = [];
    if (pages) pages.innerHTML = '';
    if (pagesTitle) pagesTitle.textContent = 'Seiten';
    if (pagesMeta) pagesMeta.textContent = 'Abschnitt wählen';
    if (title) title.textContent = 'Vorschau';
    if (meta) meta.textContent = 'Seite aus der Mitte wählen';
    if (body) {
        body.innerHTML =
            '<p class="onv-preview__empty">Noch keine Vorschau – Abschnitt und Seite wählen.</p>';
    }
}

function setOnProgress(opts) {
    const box = $('onvProgress');
    const fill = $('onvProgressFill');
    const text = $('onvProgressText');
    if (!box || !text) return;
    if (opts && opts.hidden) {
        box.hidden = true;
        box.removeAttribute('data-state');
        if (fill) fill.style.width = '0%';
        text.textContent = '';
        return;
    }
    box.hidden = false;
    const state = (opts && opts.state) || '';
    if (state) box.setAttribute('data-state', state);
    else box.removeAttribute('data-state');
    if (fill && opts && Number.isFinite(opts.pct)) {
        fill.style.width = Math.max(0, Math.min(100, opts.pct)) + '%';
    }
    text.textContent = (opts && opts.message) || '';
}

function renderSrcTree() {
    const tree = ui.onSrcTree;
    const body = $('onvSrcTree');
    const footer = $('onvSrcFooter');
    if (!body) return;

    if (!tree) {
        body.innerHTML =
            '<div class="onenote-nav__empty">„Vorlagen laden“ – dann Abschnitte anhaken.</div>';
        if (footer) footer.textContent = 'Noch nichts geladen.';
        updateNavButtons();
        return;
    }

    const parts = [];

    function sectionRow(sec, depth) {
        const color = onenoteTabColor(sec.id || sec.displayName);
        const checked = ui.onSrcChecked.has(sec.id);
        const isPrev = ui.previewSectionId === sec.id;
        parts.push(
            '<div class="onenote-row' +
                (checked ? ' is-checked' : '') +
                (isPrev ? ' is-preview' : '') +
                '" style="--on-depth:' +
                depth +
                '" data-on-section="' +
                escapeHtml(sec.id) +
                '" role="treeitem">' +
                '<input type="checkbox" class="onenote-row__check" data-on-check="' +
                escapeHtml(sec.id) +
                '"' +
                (checked ? ' checked' : '') +
                ' aria-label="' +
                escapeHtml(sec.displayName) +
                '">' +
                '<span class="onenote-tab" style="--on-tab:' +
                color +
                '" aria-hidden="true"></span>' +
                '<span class="onenote-row__label">' +
                escapeHtml(sec.displayName) +
                '</span>' +
                '</div>'
        );
    }

    (tree.sections || []).forEach((sec) => sectionRow(sec, 0));

    (tree.groups || []).forEach((g) => {
        const isOpen = ui.onSrcExpanded.has(g.id);
        parts.push(
            '<div class="onenote-row onenote-row--group" style="--on-depth:0" role="treeitem">' +
                '<button type="button" class="onenote-row__toggle" data-on-toggle="' +
                escapeHtml(g.id) +
                '" aria-expanded="' +
                (isOpen ? 'true' : 'false') +
                '" title="' +
                (isOpen ? 'Zuklappen' : 'Aufklappen') +
                '"><i class="bi bi-caret-' +
                (isOpen ? 'down' : 'right') +
                '-fill" aria-hidden="true"></i></button>' +
                '<span class="onenote-row__label">' +
                escapeHtml(g.displayName) +
                '</span></div>'
        );
        if (isOpen) {
            (g.sections || []).forEach((sec) => sectionRow(sec, 1));
        }
    });

    if (!parts.length) {
        body.innerHTML = '<div class="onenote-nav__empty">Dieses Notizbuch ist leer.</div>';
    } else {
        body.innerHTML = parts.join('');
    }

    const counts = countTreeSections(tree);
    if (footer) {
        footer.textContent =
            counts.sections +
            ' Abschnitt(e) · ' +
            counts.groups +
            ' Gruppe(n) · ' +
            ui.onSrcChecked.size +
            ' ausgewählt';
    }
    updateNavButtons();
}

async function loadPreview(sectionId) {
    const sid = String(sectionId || '').trim();
    if (!sid) return;
    ui.previewSectionId = sid;
    ui.previewPageId = '';
    renderSrcTree();

    const name = findSectionName(ui.onSrcTree, sid) || 'Abschnitt';
    const pagesTitle = $('onvPagesTitle');
    const pagesMeta = $('onvPagesMeta');
    const title = $('onvPreviewTitle');
    const meta = $('onvPreviewMeta');
    const pagesEl = $('onvPreviewPages');
    const body = $('onvPreviewBody');
    if (pagesTitle) pagesTitle.textContent = name;
    if (pagesMeta) pagesMeta.textContent = 'Seiten werden geladen …';
    if (title) title.textContent = 'Vorschau';
    if (meta) meta.textContent = 'Seite wählen';
    if (pagesEl) pagesEl.innerHTML = '';
    if (body) {
        body.innerHTML = '<p class="onv-preview__empty">Lade Seitenliste …</p>';
    }

    try {
        const pages = await listSectionPages(sid, sourceScope());
        if (ui.previewSectionId !== sid) return;
        if (!pages.length) {
            if (pagesMeta) pagesMeta.textContent = 'Keine Seiten';
            if (meta) meta.textContent = 'Leerer Abschnitt';
            if (body) {
                body.innerHTML =
                    '<p class="onv-preview__empty">Dieser Abschnitt enthält noch keine Seiten.</p>';
            }
            return;
        }
        ui.previewPages = pages;
        if (pagesEl) {
            pagesEl.innerHTML = pages
                .map(
                    (p, i) =>
                        '<li><button type="button" data-on-page="' +
                        escapeHtml(p.id) +
                        '"' +
                        (i === 0 ? ' class="is-active"' : '') +
                        '>' +
                        escapeHtml(p.title) +
                        '</button></li>'
                )
                .join('');
        }
        if (pagesMeta) pagesMeta.textContent = pages.length + ' Seite(n)';
        if (meta) meta.textContent = 'Seite 1 von ' + pages.length;
        await showPagePreview(pages[0].id, pages[0].title);
    } catch (err) {
        if (pagesMeta) pagesMeta.textContent = 'Laden fehlgeschlagen';
        if (meta) meta.textContent = 'Vorschau fehlgeschlagen';
        if (body) {
            body.innerHTML =
                '<p class="onv-preview__empty">' +
                escapeHtml((err && err.message) || String(err)) +
                '</p>';
        }
        log('Vorschau: ' + ((err && err.message) || err));
    }
}

async function showPagePreview(pageId, pageTitle) {
    const pid = String(pageId || '').trim();
    if (!pid) return;
    ui.previewPageId = pid;
    const pagesEl = $('onvPreviewPages');
    if (pagesEl) {
        pagesEl.querySelectorAll('[data-on-page]').forEach((btn) => {
            btn.classList.toggle('is-active', btn.getAttribute('data-on-page') === pid);
        });
    }
    const title = $('onvPreviewTitle');
    const meta = $('onvPreviewMeta');
    if (title) title.textContent = pageTitle || 'Vorschau';
    if (meta) {
        const n = (ui.previewPages || []).length;
        const idx = (ui.previewPages || []).findIndex((p) => p.id === pid);
        meta.textContent =
            n && idx >= 0 ? 'Seite ' + (idx + 1) + ' von ' + n : 'HTML-Vorschau';
    }
    const body = $('onvPreviewBody');
    if (body) body.innerHTML = '<p class="onv-preview__empty">Lade HTML-Vorschau …</p>';
    try {
        let html = '';
        try {
            const content = await getPageContent(pid, sourceScope());
            html = (content && content.html) || '';
        } catch {
            html = '';
        }
        if (ui.previewPageId !== pid) return;

        if (html) {
            const pageMeta = (ui.previewPages || []).find((p) => p.id === pid);
            const openUrl = pageMeta && pageMeta.webUrl ? String(pageMeta.webUrl) : '';
            if (body) {
                body.innerHTML =
                    '<div class="onv-preview__frame-wrap">' +
                    '<iframe class="onv-preview__frame" title="OneNote-Seitenvorschau" sandbox="allow-same-origin"></iframe>' +
                    '</div>' +
                    '<p class="onv-preview__note">HTML-Vorschau inkl. Tabellen. Eingebettete Bilder können fehlen (Auth).' +
                    (openUrl
                        ? ' <a href="' +
                          escapeHtml(openUrl) +
                          '" target="_blank" rel="noopener">In OneNote öffnen</a>.'
                        : '') +
                    '</p>';
                const iframe = body.querySelector('iframe');
                if (iframe) iframe.srcdoc = buildPreviewDocument(html);
            }
            return;
        }

        const prev = await getPagePreview(pid, sourceScope());
        if (ui.previewPageId !== pid) return;
        const text = (prev && prev.previewText) || '';
        const img = (prev && prev.previewImageUrl) || '';
        if (body) {
            body.innerHTML =
                '<div class="onv-preview__frame-wrap"><div class="onv-preview__fallback">' +
                (img
                    ? '<img class="onv-preview__thumb" src="' +
                      escapeHtml(img) +
                      '" alt="Seitenvorschau">'
                    : '') +
                (text
                    ? escapeHtml(text)
                    : '<em style="color:#605e5c;">Keine Vorschau für „' +
                      escapeHtml(pageTitle || 'Seite') +
                      '“.</em>') +
                '</div></div>' +
                '<p class="onv-preview__note">Nur Text-Snippet. Für Tabellen: HTML-Content war nicht verfügbar.</p>';
        }
    } catch (err) {
        if (body) {
            body.innerHTML =
                '<p class="onv-preview__empty">' +
                escapeHtml((err && err.message) || String(err)) +
                '</p>';
        }
    }
}

function renderTeamList() {
    const ul = $('onvTeamList');
    if (!ul) {
        updateNavButtons();
        return;
    }
    if (!ui.selectedTeams.length) {
        ul.innerHTML = '';
        updateNavButtons();
        return;
    }
    ul.innerHTML = ui.selectedTeams
        .map((t, idx) => {
            const warn = t.warn
                ? '<div class="onv-chip__warn"><i class="bi bi-exclamation-triangle"></i> ' +
                  escapeHtml(t.warn) +
                  '</div>'
                : '';
            const nb = t.loading
                ? 'Notizbuch wird geladen …'
                : t.notebookName
                  ? 'Kursnotizbuch: ' + t.notebookName
                  : 'Kein Notizbuch';
            return (
                '<li class="onv-chip' +
                (t.warn ? ' is-warn' : '') +
                '"><div><strong>' +
                escapeHtml(t.teamName) +
                '</strong><span>' +
                escapeHtml(nb) +
                (t.mailNickname ? ' · ' + escapeHtml(t.mailNickname) : '') +
                '</span></div>' +
                '<button type="button" class="btn btn-sm alt" data-onv-remove-team="' +
                idx +
                '" title="Entfernen"><i class="bi bi-x-lg"></i></button>' +
                warn +
                '</li>'
            );
        })
        .join('');
    updateNavButtons();
}

async function addTeam(team) {
    const id = String((team && team.id) || '').trim();
    if (!id) return;
    if (ui.selectedTeams.some((t) => t.teamId === id)) {
        toast('Team ist schon ausgewählt.');
        return;
    }
    const entry = {
        teamId: id,
        teamName: String((team && team.displayName) || id),
        mailNickname: String((team && team.mailNickname) || ''),
        notebookId: '',
        notebookName: '',
        loading: true,
        warn: ''
    };
    ui.selectedTeams.push(entry);
    renderTeamList();
    toast('Team hinzugefügt: ' + entry.teamName);
    log('Team + ' + entry.teamName);

    try {
        const notebooks = await listOnenoteNotebooks(entry.teamId);
        const nb = pickClassNotebook(notebooks, entry.teamName);
        if (!nb) {
            entry.warn = 'Kein Kursnotizbuch gefunden – bitte in Teams anlegen.';
            entry.notebookId = '';
            entry.notebookName = '';
            log('Kein Notizbuch für ' + entry.teamName);
        } else {
            entry.notebookId = nb.id;
            entry.notebookName = nb.displayName || 'Kursnotizbuch';
            entry.warn = '';
            log('Notizbuch „' + entry.notebookName + '“ für ' + entry.teamName);
        }
    } catch (err) {
        entry.warn = (err && err.message) || String(err);
        log('Notizbücher ' + entry.teamName + ': ' + entry.warn);
    } finally {
        entry.loading = false;
        renderTeamList();
        /* Verteilerliste invalidieren, wenn Teams geändert werden */
        ui.onTargets = [];
        renderTargets();
    }
}

function removeTeam(idx) {
    if (!Number.isFinite(idx) || idx < 0 || idx >= ui.selectedTeams.length) return;
    const removed = ui.selectedTeams.splice(idx, 1)[0];
    ui.onTargets = ui.onTargets.filter((t) => t.teamId !== removed.teamId);
    renderTeamList();
    renderTargets();
    toast('Team entfernt.');
}

function getDestKind() {
    const checked = document.querySelector('input[name="onvDestKind"]:checked');
    const v = checked ? checked.value : 'contentLibrary';
    if (v === 'teacherOnly' || v === 'collaboration') return v;
    return 'contentLibrary';
}

function kindLabel(kind) {
    if (kind === 'teacherOnly') return 'Nur für Lehrer';
    if (kind === 'collaboration') return 'Zusammenarbeit';
    return 'Inhaltsbibliothek';
}

function renderTargets() {
    const ul = $('onvTargetList');
    if (!ul) {
        updateNavButtons();
        return;
    }
    if (!ui.onTargets.length) {
        ul.innerHTML = '';
        updateNavButtons();
        return;
    }
    ul.innerHTML = ui.onTargets
        .map((t, idx) => {
            const st = t.status
                ? '<div class="onv-chip__status' +
                  (t.statusKind ? ' is-' + t.statusKind : '') +
                  '">' +
                  escapeHtml(t.status) +
                  '</div>'
                : '';
            return (
                '<li class="onv-chip"><div><strong>' +
                escapeHtml(t.teamName) +
                '</strong><span>' +
                escapeHtml(t.notebookName) +
                ' → ' +
                escapeHtml(t.sectionGroupName) +
                '</span></div>' +
                '<button type="button" class="btn btn-sm alt" data-onv-remove-target="' +
                idx +
                '" title="Entfernen"><i class="bi bi-x-lg"></i></button>' +
                st +
                '</li>'
            );
        })
        .join('');
    updateNavButtons();
}

async function buildTargets() {
    if (ui.buildingTargets) return;
    const teams = teamsWithNotebook();
    if (!teams.length) {
        toast('Zuerst Teams mit Kursnotizbuch in Schritt 2 wählen.');
        return;
    }
    ui.buildingTargets = true;
    ui.destKind = getDestKind();
    const btn = $('onvBtnBuildTargets');
    if (btn) btn.disabled = true;
    ui.onTargets = [];
    renderTargets();
    log('Baue Verteilerliste (' + kindLabel(ui.destKind) + ') …');

    try {
        for (const team of teams) {
            try {
                const tree = await loadOnenoteNotebookTree(team.notebookId, team.teamId);
                const group = pickSectionGroup(tree.groups || [], ui.destKind);
                if (!group) {
                    log(
                        'Keine Abschnittsgruppe für ' +
                            team.teamName +
                            ' (' +
                            kindLabel(ui.destKind) +
                            ')'
                    );
                    toast('Keine passende Abschnittsgruppe: ' + team.teamName);
                    continue;
                }
                const key = team.teamId + '|' + team.notebookId + '|' + group.id;
                ui.onTargets.push({
                    key,
                    teamId: team.teamId,
                    teamName: team.teamName,
                    notebookId: team.notebookId,
                    notebookName: team.notebookName,
                    sectionGroupId: group.id,
                    sectionGroupName: group.displayName || kindLabel(ui.destKind)
                });
                log(
                    'Ziel + ' +
                        team.teamName +
                        ' → ' +
                        team.notebookName +
                        ' / ' +
                        group.displayName
                );
            } catch (err) {
                log(
                    'Ziel ' +
                        team.teamName +
                        ': ' +
                        ((err && err.message) || err)
                );
                toast((err && err.message) || String(err));
            }
        }
        renderTargets();
        if (ui.onTargets.length) {
            toast(ui.onTargets.length + ' Ziel(e) in der Verteilerliste.');
        } else {
            toast('Keine Ziele gefunden – Abschnittsgruppen prüfen.');
        }
    } finally {
        ui.buildingTargets = false;
        if (btn) btn.disabled = false;
        updateNavButtons();
    }
}

function renderSummary() {
    const title = $('onvSummaryTitle');
    const body = $('onvSummaryBody');
    const sections = [...ui.onSrcChecked].map((id) => findSectionName(ui.onSrcTree, id) || id);
    const n = sections.length;
    const m = ui.onTargets.length;
    if (title) {
        title.textContent = n + ' Abschnitt' + (n === 1 ? '' : 'e') + ' × ' + m + ' Ziel' + (m === 1 ? '' : 'e');
    }
    if (!body) return;
    const secList =
        '<p style="margin:0 0 6px;font-weight:700;color:var(--heading);">Abschnitte</p><ul>' +
        sections.map((s) => '<li>' + escapeHtml(s) + '</li>').join('') +
        '</ul>';
    const destList =
        '<p style="margin:12px 0 6px;font-weight:700;color:var(--heading);">Ziele</p><ul>' +
        ui.onTargets
            .map(
                (t) =>
                    '<li>' +
                    escapeHtml(t.teamName) +
                    ' → ' +
                    escapeHtml(t.sectionGroupName) +
                    '</li>'
            )
            .join('') +
        '</ul>';
    body.innerHTML = secList + destList;
}

/* Wizard-Binds → onenote-verteilung-steps.js */
function init() {
    if (!$('onvStep1')) return;
    initOnenoteSteps({
        $, toast, log, showStep, canEnterStep, updateNavButtons, setSrcMode, selectNotebook,
        renderSrcTree, renderTeamList, renderTargets, buildTargets, renderSummary,
        loadPreview, showPagePreview, clearPreview, setOnProgress, addTeam, removeTeam,
        getDestKind, kindLabel, sourceScope, updateSrcPickVisibility, updatePublishBtn,
        refreshPublishedMarks, renderSrcTeamSelected, renderSrcSiteSelected, renderNotebookShelf,
        fillOnSelect, countTreeSections, findSectionName, teamsWithNotebook, buildPreviewDocument,
        canPublishSnapshots
    });
    setSrcMode('central');
    bindWizardNav();
    bindStep1();
    bindStep2();
    bindStep3();
    bindStep4();
    renderSrcTree();
    renderTeamList();
    renderTargets();
    showStep(1);
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
