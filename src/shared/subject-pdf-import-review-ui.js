/**
 * Subject-PDF im Schulregister: Auswahl-Dialog (nur Fächer, wie WebUntis-Wizard).
 */
import { escapeHtml } from './utils/strings.js';
import {
    applyRecommendedSubjectImport,
    buildSubjectPdfReviewPreview,
    rowMatchesFilter,
    subjectImportStats,
    subjectRowPassesQuickFilter
} from './webuntis-stammdaten-wizard-logic.js';

let dialogEl = null;
let state = null;

function ensureDialog() {
    if (dialogEl) return dialogEl;
    dialogEl = document.createElement('dialog');
    dialogEl.className = 'tenant-subject-pdf-review';
    dialogEl.setAttribute('aria-labelledby', 'tenantSubjectPdfReviewTitle');
    dialogEl.innerHTML =
        '<form method="dialog" class="tenant-subject-pdf-review__form">' +
        '<header class="tenant-subject-pdf-review__head">' +
        '<h2 id="tenantSubjectPdfReviewTitle">WebUntis Fächer-PDF</h2>' +
        '<p class="muted tenant-subject-pdf-review__file" id="tenantSubjectPdfReviewFile"></p>' +
        '</header>' +
        '<div class="wu-subject-import tenant-subject-pdf-review__panel">' +
        '<div class="wu-subject-import__hero" id="tenantSubjectPdfReviewHero" aria-live="polite"></div>' +
        '<div class="wu-subject-import__toolbar">' +
        '<div class="wu-subject-import__presets" role="group" aria-label="Schnellauswahl">' +
        '<button type="button" class="btn btn-sm btn-primary" data-subject-preset="recommended"><i class="bi bi-magic" aria-hidden="true"></i> Empfohlene Auswahl</button>' +
        '<button type="button" class="btn btn-sm" data-subject-preset="teaching"><i class="bi bi-mortarboard" aria-hidden="true"></i> Unterrichtsfächer</button>' +
        '<button type="button" class="btn btn-sm" data-subject-preset="all">Alle übernehmen</button>' +
        '<button type="button" class="btn btn-sm" data-subject-preset="none">Keine übernehmen</button>' +
        '</div>' +
        '<div class="wu-subject-import__view" role="group" aria-label="Listenansicht">' +
        '<button type="button" class="btn btn-sm is-active" data-subject-view="import">Nur Übernahme</button>' +
        '<button type="button" class="btn btn-sm" data-subject-view="all">Alle Fächer</button>' +
        '</div>' +
        '</div>' +
        '<p class="muted wu-stammdaten-wizard__filter-hint" id="tenantSubjectPdfReviewHint" hidden></p>' +
        '<div class="wu-import-filterbar tenant-subject-pdf-review__filterbar">' +
        '<input type="search" id="tenantSubjectPdfReviewSearch" placeholder="Kürzel oder Bezeichnung suchen …" autocomplete="off" />' +
        '<label class="tenant-subject-pdf-review__quick">' +
        '<span class="sr-only">Filter</span>' +
        '<select id="tenantSubjectPdfReviewQuickFilter" aria-label="Filter">' +
        '<option value="all">Alle anzeigen</option>' +
        '<option value="teaching">Nur Unterricht (ohne Verwaltung)</option>' +
        '<option value="numbered">Nur mit Nummer (1–20)</option>' +
        '</select>' +
        '</label>' +
        '<label class="slg-deviation-section__select-all">' +
        '<input type="checkbox" id="tenantSubjectPdfReviewSelectVisible" checked />' +
        '<span id="tenantSubjectPdfReviewSelectVisibleLabel">Sichtbare für Übernahme markieren</span>' +
        '</label>' +
        '</div>' +
        '<div class="wu-import-table-wrap tenant-subject-pdf-review__table-wrap">' +
        '<table class="teachers-table wu-import-table">' +
        '<thead><tr>' +
        '<th scope="col" style="width:2.75rem;">Übernehmen</th>' +
        '<th scope="col">Kürzel</th>' +
        '<th scope="col">Bezeichnung</th>' +
        '</tr></thead>' +
        '<tbody id="tenantSubjectPdfReviewTableBody"></tbody>' +
        '</table>' +
        '</div>' +
        '</div>' +
        '<footer class="tenant-subject-pdf-review__foot">' +
        '<p class="muted tenant-subject-pdf-review__status" id="tenantSubjectPdfReviewStatus"></p>' +
        '<div class="tenant-subject-pdf-review__actions">' +
        '<button type="submit" class="btn btn-primary" value="apply">Auswahl übernehmen</button>' +
        '<button type="button" class="btn" id="tenantSubjectPdfReviewCancel">Abbrechen</button>' +
        '</div>' +
        '</footer>' +
        '</form>';
    document.body.appendChild(dialogEl);

    const form = dialogEl.querySelector('form');
    if (form) {
        form.addEventListener('submit', function (ev) {
            if (!state || !state.preview) return;
            const n = (state.preview.subjects || []).filter(function (r) {
                return r.selected;
            }).length;
            if (!n) {
                ev.preventDefault();
                const status = dialogEl.querySelector('#tenantSubjectPdfReviewStatus');
                if (status) {
                    status.textContent =
                        'Bitte mindestens ein Fach markieren (z. B. „Empfohlene Auswahl“) oder Abbrechen.';
                }
            }
        });
    }

    dialogEl.querySelectorAll('[data-subject-preset]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            applyPreset(btn.getAttribute('data-subject-preset'));
        });
    });
    dialogEl.querySelectorAll('[data-subject-view]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            state.viewMode = btn.getAttribute('data-subject-view') === 'all' ? 'all' : 'import';
            syncViewButtons();
            render();
        });
    });
    const search = dialogEl.querySelector('#tenantSubjectPdfReviewSearch');
    if (search) search.addEventListener('input', function () {
        state.filterQuery = search.value;
        render();
    });
    const quick = dialogEl.querySelector('#tenantSubjectPdfReviewQuickFilter');
    if (quick) {
        quick.addEventListener('change', function () {
            state.quickFilter = quick.value;
            render();
        });
    }
    const selAll = dialogEl.querySelector('#tenantSubjectPdfReviewSelectVisible');
    if (selAll) {
        selAll.addEventListener('change', function () {
            const visible = visibleRows();
            const on = !!selAll.checked;
            visible.forEach(function (r) {
                r.selected = on;
            });
            render();
        });
    }
    const cancel = dialogEl.querySelector('#tenantSubjectPdfReviewCancel');
    if (cancel) {
        cancel.addEventListener('click', function () {
            dialogEl.close('cancel');
        });
    }
    dialogEl.addEventListener('close', function () {
        if (state && state.resolve) {
            const v = dialogEl.returnValue;
            if (v === 'apply' && state.preview) {
                const picked = (state.preview.subjects || []).filter(function (r) {
                    return r.selected;
                });
                state.resolve({
                    subjects: picked.map(function (r) {
                        return { code: r.code, name: r.name, admin: r.admin };
                    })
                });
            } else {
                state.resolve(null);
            }
            state = null;
        }
    });

    return dialogEl;
}

function syncViewButtons() {
    if (!dialogEl || !state) return;
    dialogEl.querySelectorAll('[data-subject-view]').forEach(function (b) {
        const mode = b.getAttribute('data-subject-view') === 'all' ? 'all' : 'import';
        b.classList.toggle('is-active', mode === state.viewMode);
    });
}

function applyPreset(preset) {
    if (!state || !state.preview) return;
    const rows = state.preview.subjects || [];
    if (preset === 'recommended') {
        applyRecommendedSubjectImport(rows);
    } else if (preset === 'teaching') {
        rows.forEach(function (r) {
            r.selected = !r.admin;
        });
    } else if (preset === 'all') {
        rows.forEach(function (r) {
            r.selected = true;
        });
    } else if (preset === 'none') {
        rows.forEach(function (r) {
            r.selected = false;
        });
    }
    render();
}

function visibleRows() {
    if (!state || !state.preview) return [];
    const rows = state.preview.subjects || [];
    const q = state.filterQuery || '';
    const qf = state.quickFilter || 'all';
    return rows.filter(function (r) {
        if (state.viewMode === 'import' && !r.selected) return false;
        if (!subjectRowPassesQuickFilter(r, qf)) return false;
        return rowMatchesFilter(r, q, ['code', 'name']);
    });
}

function renderHero() {
    const host = dialogEl.querySelector('#tenantSubjectPdfReviewHero');
    if (!host || !state || !state.preview) return;
    const stats = subjectImportStats(state.preview.subjects || []);
    const pct = stats.total ? Math.round((stats.selected / stats.total) * 100) : 0;
    host.innerHTML =
        '<div class="wu-subject-import__hero-main">' +
        '<p class="wu-subject-import__hero-count"><strong>' +
        escapeHtml(String(stats.selected)) +
        '</strong><span class="wu-subject-import__hero-of">von ' +
        escapeHtml(String(stats.total)) +
        ' Fächern werden übernommen</span></p>' +
        '<p class="muted wu-subject-import__hero-sub">' +
        escapeHtml(String(stats.excluded)) +
        ' nicht übernommen · ' +
        escapeHtml(String(pct)) +
        '% der PDF-Liste</p>' +
        '</div>' +
        '<div class="wu-subject-import__hero-meter" aria-hidden="true">' +
        '<div class="wu-subject-import__hero-meter-fill" style="width:' +
        pct +
        '%"></div></div>';
}

function render() {
    if (!dialogEl || !state || !state.preview) return;
    renderHero();
    syncViewButtons();
    const hint = dialogEl.querySelector('#tenantSubjectPdfReviewHint');
    if (hint) {
        const text = state.preview.subjectImportHint || '';
        if (text) {
            hint.hidden = false;
            hint.textContent = text;
        } else {
            hint.hidden = true;
            hint.textContent = '';
        }
    }
    const status = dialogEl.querySelector('#tenantSubjectPdfReviewStatus');
    const all = state.preview.subjects || [];
    const sel = all.filter(function (r) {
        return r.selected;
    }).length;
    const visible = visibleRows();
    if (status) {
        status.textContent =
            sel + ' Fächer für die Fächerliste vorgemerkt · ' + visible.length + ' Zeilen in der Tabelle';
    }
    const body = dialogEl.querySelector('#tenantSubjectPdfReviewTableBody');
    if (!body) return;
    body.replaceChildren();
    if (!visible.length) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.colSpan = 3;
        td.className = 'muted';
        if (state.viewMode === 'import' && sel === 0) {
            td.textContent =
                'Noch keine Fächer markiert. „Empfohlene Auswahl“ nutzen oder „Alle Fächer“ anzeigen und Häkchen setzen.';
        } else {
            td.textContent = 'Keine Treffer – Suche oder Filter anpassen.';
        }
        tr.appendChild(td);
        body.appendChild(tr);
        return;
    }
    visible.forEach(function (row) {
        const tr = document.createElement('tr');
        tr.classList.add(row.selected ? 'wu-import-row--import-yes' : 'wu-import-row--import-no');
        const tdCb = document.createElement('td');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = !!row.selected;
        cb.setAttribute('aria-label', 'Fach ' + row.code + ' übernehmen');
        cb.addEventListener('change', function () {
            row.selected = !!cb.checked;
            render();
        });
        tdCb.appendChild(cb);
        tr.appendChild(tdCb);
        const tdCode = document.createElement('td');
        tdCode.textContent = row.code || '';
        if (row.admin) tdCode.title = 'Verwaltungsfach';
        tr.appendChild(tdCode);
        const tdName = document.createElement('td');
        tdName.textContent = row.name || '';
        tr.appendChild(tdName);
        body.appendChild(tr);
    });
}

/**
 * @param {{ subjects: object[], sourceFileName?: string, meta?: object }} options
 * @returns {Promise<{ subjects: object[] }|null>}
 */
export function openSubjectPdfImportReview(options) {
    const raw = (options && options.subjects) || [];
    if (!raw.length) return Promise.resolve(null);
    const dlg = ensureDialog();
    const preview = buildSubjectPdfReviewPreview(raw, {
        sourceFileName: options && options.sourceFileName,
        deselectThreshold: 25
    });
    const fileEl = dlg.querySelector('#tenantSubjectPdfReviewFile');
    if (fileEl) {
        fileEl.textContent = options && options.sourceFileName ? 'Datei: ' + options.sourceFileName : '';
    }
    const search = dlg.querySelector('#tenantSubjectPdfReviewSearch');
    if (search) search.value = '';
    const quick = dlg.querySelector('#tenantSubjectPdfReviewQuickFilter');
    if (quick) quick.value = 'all';

    return new Promise(function (resolve) {
        state = {
            preview: preview,
            viewMode: (preview.subjects || []).length > 25 ? 'import' : 'all',
            filterQuery: '',
            quickFilter: 'all',
            resolve: resolve
        };
        render();
        if (typeof dlg.showModal === 'function') {
            dlg.showModal();
        } else {
            dlg.setAttribute('open', 'open');
        }
    });
}
