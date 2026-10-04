/**
 * Vollseite: WebUntis-Stammdaten importieren (Upload, Kacheln, Prüfung, Übernahme per SessionStorage).
 */
import {
    WIZARD_TILES,
    buildWizardPreview,
    compileWizardApplyPayload,
    rowMatchesFilter,
    applyRecommendedSubjectImport,
    subjectImportStats
} from '../../shared/webuntis-stammdaten-wizard-logic.js';
import { stashWebuntisImportPayload, resolveReturnUrl } from '../../shared/webuntis-stammdaten-import-handoff.js';
import { escapeHtml } from '../../shared/utils/strings.js';
import { showToast } from '../../shared/utils/dom.js';
import { mountWebuntisSubjectCleanup } from '../../shared/webuntis-subject-cleanup-ui.js';
import '../kursteams/kursteam-subject-filter-logic.js';

/** @type {object|null} */
let previewState = null;
let activeTile = 'teachers';
let phase = 'upload';
/** @type {'import'|'all'} */
let subjectViewMode = 'import';
/** @type {File[]} */
let pendingFiles = [];
/** @type {ReturnType<typeof mountWebuntisSubjectCleanup>|null} */
let subjectCleanup = null;

const fromParam = () => {
    const q = new URLSearchParams(window.location.search).get('from');
    return q ? String(q).toLowerCase() : 'tenant';
};

function toast(msg, opts) {
    showToast(msg, opts || { durationMs: 4500 });
}

/** Fortschritts-Overlay beim Auswerten */
function createImportParseProgress() {
    const root = document.getElementById('wuImportParseProgress');
    const fill = document.getElementById('wuImportProgressFill');
    const pctEl = document.getElementById('wuImportProgressPct');
    const detail = document.getElementById('wuImportProgressDetail');
    const eta = document.getElementById('wuImportProgressEta');
    const stepsEl = document.getElementById('wuImportProgressSteps');
    const btnNext = document.getElementById('wuStammdatenBtnNext');
    /** @type {Map<string, { weight: number, el: HTMLElement }>} */
    const stepMap = new Map();
    let startMs = 0;
    let doneWeight = 0;
    let totalWeight = 0;

    function setBar() {
        const pct = totalWeight > 0 ? Math.min(100, Math.round((doneWeight / totalWeight) * 100)) : 0;
        if (fill) fill.style.width = pct + '%';
        if (pctEl) pctEl.textContent = pct + ' %';
        if (!eta || !startMs) return;
        if (pct >= 100) {
            eta.textContent = 'Fast fertig …';
            return;
        }
        if (doneWeight <= 0 || pct < 3) {
            eta.textContent = 'Dauer hängt von Dateigröße und PDF-Anzahl ab …';
            return;
        }
        const elapsed = Date.now() - startMs;
        const remaining = Math.max(0, Math.round((elapsed / doneWeight) * (totalWeight - doneWeight) / 1000));
        if (remaining < 5) eta.textContent = 'Noch wenige Sekunden …';
        else if (remaining < 90) eta.textContent = 'Noch etwa ' + remaining + ' Sekunden …';
        else eta.textContent = 'Noch etwa ' + Math.ceil(remaining / 60) + ' Minute(n) …';
    }

    function setStepState(id, state) {
        const rec = stepMap.get(id);
        if (!rec || !rec.el) return;
        rec.el.classList.remove('is-pending', 'is-active', 'is-done');
        rec.el.classList.add('is-' + state);
        const icon = rec.el.querySelector('.bi');
        if (!icon) return;
        if (state === 'done') icon.className = 'bi bi-check-circle-fill';
        else if (state === 'active') icon.className = 'bi bi-arrow-repeat';
        else icon.className = 'bi bi-circle';
    }

    return {
        begin: function (steps) {
            if (!root) return;
            stepMap.clear();
            doneWeight = 0;
            totalWeight = 0;
            startMs = Date.now();
            (steps || []).forEach(function (s) {
                totalWeight += s.weight || 1;
            });
            if (stepsEl) {
                stepsEl.innerHTML = (steps || [])
                    .map(function (s) {
                        return (
                            '<li class="is-pending" data-step-id="' +
                            escapeHtml(s.id) +
                            '"><i class="bi bi-circle" aria-hidden="true"></i><span>' +
                            escapeHtml(s.label) +
                            '</span></li>'
                        );
                    })
                    .join('');
                stepsEl.querySelectorAll('[data-step-id]').forEach(function (li) {
                    const id = li.getAttribute('data-step-id');
                    const step = (steps || []).find(function (x) {
                        return x.id === id;
                    });
                    if (step) stepMap.set(id, { weight: step.weight || 1, el: li });
                });
            }
            if (detail) detail.textContent = 'Vorbereitung …';
            if (eta) eta.textContent = '';
            setBar();
            root.hidden = false;
            document.body.classList.add('wu-import-parse-busy');
            if (btnNext) btnNext.disabled = true;
        },
        active: function (stepId, detailText) {
            stepMap.forEach(function (rec, id) {
                if (id === stepId) setStepState(id, 'active');
                else if (!rec.el.classList.contains('is-done')) setStepState(id, 'pending');
            });
            if (detailText && detail) detail.textContent = detailText;
            setBar();
        },
        complete: function (stepId) {
            const rec = stepMap.get(stepId);
            if (rec) {
                doneWeight += rec.weight;
                setStepState(stepId, 'done');
            }
            setBar();
        },
        end: function () {
            if (root) root.hidden = true;
            document.body.classList.remove('wu-import-parse-busy');
            if (btnNext) btnNext.disabled = false;
        }
    };
}

function getEmailOptsFromSettings() {
    const domain =
        typeof window.ms365SchoolDomainForEmail === 'function' ? window.ms365SchoolDomainForEmail() : '';
    return { domain: domain, pattern: 'vorname.nachname', firstNameMode: 'first' };
}

function getTeachersForClassEnrichment() {
    const load = window.ms365TenantSettingsLoad;
    if (typeof load !== 'function') return [];
    const s = load();
    return Array.isArray(s && s.teachers) ? s.teachers : [];
}

function getSchoolYearEndForClassPdf() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.load === 'function') {
            const c = window.ms365AppDataV2.load();
            const label = c && c.years && c.years.current;
            if (label && window.ms365SchoolYear && typeof window.ms365SchoolYear.parseSchoolYearStartYear === 'function') {
                const start = window.ms365SchoolYear.parseSchoolYearStartYear(label);
                if (isFinite(start)) return start + 1;
            }
        }
    } catch {
        /* ignore */
    }
    if (window.ms365SchoolYear && typeof window.ms365SchoolYear.schoolYearStartYear === 'function') {
        return window.ms365SchoolYear.schoolYearStartYear() + 1;
    }
    const y = new Date().getFullYear();
    return new Date().getMonth() < 8 ? y : y + 1;
}

function setPhase(next) {
    phase = next;
    const upload = document.getElementById('wuStammdatenWizardUpload');
    const review = document.getElementById('wuStammdatenWizardReview');
    const stepU = document.getElementById('wuImportStepUpload');
    const stepR = document.getElementById('wuImportStepReview');

    if (upload) upload.hidden = phase !== 'upload';
    if (review) review.hidden = phase !== 'review';

    if (stepU) {
        stepU.classList.toggle('is-active', phase === 'upload');
        stepU.classList.toggle('is-done', phase === 'review');
    }
    if (stepR) stepR.classList.toggle('is-active', phase === 'review');

    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function renderFileList() {
    const ul = document.getElementById('wuStammdatenFileList');
    if (!ul) return;
    if (!pendingFiles.length) {
        ul.innerHTML = '<li class="muted">Noch keine Dateien gewählt.</li>';
        return;
    }
    ul.innerHTML = pendingFiles
        .map(function (f) {
            return '<li><i class="bi bi-file-earmark"></i> ' + escapeHtml(f.name) + '</li>';
        })
        .join('');
}

function tileCounts() {
    if (!previewState) {
        return { subjects: 0, classes: 0, teachers: 0, students: 0, guardians: 0 };
    }
    return {
        subjects: (previewState.subjects || []).length,
        classes: (previewState.classes || []).length,
        teachers: (previewState.teachers || []).length,
        students: (previewState.students || []).length,
        guardians: (previewState.guardians || []).length
    };
}

function tileMeta(tile, n) {
    if (n > 0) return { sub: n + ' Datensätze', status: 'erkannt', statusClass: 'has-files' };
    if (tile.id === 'guardians') {
        return { sub: 'nur nicht zugeordnete', status: 'optional', statusClass: '' };
    }
    return { sub: 'optional', status: 'keine Datei', statusClass: '' };
}

function wireAreaTabs() {
    if (wireAreaTabs._done) return;
    const host = document.getElementById('wuImportAreaTabs');
    if (!host) return;
    wireAreaTabs._done = true;
    host.addEventListener('click', function (ev) {
        const btn = ev.target.closest('[data-tile]');
        if (!btn || phase !== 'review') return;
        activeTile = btn.getAttribute('data-tile') || 'teachers';
        renderAreaTabs();
        renderReviewTable();
    });
}

function renderAreaTabs() {
    const host = document.getElementById('wuImportAreaTabs');
    if (!host) return;
    const counts = tileCounts();
    const showTabs = phase === 'review' && !!previewState;
    host.hidden = !showTabs;
    if (!showTabs) {
        host.innerHTML = '';
        return;
    }

    host.innerHTML = WIZARD_TILES.map(function (tile) {
        const n = counts[tile.id] || 0;
        const selected =
            previewState && Array.isArray(previewState[tile.id])
                ? previewState[tile.id].filter(function (r) {
                      return r.selected;
                  }).length
                : 0;
        const isActive = activeTile === tile.id;
        const badgeNum = tile.id === 'subjects' && n > 0 ? selected : n;
        const badgeTitle =
            tile.id === 'subjects' && n > 0 ? selected + ' von ' + n + ' werden übernommen' : selected + ' ausgewählt';
        const badge =
            n > 0
                ? '<span class="wu-import-tab-badge" title="' + escapeHtml(badgeTitle) + '">' + escapeHtml(String(badgeNum)) + '</span>'
                : '<span class="wu-import-tab-badge wu-import-tab-badge--muted">–</span>';
        return (
            '<button type="button" class="tab-btn' +
            (isActive ? ' active' : '') +
            '" role="tab" data-tile="' +
            tile.id +
            '" aria-selected="' +
            (isActive ? 'true' : 'false') +
            '" aria-controls="wuImportReviewBody" id="wuImportTab-' +
            tile.id +
            '">' +
            '<i class="bi ' +
            tile.icon +
            '" aria-hidden="true"></i>' +
            escapeHtml(tile.label) +
            badge +
            '</button>'
        );
    }).join('');

    const panel = document.getElementById('wuImportReviewBody');
    if (panel) panel.setAttribute('aria-labelledby', 'wuImportTab-' + activeTile);
}

function renderTiles() {
    renderAreaTabs();
}

function getRowsForTile(tileId) {
    if (!previewState) return [];
    return previewState[tileId] || [];
}

function filterQuery() {
    const el = document.getElementById('wuStammdatenFilter');
    return el ? el.value : '';
}

function visibleRowsForTile() {
    const rows = getRowsForTile(activeTile);
    const q = filterQuery();
    const keysByTile = {
        subjects: ['code', 'name'],
        classes: ['code', 'name', 'year', 'headName', 'headEmail'],
        teachers: ['code', 'name', 'email'],
        students: ['klasse', 'name', 'email', 'parentSummary'],
        guardians: ['name', 'email', 'phone', 'studentRef', 'note']
    };
    const keys = keysByTile[activeTile] || ['name'];
    return rows.filter(function (r) {
        if (activeTile === 'subjects' && subjectViewMode === 'import' && !r.selected) return false;
        return rowMatchesFilter(r, q, keys);
    });
}

function renderSubjectImportHero() {
    const host = document.getElementById('wuSubjectImportHero');
    if (!host || !previewState || activeTile !== 'subjects') return;
    const stats = subjectImportStats(previewState.subjects || []);
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
        '% der WebUntis-Liste</p>' +
        '</div>' +
        '<div class="wu-subject-import__hero-meter" aria-hidden="true">' +
        '<div class="wu-subject-import__hero-meter-fill" style="width:' +
        pct +
        '%"></div></div>';
}

function applySubjectPreset(preset) {
    if (!previewState || !previewState.subjects) return;
    const rows = previewState.subjects;
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
    if (subjectCleanup && typeof subjectCleanup.setExcluded === 'function') {
        const ex = rows
            .filter(function (r) {
                return !r.selected;
            })
            .map(function (r) {
                return String(r.code || '').toUpperCase();
            });
        subjectCleanup.setExcluded(ex);
    }
    renderSubjectImportHero();
    renderTiles();
    renderReviewTable();
}

function initSubjectCleanup() {
    if (subjectCleanup) return subjectCleanup;
    subjectCleanup = mountWebuntisSubjectCleanup({
        getSubjectRows: function () {
            return previewState && previewState.subjects ? previewState.subjects : [];
        },
        onExcludedChange: function (codes) {
            applySubjectExclusions(codes);
        }
    });
    return subjectCleanup;
}

function applySubjectExclusions(excludedCodes) {
    if (!previewState || !previewState.subjects) return;
    const ex = new Set(
        (excludedCodes || []).map(function (c) {
            return String(c || '').toUpperCase();
        })
    );
    previewState.subjects.forEach(function (r) {
        const code = String(r.code || '').toUpperCase();
        r.selected = !ex.has(code);
    });
    renderSubjectImportHero();
    renderTiles();
    if (phase === 'review' && activeTile === 'subjects') {
        const body = document.getElementById('wuStammdatenTableBody');
        if (body) renderReviewTable();
    }
}

function updateSubjectCleanupChrome() {
    const host = document.getElementById('wuStammdatenSubjectCleanup');
    const show = phase === 'review' && activeTile === 'subjects' && previewState;
    if (host) host.hidden = !show;
    if (show) initSubjectCleanup().refresh();
}

function updateSubjectImportChrome() {
    const panel = document.getElementById('wuSubjectImportPanel');
    const hint = document.getElementById('wuStammdatenFilterHint');
    const filterInput = document.getElementById('wuStammdatenFilter');
    const selectWrap = document.getElementById('wuStammdatenSelectAllWrap');
    const showSubjects = phase === 'review' && activeTile === 'subjects' && previewState;
    if (panel) panel.hidden = !showSubjects;
    if (showSubjects) {
        renderSubjectImportHero();
        document.querySelectorAll('[data-subject-view]').forEach(function (b) {
            b.classList.toggle('is-active', b.getAttribute('data-subject-view') === subjectViewMode);
        });
    }
    updateSubjectCleanupChrome();
    if (hint) {
        if (showSubjects && previewState && previewState.subjectImportHint) {
            hint.hidden = false;
            hint.textContent = previewState.subjectImportHint;
        } else {
            hint.hidden = true;
            hint.textContent = '';
        }
    }
    if (filterInput) {
        filterInput.placeholder =
            showSubjects ? 'In der Liste suchen (Kürzel, Bezeichnung) …' : 'Textsuche (Name, Kürzel, Klasse …)';
    }
    if (selectWrap) {
        const label = document.getElementById('wuStammdatenSelectAllLabel');
        if (label) {
            label.textContent = showSubjects
                ? 'Alle sichtbaren für Übernahme markieren'
                : 'Sichtbare für Übernahme markieren';
        }
    }
}

function renderReviewTable() {
    const head = document.getElementById('wuStammdatenTableHead');
    const body = document.getElementById('wuStammdatenTableBody');
    if (!head || !body) return;

    const tileDef = WIZARD_TILES.find(function (t) {
        return t.id === activeTile;
    });
    const titleEl = document.getElementById('wuImportReviewTitle');
    if (titleEl && tileDef) titleEl.textContent = 'Prüfung: ' + tileDef.label;
    const tabMeta = document.getElementById('wuImportTabMeta');

    const cols = {
        subjects: [
            ['', 'Übernehmen'],
            ['code', 'Kürzel'],
            ['name', 'Bezeichnung']
        ],
        classes: [
            ['', ''],
            ['code', 'Kürzel'],
            ['year', 'Abschluss'],
            ['name', 'Name'],
            ['headCode', 'KV-Kürzel'],
            ['headName', 'Klassenvorstand'],
            ['headEmail', 'KV-E-Mail']
        ],
        teachers: [
            ['', ''],
            ['code', 'Kürzel'],
            ['name', 'Name'],
            ['email', 'E-Mail']
        ],
        students: [
            ['', ''],
            ['klasse', 'Klasse'],
            ['name', 'Name'],
            ['email', 'E-Mail'],
            ['parentSummary', 'Eltern']
        ],
        guardians: [
            ['', ''],
            ['name', 'Name'],
            ['email', 'E-Mail'],
            ['phone', 'Telefon'],
            ['studentRef', 'Schüler (aus Export)'],
            ['note', 'Hinweis']
        ]
    };
    const spec = cols[activeTile] || cols.teachers;
    head.innerHTML =
        '<tr>' +
        spec
            .map(function (c) {
                if (!c[0]) return '<th scope="col" style="width:2.5rem;"></th>';
                return '<th scope="col">' + escapeHtml(c[1]) + '</th>';
            })
            .join('') +
        '</tr>';

    updateSubjectImportChrome();

    const visible = visibleRowsForTile();
    const total = getRowsForTile(activeTile).length;
    const statusEl = document.getElementById('wuStammdatenWizardStatus');
    const sel = getRowsForTile(activeTile).filter(function (r) {
        return r.selected;
    }).length;
    const statusLine =
        activeTile === 'subjects'
            ? sel + ' werden übernommen · ' + visible.length + ' in der Liste'
            : 'Angezeigt: ' + visible.length + ' von ' + total + ' · ausgewählt: ' + sel;
    if (statusEl && phase === 'review' && previewState) {
        statusEl.textContent = statusLine;
    }
    if (tabMeta && tileDef && phase === 'review') {
        const hint = tileMeta(tileDef, total);
        if (activeTile === 'subjects' && total > 0) {
            tabMeta.textContent = sel + ' von ' + total + ' Fächern für die Stammdaten vorgemerkt.';
        } else {
            tabMeta.textContent =
                total === 0 && hint.sub
                    ? tileDef.label + ' · ' + hint.sub + (hint.status ? ' · ' + hint.status : '')
                    : tileDef.label + ': ' + statusLine;
        }
    }

    body.replaceChildren();
    if (!visible.length) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.colSpan = spec.length;
        td.className = 'muted';
        if (activeTile === 'guardians') {
            td.textContent = 'Keine nicht zugeordneten Eltern – gut!';
        } else if (activeTile === 'subjects' && subjectViewMode === 'import' && sel === 0) {
            td.textContent =
                'Noch keine Fächer für die Übernahme markiert. „Empfohlene Auswahl“ nutzen oder auf „Alle Fächer“ wechseln und Häkchen setzen.';
        } else if (activeTile === 'subjects' && subjectViewMode === 'import') {
            td.textContent = 'Keine Treffer in der Übernahme-Liste – Suche anpassen oder „Alle Fächer“ anzeigen.';
        } else {
            td.textContent = 'Keine Einträge in dieser Kategorie (Filter oder Datei prüfen).';
        }
        tr.appendChild(td);
        body.appendChild(tr);
        return;
    }

    visible.forEach(function (row) {
        const tr = document.createElement('tr');
        if (activeTile === 'subjects') {
            tr.classList.add(row.selected ? 'wu-import-row--import-yes' : 'wu-import-row--import-no');
        }
        spec.forEach(function (c) {
            const td = document.createElement('td');
            if (!c[0]) {
                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.checked = !!row.selected;
                cb.disabled = activeTile === 'guardians';
                cb.addEventListener('change', function () {
                    row.selected = cb.checked;
                    if (activeTile === 'subjects') {
                        tr.classList.toggle('wu-import-row--import-yes', row.selected);
                        tr.classList.toggle('wu-import-row--import-no', !row.selected);
                        renderSubjectImportHero();
                    }
                    renderTiles();
                });
                td.appendChild(cb);
            } else {
                td.textContent = row[c[0]] != null ? String(row[c[0]]) : '';
            }
            tr.appendChild(td);
        });
        body.appendChild(tr);
    });
}

function ingestFiles(fileList) {
    const list = Array.from(fileList || []);
    list.forEach(function (f) {
        if (
            !pendingFiles.some(function (p) {
                return p.name === f.name && p.size === f.size;
            })
        ) {
            pendingFiles.push(f);
        }
    });
    renderFileList();
}

function isSubjectPdfFile(file) {
    const n = String(file.name || '').toLowerCase();
    return n.endsWith('.pdf') && (/subject/.test(n) || /^fach/.test(n) || /fächer/.test(n));
}

function isClassPdfFile(file) {
    const n = String(file.name || '').toLowerCase();
    if (!n.endsWith('.pdf') || isSubjectPdfFile(file)) return false;
    return /class|klassen|klasse/.test(n) || !/subject|fach|fächer/.test(n);
}

async function parseSubjectPdfs(files, progress, stepIds) {
    const wu = window.ms365WebuntisExportImport;
    const pdfApi = window.ms365PdfText;
    const blocks = [];
    let idx = 0;
    for (const file of files) {
        if (!isSubjectPdfFile(file)) continue;
        const stepId = stepIds && stepIds[idx];
        if (progress && stepId) {
            progress.active(stepId, 'Fächer-PDF: ' + file.name);
        }
        if (!wu || !pdfApi) throw new Error('PDF-Modul fehlt – Seite neu laden.');
        const content = await pdfApi.extractPdfFile(file);
        const result =
            typeof wu.importSubjectsFromWebuntisPdf === 'function'
                ? wu.importSubjectsFromWebuntisPdf({ text: content.text, words: content.words })
                : { subjects: [] };
        blocks.push({ source: file.name, subjects: result.subjects || [] });
        if (progress && stepId) progress.complete(stepId);
        idx += 1;
    }
    return blocks;
}

async function readClassPdfImports(files, progress, stepIds) {
    const pdfApi = window.ms365PdfText;
    const blocks = [];
    let idx = 0;
    for (const file of files) {
        if (!isClassPdfFile(file)) continue;
        const stepId = stepIds && stepIds[idx];
        if (progress && stepId) {
            progress.active(stepId, 'Klassen-PDF: ' + file.name + ' (KV & Abschlussjahr)');
        }
        if (!pdfApi) throw new Error('PDF-Modul fehlt – Seite neu laden.');
        const content = await pdfApi.extractPdfFile(file);
        blocks.push({
            source: file.name,
            text: content.text || '',
            words: content.words && content.words.length ? content.words : null
        });
        if (progress && stepId) progress.complete(stepId);
        idx += 1;
    }
    return blocks;
}

async function runParseAndReview() {
    if (!pendingFiles.length) {
        toast('Bitte mindestens eine WebUntis-Datei wählen.');
        return;
    }
    const io = window.ms365StammdatenFileIO;
    if (!io || typeof io.readFilesToAllSheets !== 'function') {
        toast('Import-Modul nicht geladen – Seite neu laden.');
        return;
    }
    const spreadsheetFiles = pendingFiles.filter(function (f) {
        return !String(f.name || '').toLowerCase().endsWith('.pdf');
    });
    const pdfFiles = pendingFiles.filter(function (f) {
        return String(f.name || '').toLowerCase().endsWith('.pdf');
    });
    const subjectPdfFiles = pdfFiles.filter(isSubjectPdfFile);
    const classPdfFiles = pdfFiles.filter(isClassPdfFile);

    const progressSteps = [];
    const subjectStepIds = [];
    const classStepIds = [];
    if (spreadsheetFiles.length) {
        progressSteps.push({
            id: 'csv',
            label: 'CSV/Excel einlesen (' + spreadsheetFiles.length + ' Datei(en))',
            weight: Math.max(1, spreadsheetFiles.length)
        });
    }
    subjectPdfFiles.forEach(function (f, i) {
        const id = 'spdf-' + i;
        subjectStepIds.push(id);
        progressSteps.push({ id: id, label: 'Fächer-PDF: ' + f.name, weight: 3 });
    });
    classPdfFiles.forEach(function (f, i) {
        const id = 'cpdf-' + i;
        classStepIds.push(id);
        progressSteps.push({ id: id, label: 'Klassen-PDF: ' + f.name, weight: 4 });
    });
    progressSteps.push({ id: 'merge', label: 'Daten zusammenführen (Lehrer, Schüler, KV, Abschlussjahre)', weight: 3 });

    const progress = createImportParseProgress();
    progress.begin(progressSteps);

    try {
        let sheets = [];
        if (spreadsheetFiles.length) {
            progress.active('csv', spreadsheetFiles.map(function (f) {
                return f.name;
            }).join(', '));
            sheets = await io.readFilesToAllSheets(spreadsheetFiles);
            progress.complete('csv');
        }
        const subjectPdfImports = await parseSubjectPdfs(subjectPdfFiles, progress, subjectStepIds);
        const classPdfImports = await readClassPdfImports(classPdfFiles, progress, classStepIds);
        progress.active('merge', 'Lehrer, Schüler, Eltern, Fächer und Klassen werden verknüpft …');
        const emailOpts = getEmailOptsFromSettings();
        previewState = buildWizardPreview(
            {
                sheets: sheets,
                classPdfImports: classPdfImports,
                kvTeachers: getTeachersForClassEnrichment(),
                schoolYearEnd: getSchoolYearEndForClassPdf(),
                subjectPdfImports: subjectPdfImports,
                emailOpts: emailOpts
            },
            {
                webuntis: window.ms365WebuntisExportImport,
                schoolSis: window.ms365SchoolSisImport
            }
        );
        progress.complete('merge');

        const warnEl = document.getElementById('wuStammdatenWizardWarn');
        const warnings = (previewState.warnings || []).concat(
            (previewState.fileSummary || [])
                .filter(function (f) {
                    return f.kind === 'unknown';
                })
                .map(function (f) {
                    return 'Unbekanntes Format: ' + f.name;
                })
        );
        if (warnEl) {
            if (warnings.length) {
                warnEl.hidden = false;
                warnEl.textContent = warnings.join(' · ');
            } else {
                warnEl.hidden = true;
                warnEl.textContent = '';
            }
        }

        const status = document.getElementById('wuStammdatenWizardStatus');
        if (status && previewState.meta) {
            status.textContent =
                'Vorschau: ' +
                previewState.meta.teacherCount +
                ' Lehrer, ' +
                previewState.meta.studentCount +
                ' Schüler, ' +
                previewState.meta.subjectCount +
                ' Fächer, ' +
                previewState.meta.classCount +
                ' Klassen' +
                (previewState.meta.guardianRows
                    ? ' · Eltern: ' +
                      (previewState.meta.guardiansLinked || 0) +
                      ' zugeordnet, ' +
                      (previewState.meta.unmatchedGuardians || 0) +
                      ' offen'
                    : '') +
                (previewState.meta.privateStudentEmailsIgnored
                    ? ' · ' + previewState.meta.privateStudentEmailsIgnored + ' private Schüler-Mail(s) ignoriert'
                    : '');
        }

        if (previewState.meta && previewState.meta.subjectCount > previewState.meta.teacherCount) {
            activeTile = 'subjects';
            subjectViewMode = previewState.meta.subjectCount > 25 ? 'import' : 'all';
        } else {
            activeTile = 'teachers';
            subjectViewMode = 'all';
        }
        setPhase('review');
        renderTiles();
        const selectAllEl = document.getElementById('wuStammdatenSelectAll');
        if (selectAllEl && previewState.meta && previewState.meta.subjectCount > 25) {
            selectAllEl.checked = true;
        }
        renderReviewTable();
        toast('Auswertung abgeschlossen – bitte Daten prüfen.', { kind: 'success', durationMs: 3500 });
    } catch (err) {
        toast('Auswertung fehlgeschlagen: ' + (err && err.message ? err.message : String(err)), { kind: 'error' });
    } finally {
        progress.end();
    }
}

function applyImport() {
    if (!previewState) return;
    const payload = compileWizardApplyPayload(previewState, {
        webuntis: window.ms365WebuntisExportImport,
        schoolSis: window.ms365SchoolSisImport
    });
    const returnUrl = resolveReturnUrl(fromParam());
    stashWebuntisImportPayload(payload, returnUrl);
    toast(
        'Weiterleitung – dort werden die Daten mit bestehenden Stammdaten abgeglichen (kein Doppelimport). Bitte noch „Speichern“ klicken (' +
            payload.counts.teachers +
            ' Lehrer, ' +
            payload.counts.students +
            ' Schüler, …).'
    );
    window.location.href = returnUrl;
}

function wireSubjectImportUx() {
    const panel = document.getElementById('wuSubjectImportPanel');
    if (!panel || panel.dataset.wired) return;
    panel.dataset.wired = '1';
    panel.querySelectorAll('[data-subject-preset]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            applySubjectPreset(btn.getAttribute('data-subject-preset') || 'recommended');
        });
    });
    panel.querySelectorAll('[data-subject-view]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            subjectViewMode = btn.getAttribute('data-subject-view') === 'all' ? 'all' : 'import';
            panel.querySelectorAll('[data-subject-view]').forEach(function (b) {
                b.classList.toggle('is-active', b === btn);
            });
            renderReviewTable();
        });
    });
}

function init() {
    const back = document.getElementById('wuImportBackLink');
    if (back) back.href = resolveReturnUrl(fromParam());

    document.getElementById('wuStammdatenBtnBack').addEventListener('click', function () {
        setPhase('upload');
        renderTiles();
    });
    document.getElementById('wuStammdatenBtnNext').addEventListener('click', function () {
        runParseAndReview();
    });
    document.getElementById('wuStammdatenBtnApply').addEventListener('click', applyImport);

    const drop = document.getElementById('wuStammdatenDrop');
    const input = document.getElementById('wuStammdatenFileInput');
    drop.addEventListener('click', function () {
        input.click();
    });
    drop.addEventListener('dragover', function (e) {
        e.preventDefault();
        drop.classList.add('dragover');
    });
    drop.addEventListener('dragleave', function () {
        drop.classList.remove('dragover');
    });
    drop.addEventListener('drop', function (e) {
        e.preventDefault();
        drop.classList.remove('dragover');
        if (e.dataTransfer && e.dataTransfer.files) ingestFiles(e.dataTransfer.files);
    });
    input.addEventListener('change', function (e) {
        if (e.target.files) ingestFiles(e.target.files);
        input.value = '';
    });

    wireAreaTabs();
    wireSubjectImportUx();
    document.getElementById('wuStammdatenFilter').addEventListener('input', function () {
        renderReviewTable();
    });
    document.getElementById('wuStammdatenSelectAll').addEventListener('change', function () {
        const checked = !!this.checked;
        visibleRowsForTile().forEach(function (r) {
            r.selected = checked;
        });
        if (activeTile === 'subjects') renderSubjectImportHero();
        renderReviewTable();
        renderTiles();
    });

    renderFileList();
    renderTiles();
    setPhase('upload');
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
