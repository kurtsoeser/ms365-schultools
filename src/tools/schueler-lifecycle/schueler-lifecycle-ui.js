/**
 * Schüler-Lifecycle UI (Analyse 04 A1) – Diff-first, kein Apply ohne Preview.
 */
import {
    diffStudents,
    hasMembershipWork,
    previewMemberships,
    summarizePreview
} from '../../shared/student-class-lifecycle.js';

const SNAPSHOT_KEY = 'ms365-student-lifecycle-prev-v1';

function $(id) {
    return document.getElementById(id);
}

function escapeHtml(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function loadStudentsFromTenant() {
    if (typeof window.ms365TenantSettingsLoad !== 'function') return [];
    const s = window.ms365TenantSettingsLoad();
    return Array.isArray(s && s.students) ? s.students : [];
}

function loadClassTeamsHint() {
    const out = [];
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad();
            const classes = Array.isArray(s && s.classes) ? s.classes : [];
            classes.forEach(function (c) {
                const code = String((c && (c.code || c.name)) || '').trim();
                if (!code) return;
                out.push({
                    classCode: code,
                    graphGroupId: String((c && (c.graphGroupId || c.groupId)) || '').trim(),
                    displayName: String((c && c.name) || code).trim()
                });
            });
        }
    } catch {
        /* ignore */
    }
    return out;
}

function loadPrevSnapshot() {
    try {
        const raw = localStorage.getItem(SNAPSHOT_KEY);
        if (!raw) return null;
        const o = JSON.parse(raw);
        if (o && Array.isArray(o.students)) return o;
    } catch {
        /* ignore */
    }
    return null;
}

function savePrevSnapshot(students, label) {
    try {
        localStorage.setItem(
            SNAPSHOT_KEY,
            JSON.stringify({
                savedAt: new Date().toISOString(),
                label: label || '',
                students: Array.isArray(students) ? students : []
            })
        );
    } catch {
        /* ignore */
    }
}

function parseCsvStudents(text) {
    const lines = String(text || '')
        .replace(/^\uFEFF/, '')
        .split(/\r?\n/)
        .map(function (l) {
            return l.trim();
        })
        .filter(Boolean);
    if (!lines.length) return [];
    const sep = lines[0].indexOf(';') >= 0 ? ';' : ',';
    const headers = lines[0].split(sep).map(function (h) {
        return h.trim().toLowerCase();
    });
    const iName = headers.findIndex(function (h) {
        return /name|nachname|schüler|schueler/.test(h);
    });
    const iEmail = headers.findIndex(function (h) {
        return /mail|email|upn/.test(h);
    });
    const iKlasse = headers.findIndex(function (h) {
        return /klasse|class|jahrgang/.test(h);
    });
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
        const cols = lines[i].split(sep);
        const email = String(cols[iEmail >= 0 ? iEmail : 1] || '').trim();
        const name = String(cols[iName >= 0 ? iName : 0] || '').trim();
        const klasse = String(cols[iKlasse >= 0 ? iKlasse : 2] || '').trim();
        if (!email && !name) continue;
        rows.push({ name: name, email: email.toLowerCase(), klasse: klasse });
    }
    return rows;
}

function renderList(el, items, formatter) {
    if (!el) return;
    el.replaceChildren();
    if (!items.length) {
        const p = document.createElement('p');
        p.className = 'muted';
        p.textContent = 'Keine Einträge.';
        el.appendChild(p);
        return;
    }
    const ul = document.createElement('ul');
    ul.className = 'slc-list';
    items.slice(0, 80).forEach(function (it) {
        const li = document.createElement('li');
        li.innerHTML = formatter(it);
        ul.appendChild(li);
    });
    if (items.length > 80) {
        const li = document.createElement('li');
        li.className = 'muted';
        li.textContent = '… und ' + (items.length - 80) + ' weitere';
        ul.appendChild(li);
    }
    el.appendChild(ul);
}

function setStatus(msg, isError) {
    const el = $('slcStatus');
    if (!el) return;
    el.textContent = msg || '';
    el.className = isError ? 'slc-status slc-status--err' : 'slc-status';
}

let lastDiff = null;
let lastPreview = null;

function runPreview() {
    const next = loadStudentsFromTenant();
    let prev = [];
    const mode = ($('slcPrevMode') && $('slcPrevMode').value) || 'snapshot';
    if (mode === 'snapshot') {
        const snap = loadPrevSnapshot();
        prev = snap && Array.isArray(snap.students) ? snap.students : [];
        if (!prev.length) {
            setStatus('Kein gespeicherter Vorher-Stand. Bitte zuerst „Aktuellen Stand als Vorher speichern“ oder CSV laden.', true);
            return;
        }
    } else if (mode === 'csv') {
        const ta = $('slcCsv');
        prev = parseCsvStudents(ta && ta.value);
        if (!prev.length) {
            setStatus('CSV leer oder nicht lesbar (Spalten: Name;E-Mail;Klasse).', true);
            return;
        }
    }

    lastDiff = diffStudents(prev, next);
    const teams = loadClassTeamsHint();
    const sammel = String(($('slcSammelGroupId') && $('slcSammelGroupId').value) || '').trim();
    lastPreview = previewMemberships(lastDiff, teams, sammel);

    const sum = summarizePreview(lastPreview);
    $('slcDiffSummary').innerHTML =
        '<strong>' +
        lastDiff.added.length +
        '</strong> Zugang · <strong>' +
        lastDiff.removed.length +
        '</strong> Abgang · <strong>' +
        lastDiff.classChanged.length +
        '</strong> Klassenwechsel · Graph-Vorschau: <strong>' +
        sum.join +
        '</strong> Join / <strong>' +
        sum.leave +
        '</strong> Leave in <strong>' +
        sum.groupCount +
        '</strong> Gruppen';

    renderList($('slcAdded'), lastDiff.added, function (s) {
        return escapeHtml(s.name || s.email) + ' → <code>' + escapeHtml(s.klasse) + '</code>';
    });
    renderList($('slcRemoved'), lastDiff.removed, function (s) {
        return escapeHtml(s.name || s.email) + ' (<code>' + escapeHtml(s.klasse) + '</code>)';
    });
    renderList($('slcChanged'), lastDiff.classChanged, function (row) {
        const s = row.student || {};
        return (
            escapeHtml(s.name || s.email) +
            ': <code>' +
            escapeHtml(row.fromClass) +
            '</code> → <code>' +
            escapeHtml(row.toClass) +
            '</code>'
        );
    });
    renderList($('slcPreviewGroups'), lastPreview.groups || [], function (g) {
        return (
            '<strong>' +
            escapeHtml(g.label || g.groupId) +
            '</strong> · +' +
            (g.join || []).length +
            ' / −' +
            (g.leave || []).length +
            (g.groupId ? ' <span class="muted">(' + escapeHtml(g.groupId) + ')</span>' : ' <span class="muted">(keine Group-ID – nur Info)</span>')
        );
    });

    const panel = $('slcPreviewPanel');
    if (panel) panel.hidden = false;

    const allowLeave = $('slcAllowLeave');
    const applyBtn = $('slcApply');
    const dryRun = !$('slcDryRun') || $('slcDryRun').checked;
    const work = hasMembershipWork(lastPreview);
    const leaveBlocked = sum.leave > 0 && allowLeave && !allowLeave.checked;

    if (applyBtn) {
        applyBtn.disabled = !work || leaveBlocked || dryRun;
        applyBtn.title = dryRun
            ? 'Dry-Run aktiv – Apply gesperrt (nur Vorschau).'
            : leaveBlocked
              ? 'Entfernen erst nach Checkbox „Leave erlauben“.'
              : work
                ? 'Mitgliedschaften in Microsoft 365 anpassen (Graph).'
                : 'Keine Graph-Änderungen nötig.';
    }

    if (typeof window.ms365TruncationUi === 'object' && window.ms365TruncationUi.hideTruncationBanner) {
        window.ms365TruncationUi.hideTruncationBanner('slcBannerHost');
    }

    setStatus(
        dryRun
            ? 'Vorschau bereit (Dry-Run). Zum echten Apply Dry-Run abwählen und bei Leaves die Checkbox setzen.'
            : leaveBlocked
              ? 'Vorschau bereit – Leave ist gesperrt, bis Sie es erlauben.'
              : 'Vorschau bereit.'
    );
}

function wire() {
    const saveBtn = $('slcSaveSnapshot');
    if (saveBtn) {
        saveBtn.addEventListener('click', function () {
            const students = loadStudentsFromTenant();
            savePrevSnapshot(students, 'Stammdaten');
            setStatus('Vorher-Stand gespeichert (' + students.length + ' Schüler).');
            const meta = $('slcSnapshotMeta');
            if (meta) meta.textContent = 'Gespeichert: ' + new Date().toLocaleString('de-AT') + ' · ' + students.length + ' Personen';
        });
    }

    const previewBtn = $('slcPreview');
    if (previewBtn) previewBtn.addEventListener('click', runPreview);

    const allowLeave = $('slcAllowLeave');
    if (allowLeave) allowLeave.addEventListener('change', function () {
        if (lastDiff) runPreview();
    });
    const dryRun = $('slcDryRun');
    if (dryRun) dryRun.addEventListener('change', function () {
        if (lastDiff) runPreview();
    });

    const applyBtn = $('slcApply');
    if (applyBtn) {
        applyBtn.addEventListener('click', function () {
            setStatus(
                'Graph-Apply ist in dieser Version bewusst an den bestehenden Sync (Schüler/Lehrkräfte, Jahrgangsgruppen) gekoppelt. Nutzen Sie die Deep-Links unten mit der Vorschau als Checkliste – kein stilles Massen-Remove.',
                false
            );
        });
    }

    const snap = loadPrevSnapshot();
    const meta = $('slcSnapshotMeta');
    if (meta && snap) {
        meta.textContent =
            'Vorher: ' +
            (snap.savedAt ? new Date(snap.savedAt).toLocaleString('de-AT') : '–') +
            ' · ' +
            (snap.students || []).length +
            ' Personen';
    }

    const countEl = $('slcCurrentCount');
    if (countEl) {
        const n = loadStudentsFromTenant().length;
        countEl.textContent = n ? n + ' Schüler in Stammdaten' : 'Keine Schüler in Stammdaten';
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wire);
} else {
    wire();
}

export { runPreview, parseCsvStudents, SNAPSHOT_KEY };
