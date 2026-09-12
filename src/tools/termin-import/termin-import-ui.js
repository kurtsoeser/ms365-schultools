/**
 * CSV-Import-UI für SharePoint-Schultermine.
 */
import { parseTermineCsvText, terminToListFields } from './termin-import-logic.js';

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

function targetLabel(t) {
    if (t === 'teachers') return 'Lehrer';
    if (t === 'both') return 'Beide';
    return 'Schule';
}

/** @type {{ rows: ReturnType<typeof parseTermineCsvText>['rows'], summary: object } | null} */
let lastParsed = null;

function renderPreview() {
    const body = $('sptImportBody');
    const meta = $('sptImportMeta');
    if (!body) return;
    body.replaceChildren();
    if (!lastParsed || !lastParsed.rows.length) {
        body.innerHTML = '<tr><td colspan="5" class="muted">Noch keine Datei geladen.</td></tr>';
        if (meta) meta.textContent = '';
        return;
    }
    const s = lastParsed.summary;
    if (meta) {
        meta.textContent =
            s.total +
            ' Zeilen · gültig ' +
            s.ok +
            ' · fehlerhaft ' +
            s.bad +
            ' · Schule ' +
            s.school +
            ' · Lehrer ' +
            s.teachers;
    }
    lastParsed.rows.slice(0, 80).forEach(function (r) {
        const tr = document.createElement('tr');
        tr.innerHTML =
            '<td>' +
            escapeHtml(r.title || '—') +
            '</td><td>' +
            escapeHtml(r.start || '') +
            '</td><td>' +
            escapeHtml(r.end || '') +
            '</td><td>' +
            escapeHtml(targetLabel(r.target)) +
            '</td><td>' +
            (r.ok ? 'OK' : escapeHtml((r.issues || []).join(', '))) +
            '</td>';
        if (!r.ok) tr.style.opacity = '0.75';
        body.appendChild(tr);
    });
    if (lastParsed.rows.length > 80) {
        const tr = document.createElement('tr');
        tr.innerHTML =
            '<td colspan="5" class="muted">… weitere ' +
            (lastParsed.rows.length - 80) +
            ' Zeilen in der Datei</td>';
        body.appendChild(tr);
    }
}

function filterOkRows() {
    const prefer = String(($('sptImportTargetFilter') && $('sptImportTargetFilter').value) || 'all');
    if (!lastParsed) return [];
    return lastParsed.rows.filter(function (r) {
        if (!r.ok) return false;
        if (prefer === 'all') return true;
        if (prefer === 'school') return r.target === 'school' || r.target === 'both';
        if (prefer === 'teachers') return r.target === 'teachers' || r.target === 'both';
        return true;
    });
}

async function runImport() {
    const api = window.ms365SpoSchultermine;
    if (!api || typeof api.importTermine !== 'function') {
        throw new Error('Schultermine-Modul nicht geladen.');
    }
    const rows = filterOkRows();
    if (!rows.length) throw new Error('Keine gültigen Termine zum Import (Filter prüfen).');
    const webUrl = String(($('sptSiteUrl') && $('sptSiteUrl').value) || '').trim();
    const listTitle = String(($('sptListName') && $('sptListName').value) || '').trim() || 'Schultermine';
    if (!webUrl) throw new Error('SharePoint-Website fehlt.');
    if (
        !window.confirm(
            rows.length +
                ' Termin(e) in die Liste „' +
                listTitle +
                '“ schreiben?\n(Zeilen mit Ziel Lehrer erhalten SyncStatus=pending.)'
        )
    ) {
        return;
    }
    const fieldsList = rows.map(terminToListFields);
    await api.importTermine(webUrl, listTitle, fieldsList);
    toast(rows.length + ' Termine geschrieben.');
}

function onFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function () {
        try {
            lastParsed = parseTermineCsvText(String(reader.result || ''));
            renderPreview();
            toast('CSV gelesen: ' + lastParsed.summary.ok + ' gültig.');
        } catch (e) {
            lastParsed = null;
            renderPreview();
            toast('CSV-Fehler: ' + (e && e.message ? e.message : e));
        }
    };
    reader.readAsText(file, 'UTF-8');
}

export function mountTerminImportUi() {
    const file = $('sptImportFile');
    const btn = $('sptBtnImport');
    if (file && file.dataset.bound !== '1') {
        file.dataset.bound = '1';
        file.addEventListener('change', function () {
            const f = file.files && file.files[0];
            onFile(f || null);
            file.value = '';
        });
    }
    if (btn && btn.dataset.bound !== '1') {
        btn.dataset.bound = '1';
        btn.addEventListener('click', function () {
            runImport().catch(function (e) {
                toast(e.message || String(e));
            });
        });
    }
    renderPreview();
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', mountTerminImportUi);
    } else {
        mountTerminImportUi();
    }
}
