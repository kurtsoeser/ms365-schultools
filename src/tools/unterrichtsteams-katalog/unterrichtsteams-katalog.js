/**
 * Unterrichtsteams-Katalog – Übersicht & Bearbeitung (Unterrichtsbelegung).
 */
import {
    filterUnterrichtsteamRows,
    uniqueFieldValues,
    sortUnterrichtsteamRows,
    rowStableKey,
    updateRowByKey,
    removeRowByKey,
    snapshotFromRows,
    mergeWizardStateIntoSnapshot,
    loadWizardMergeFromBrowser,
    countLinked
} from './unterrichtsteams-katalog-logic.js';

/** @type {object[]} */
let allRows = [];
let dirty = false;
let metaYearPrefix = '';

const filters = {
    q: '',
    klasse: '',
    fach: '',
    lehrerCode: '',
    linkedOnly: false
};

function $(id) {
    return document.getElementById(id);
}

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else window.alert(msg);
}

function escapeHtml(text) {
    const d = document.createElement('div');
    d.textContent = text;
    return d.innerHTML;
}

function loadFromApp() {
    const api = window.ms365AppDataV2;
    if (!api || typeof api.getUnterrichtsbelegung !== 'function') {
        allRows = [];
        metaYearPrefix = '';
        return;
    }
    const snap = api.getUnterrichtsbelegung();
    metaYearPrefix = snap && snap.yearPrefix ? String(snap.yearPrefix) : '';
    allRows = sortUnterrichtsteamRows(snap && Array.isArray(snap.rows) ? snap.rows : []);
    dirty = false;
}

function saveToApp(pushM365Names) {
    const api = window.ms365AppDataV2;
    if (!api || typeof api.setUnterrichtsbelegung !== 'function') {
        toast('App-Daten nicht verfügbar.');
        return;
    }
    const snap = snapshotFromRows(allRows, {
        yearPrefix: metaYearPrefix || $('utkYearPrefix')?.value || '',
        source: 'unterrichtsteams-katalog'
    });
    if (!snap) {
        toast('Nichts zu speichern.');
        return;
    }
    api.setUnterrichtsbelegung(snap);
    metaYearPrefix = snap.yearPrefix || metaYearPrefix;
    dirty = false;
    toast('Unterrichtsbelegung gespeichert (' + snap.rows.length + ' Teams).');
    if (pushM365Names) {
        pushDisplayNamesToM365().catch((e) => toast('M365: ' + (e.message || e)));
    }
    render();
}

async function pushDisplayNamesToM365() {
    const linked = allRows.filter((r) => r.graphGroupId && r.teamName);
    if (!linked.length) {
        toast('Keine verknüpften Teams mit Anzeigenamen.');
        return;
    }
    const { getGraphToken } = await import('../kursteams/kursteam-graph.js');
    const token = await getGraphToken();
    let ok = 0;
    let fail = 0;
    for (let i = 0; i < linked.length; i++) {
        const r = linked[i];
        try {
            const res = await fetch('https://graph.microsoft.com/v1.0/groups/' + encodeURIComponent(r.graphGroupId), {
                method: 'PATCH',
                headers: {
                    Authorization: 'Bearer ' + token,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ displayName: r.teamName })
            });
            if (!res.ok) throw new Error(await res.text());
            ok++;
        } catch {
            fail++;
        }
    }
    toast('M365 Anzeigename: ' + ok + ' aktualisiert' + (fail ? ', ' + fail + ' fehlgeschlagen' : '') + '.');
}

function readFiltersFromUi() {
    filters.q = $('utkSearch')?.value || '';
    filters.klasse = $('utkFilterKlasse')?.value || '';
    filters.fach = $('utkFilterFach')?.value || '';
    filters.lehrerCode = $('utkFilterLehrer')?.value || '';
    filters.linkedOnly = !!$('utkFilterLinked')?.checked;
}

function fillFilterSelects() {
    const kEl = $('utkFilterKlasse');
    const fEl = $('utkFilterFach');
    const lEl = $('utkFilterLehrer');
    if (!kEl || !fEl || !lEl) return;

    const klasse = uniqueFieldValues(allRows, 'klasse');
    const fach = uniqueFieldValues(allRows, 'fach');
    const lehrer = uniqueFieldValues(allRows, 'lehrerCode');

    function fill(sel, values, current) {
        const opts = ['<option value="">Alle</option>']
            .concat(values.map((v) => '<option value="' + escapeHtml(v) + '">' + escapeHtml(v) + '</option>'))
            .join('');
        sel.innerHTML = opts;
        sel.value = current && values.indexOf(current) !== -1 ? current : '';
    }
    fill(kEl, klasse, filters.klasse);
    fill(fEl, fach, filters.fach);
    fill(lEl, lehrer, filters.lehrerCode);
}

function onFieldChange(key, field, value) {
    const patch = {};
    patch[field] = value;
    const res = updateRowByKey(allRows, key, patch);
    if (res.ok) {
        allRows = res.rows;
        dirty = true;
        const hint = $('utkDirtyHint');
        if (hint) hint.hidden = false;
    }
}

function renderTable(filtered) {
    const host = $('utkTableHost');
    if (!host) return;
    if (!filtered.length) {
        host.innerHTML =
            '<p class="muted" style="margin:12px;">Keine Einträge für die aktuelle Filterung.</p>';
        return;
    }

    const rows = filtered
        .map((r) => {
            const key = rowStableKey(r);
            const linked = !!String(r.graphGroupId || '').trim();
            return (
                '<tr data-utk-key="' +
                escapeHtml(key) +
                '">' +
                '<td><input type="text" data-field="klasse" value="' +
                escapeHtml(r.klasse || '') +
                '" aria-label="Klasse"></td>' +
                '<td><input type="text" data-field="fach" value="' +
                escapeHtml(r.fach || '') +
                '" aria-label="Fach"></td>' +
                '<td><input type="text" data-field="lehrerCode" value="' +
                escapeHtml(r.lehrerCode || '') +
                '" aria-label="Lehrkraft-Kürzel"></td>' +
                '<td><input type="text" class="utk-input--wide" data-field="teamName" value="' +
                escapeHtml(r.teamName || '') +
                '" aria-label="Anzeigename"></td>' +
                '<td><input type="text" data-field="gruppenmail" value="' +
                escapeHtml(r.gruppenmail || '') +
                '" aria-label="Mail-Nickname"></td>' +
                '<td><input type="text" data-field="lehrerEmail" value="' +
                escapeHtml(r.lehrerEmail || '') +
                '" aria-label="Lehrer E-Mail"></td>' +
                '<td>' +
                (linked
                    ? '<span class="utk-badge" title="' +
                      escapeHtml(r.graphGroupId) +
                      '">M365</span>'
                    : '<span class="utk-badge utk-badge--muted">—</span>') +
                '</td>' +
                '<td><button type="button" class="btn btn-small btn-danger utk-del" title="Zeile entfernen"><i class="bi bi-trash"></i></button></td>' +
                '</tr>'
            );
        })
        .join('');

    host.innerHTML =
        '<div class="utk-table-wrap"><table class="utk-table"><thead><tr>' +
        '<th>Klasse</th><th>Fach</th><th>LK</th><th>Anzeigename</th><th>Mail-Nickname</th><th>Lehrer E-Mail</th><th>M365</th><th></th>' +
        '</tr></thead><tbody>' +
        rows +
        '</tbody></table></div>';

    host.querySelectorAll('tbody tr').forEach((tr) => {
        const key = tr.getAttribute('data-utk-key');
        tr.querySelectorAll('input[data-field]').forEach((inp) => {
            inp.addEventListener('change', () => {
                onFieldChange(key, inp.getAttribute('data-field'), inp.value);
            });
        });
        const del = tr.querySelector('.utk-del');
        if (del) {
            del.addEventListener('click', () => {
                if (!window.confirm('Diese Zeile aus der Unterrichtsbelegung entfernen?')) return;
                const res = removeRowByKey(allRows, key);
                allRows = res.rows;
                dirty = true;
                render();
            });
        }
    });
}

function render() {
    readFiltersFromUi();
    fillFilterSelects();
    const filtered = filterUnterrichtsteamRows(allRows, filters);
    const stats = $('utkStats');
    if (stats) {
        stats.innerHTML =
            '<span><strong>' +
            allRows.length +
            '</strong> Teams gesamt</span>' +
            '<span><strong>' +
            filtered.length +
            '</strong> angezeigt</span>' +
            '<span><strong>' +
            countLinked(allRows) +
            '</strong> mit M365-ID</span>' +
            (metaYearPrefix ? '<span>Schuljahr-Präfix: <strong>' + escapeHtml(metaYearPrefix) + '</strong></span>' : '');
    }
    const empty = $('utkEmpty');
    if (empty) empty.hidden = allRows.length > 0;
    const hint = $('utkDirtyHint');
    if (hint) hint.hidden = !dirty;
    const yp = $('utkYearPrefix');
    if (yp && metaYearPrefix && !yp.value) yp.value = metaYearPrefix;
    renderTable(filtered);
}

function pullFromWizard() {
    const state = loadWizardMergeFromBrowser();
    const api = window.ms365AppDataV2;
    const existing = api && typeof api.getUnterrichtsbelegung === 'function' ? api.getUnterrichtsbelegung() : null;
    const { snapshot, imported, source } = mergeWizardStateIntoSnapshot(existing, state);
    if (!imported) {
        toast('Kein Kursteam-/WebUntis-Stand im Browser – zuerst in Kursteams importieren.');
        return;
    }
    if (snapshot && snapshot.rows) {
        allRows = sortUnterrichtsteamRows(snapshot.rows);
        metaYearPrefix = snapshot.yearPrefix || metaYearPrefix;
        dirty = true;
        toast(imported + ' Einträge aus Assistent übernommen (' + source + '). Bitte speichern.');
        render();
    }
}

function addEmptyRow() {
    allRows.push({
        klasse: '',
        fach: '',
        lehrerCode: '',
        lehrerEmail: '',
        gruppe: '',
        teamName: '',
        gruppenmail: '',
        graphGroupId: ''
    });
    dirty = true;
    render();
}

function wireUi() {
    $('utkRefresh')?.addEventListener('click', () => {
        if (dirty && !window.confirm('Ungespeicherte Änderungen verwerfen und neu laden?')) return;
        loadFromApp();
        render();
    });
    $('utkSave')?.addEventListener('click', () => {
        const push = !!$('utkPushM365')?.checked;
        metaYearPrefix = $('utkYearPrefix')?.value || metaYearPrefix;
        saveToApp(push);
    });
    $('utkPullWizard')?.addEventListener('click', () => pullFromWizard());
    $('utkAddRow')?.addEventListener('click', () => addEmptyRow());
    ['utkSearch', 'utkFilterKlasse', 'utkFilterFach', 'utkFilterLehrer', 'utkFilterLinked'].forEach((id) => {
        $(id)?.addEventListener('input', () => render());
        $(id)?.addEventListener('change', () => render());
    });
    $('utkResetFilters')?.addEventListener('click', () => {
        filters.q = '';
        filters.klasse = '';
        filters.fach = '';
        filters.lehrerCode = '';
        filters.linkedOnly = false;
        if ($('utkSearch')) $('utkSearch').value = '';
        if ($('utkFilterLinked')) $('utkFilterLinked').checked = false;
        render();
    });
}

function init() {
    wireUi();
    loadFromApp();
    render();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}

window.ms365UnterrichtsteamsKatalog = {
    filterUnterrichtsteamRows,
    reload: () => {
        loadFromApp();
        render();
    }
};
