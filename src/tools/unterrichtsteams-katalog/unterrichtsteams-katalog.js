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
    loadWizardMergeFromBrowser,
    hydrateUnterrichtsbelegungFromKursteamState,
    findUnterrichtsbelegungInOtherYears,
    countLinked
} from './unterrichtsteams-katalog-logic.js';
import {
    autoLinkRows,
    buildAbgleichForRows,
    entraGroupUrl,
    fetchM365KursteamRows,
    getCachedM365Rows,
    linkPatchForSingleRow
} from './unterrichtsteams-katalog-m365.js';
import {
    buildTeacherLookupFromList,
    enrichBelegungRow,
    enrichBelegungRowsInPlace
} from '../../shared/unterrichtsbelegung-teacher-enrich-logic.js';

/** @type {object[]} */
let allRows = [];
let dirty = false;
let metaYearPrefix = '';
/** @type {ReturnType<typeof buildAbgleichForRows>|null} */
let lastAbgleichReport = null;

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

function waitForAppDataV2(maxMs) {
    const limit = Number.isFinite(maxMs) ? maxMs : 8000;
    return new Promise((resolve) => {
        if (window.ms365AppDataV2) {
            resolve(window.ms365AppDataV2);
            return;
        }
        const t0 = Date.now();
        const tick = () => {
            if (window.ms365AppDataV2) {
                resolve(window.ms365AppDataV2);
                return;
            }
            if (Date.now() - t0 >= limit) {
                resolve(null);
                return;
            }
            requestAnimationFrame(tick);
        };
        tick();
    });
}

/**
 * App-Belegung leer, aber Kursteam-Zwischenstand im Browser → in App-Daten schreiben.
 * @returns {{ ok: boolean, count?: number, source?: string }}
 */
function tryHydrateFromKursteamBrowserState() {
    if (allRows.length > 0) return { ok: false };
    const api = window.ms365AppDataV2;
    if (!api || typeof api.getUnterrichtsbelegung !== 'function' || typeof api.setUnterrichtsbelegung !== 'function') {
        return { ok: false };
    }
    const state = loadWizardMergeFromBrowser();
    if (!state) return { ok: false };
    const existing = api.getUnterrichtsbelegung();
    const { ok, snapshot, imported, source } = hydrateUnterrichtsbelegungFromKursteamState(existing, state);
    if (!ok || !snapshot) return { ok: false };
    api.setUnterrichtsbelegung(snapshot);
    metaYearPrefix = snapshot.yearPrefix || metaYearPrefix;
    allRows = sortUnterrichtsteamRows(snapshot.rows);
    dirty = false;
    return { ok: true, count: imported, source };
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

function getTeacherLookup() {
    try {
        const load =
            typeof window.ms365TenantSettingsLoad === 'function' ? window.ms365TenantSettingsLoad : null;
        const data = load ? load() : null;
        return buildTeacherLookupFromList(data && Array.isArray(data.teachers) ? data.teachers : []);
    } catch {
        return buildTeacherLookupFromList([]);
    }
}

function enrichAllRowsFromTenant(options) {
    const quiet = !!(options && options.quiet);
    const persist = !!(options && options.persist);
    const lookup = getTeacherLookup();
    const { rows, changed } = enrichBelegungRowsInPlace(allRows, lookup);
    if (!changed) {
        if (!quiet) toast('Keine fehlenden Namen/E-Mails – Stammdaten (Lehrerliste) prüfen oder Kürzel eintragen.');
        return 0;
    }
    allRows = sortUnterrichtsteamRows(rows);
    if (persist) {
        const api = window.ms365AppDataV2;
        if (api && typeof api.setUnterrichtsbelegung === 'function') {
            const snap = snapshotFromRows(allRows, {
                yearPrefix: metaYearPrefix || $('utkYearPrefix')?.value || '',
                source: 'unterrichtsteams-katalog'
            });
            if (snap) {
                api.setUnterrichtsbelegung(snap);
                dirty = false;
            }
        } else {
            dirty = true;
        }
    } else {
        dirty = true;
    }
    if (!quiet) toast(changed + ' Zeile(n): Lehrkraft-Name und/oder E-Mail ergänzt.');
    return changed;
}

function yearPrefixForM365() {
    const fromUi = $('utkYearPrefix') && $('utkYearPrefix').value ? String($('utkYearPrefix').value).trim() : '';
    return fromUi || metaYearPrefix || 'SJ26-27';
}

function renderAbgleichPanel() {
    const host = $('utkAbgleichHost');
    const summary = $('utkAbgleichSummary');
    if (!host) return;
    if (!lastAbgleichReport) {
        host.innerHTML =
            '<p class="muted" style="margin:0;">Nach „Teams aus M365 laden“ sehen Sie hier den Vergleich Katalog ↔ Tenant (wie im Kursteam-Import).</p>';
        if (summary) summary.textContent = '';
        return;
    }
    const c = lastAbgleichReport.counts || {};
    if (summary) {
        summary.textContent =
            (c.planned || 0) +
            ' im Katalog · ' +
            (c.m365 || 0) +
            ' in M365 · ' +
            (c.matched || 0) +
            ' passend · ' +
            (c.missingInM365 || 0) +
            ' fehlen in M365 · ' +
            (c.onlyInM365 || 0) +
            ' nur in M365';
    }
    const miss = (lastAbgleichReport.missingInM365 || []).slice(0, 8);
    const missMore = (lastAbgleichReport.missingInM365 || []).length - miss.length;
    let missHtml = '';
    if (miss.length) {
        missHtml =
            '<p style="margin:8px 0 4px;font-size:0.85em;"><strong>Ohne M365-Team (Auszug):</strong></p><ul class="utk-abgleich-list">' +
            miss
                .map(function (r) {
                    return (
                        '<li>' +
                        escapeHtml([r.klasse, r.fach, r.lehrerCode].filter(Boolean).join(' · ')) +
                        (r.gruppenmail ? ' <code>' + escapeHtml(r.gruppenmail) + '</code>' : '') +
                        '</li>'
                    );
                })
                .join('') +
            (missMore > 0 ? '<li class="muted">… und ' + missMore + ' weitere</li>' : '') +
            '</ul>';
    }
    host.innerHTML =
        '<div class="utk-abgleich-stats">' +
        '<span class="utk-abgleich-pill utk-abgleich-pill--ok">' +
        (c.matched || 0) +
        ' passend</span>' +
        '<span class="utk-abgleich-pill utk-abgleich-pill--warn">' +
        (c.missingInM365 || 0) +
        ' fehlen in M365</span>' +
        '<span class="utk-abgleich-pill">' +
        (c.onlyInM365 || 0) +
        ' nur in M365</span>' +
        '</div>' +
        missHtml;
}

function refreshAbgleichFromRows() {
    if (!getCachedM365Rows().length) {
        lastAbgleichReport = null;
        renderAbgleichPanel();
        return;
    }
    lastAbgleichReport = buildAbgleichForRows(allRows);
    renderAbgleichPanel();
}

async function loadM365Teams() {
    const btn = $('utkM365Load');
    if (btn) btn.disabled = true;
    const status = $('utkM365Status');
    if (status) status.textContent = 'Lade Teams aus Microsoft 365 …';
    try {
        const yp = yearPrefixForM365();
        const n = (await fetchM365KursteamRows(yp)).length;
        refreshAbgleichFromRows();
        toast(n ? n + ' Kursteam(s) aus M365 geladen (Präfix „' + yp + '“).' : 'Keine Kursteams mit diesem Präfix in M365 gefunden.');
    } catch (e) {
        toast('M365: ' + (e && e.message ? e.message : e));
    } finally {
        if (btn) btn.disabled = false;
        if (status) status.textContent = '';
    }
}

function linkAllFromAbgleich() {
    if (!getCachedM365Rows().length) {
        toast('Zuerst „Teams aus M365 laden“ (Anmeldung nötig).');
        return;
    }
    const { rows, linked, report } = autoLinkRows(allRows);
    allRows = sortUnterrichtsteamRows(rows);
    lastAbgleichReport = report;
    if (linked > 0) {
        dirty = true;
        const hint = $('utkDirtyHint');
        if (hint) hint.hidden = false;
    }
    renderAbgleichPanel();
    toast(linked ? linked + ' Zeile(n) mit M365 verknüpft – bitte speichern.' : 'Keine neuen Verknüpfungen (bereits verknüpft oder kein Treffer).');
    render();
}

function linkOneRowByKey(key) {
    if (!getCachedM365Rows().length) {
        toast('Zuerst „Teams aus M365 laden“.');
        return;
    }
    const row = allRows.find((r) => rowStableKey(r) === key);
    if (!row) return;
    const patch = linkPatchForSingleRow(row);
    if (!patch) {
        toast('Kein passendes Team in M365 (Mail-Nickname oder Klasse/Fach/LK prüfen).');
        return;
    }
    const res = updateRowByKey(allRows, key, patch);
    if (res.ok) {
        allRows = sortUnterrichtsteamRows(res.rows);
        dirty = true;
        refreshAbgleichFromRows();
        render();
        toast('Verknüpft – bitte speichern.');
    }
}

function unlinkRowByKey(key) {
    const res = updateRowByKey(allRows, key, { graphGroupId: '' });
    if (res.ok) {
        allRows = res.rows;
        dirty = true;
        refreshAbgleichFromRows();
        render();
    }
}

function shortGraphGroupId(gid) {
    const id = String(gid || '').trim();
    if (!id) return '';
    return id.length > 14 ? id.slice(0, 12) + '…' : id;
}

function m365StatusCellHtml(row) {
    const gid = String(row.graphGroupId || '').trim();
    if (!gid) {
        return '<span class="utk-m365-none" title="Noch nicht mit einem Microsoft-365-Team verknüpft">–</span>';
    }
    const title =
        (row.teamName ? String(row.teamName) : '') +
        (row.gruppenmail ? '\nAlias: ' + String(row.gruppenmail) : '') +
        '\nGroup-ID: ' +
        gid;
    return (
        '<span class="utk-m365-linked" title="' +
        escapeHtml(title) +
        '"><span class="utk-m365-check" aria-hidden="true">✓</span> <code class="utk-m365-id">' +
        escapeHtml(shortGraphGroupId(gid)) +
        '</code></span>'
    );
}

function m365ActionCellHtml(linked, graphGroupId) {
    const entra =
        linked && graphGroupId
            ? '<a class="utk-mini-btn utk-mini-btn--entra utk-m365-entra" href="' +
              escapeHtml(entraGroupUrl(graphGroupId)) +
              '" target="_blank" rel="noopener" title="In Entra öffnen"><i class="bi bi-box-arrow-up-right" aria-hidden="true"></i></a>'
            : '';
    const unlink = linked
        ? '<button type="button" class="utk-mini-btn utk-mini-btn--warn utk-m365-unlink" title="Verknüpfung lösen"><i class="bi bi-x-lg" aria-hidden="true"></i></button>'
        : '';
    return (
        '<td class="utk-action-cell">' +
        '<button type="button" class="utk-mini-btn utk-mini-btn--brand utk-m365-link" title="Team in Microsoft 365 prüfen oder verknüpfen">' +
        '<i class="bi bi-microsoft" aria-hidden="true"></i></button>' +
        entra +
        unlink +
        '<button type="button" class="utk-mini-btn utk-mini-btn--danger utk-del" title="Zeile entfernen">' +
        '<i class="bi bi-trash" aria-hidden="true"></i></button></td>'
    );
}

function onFieldChange(key, field, value) {
    const patch = {};
    patch[field] = value;
    const res = updateRowByKey(allRows, key, patch);
    if (res.ok) {
        allRows = res.rows;
        if (field === 'lehrerCode') {
            const idx = allRows.findIndex((r) => rowStableKey(r) === key);
            if (idx >= 0) {
                allRows[idx] = enrichBelegungRow(allRows[idx], getTeacherLookup(), {});
            }
        }
        dirty = true;
        const hint = $('utkDirtyHint');
        if (hint) hint.hidden = false;
        if (field === 'lehrerCode') render();
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
                '"' +
                (linked ? ' class="utk-row-linked"' : '') +
                '>' +
                '<td><input type="text" data-field="klasse" value="' +
                escapeHtml(r.klasse || '') +
                '" aria-label="Klasse"></td>' +
                '<td><input type="text" data-field="fach" value="' +
                escapeHtml(r.fach || '') +
                '" aria-label="Fach"></td>' +
                '<td><input type="text" data-field="lehrerCode" value="' +
                escapeHtml(r.lehrerCode || '') +
                '" aria-label="Lehrkraft-Kürzel" title="Kürzel"></td>' +
                '<td class="utk-cell-readonly" title="Aus Stammdaten">' +
                escapeHtml(r.lehrerName || '—') +
                '</td>' +
                '<td><input type="text" data-field="lehrerEmail" value="' +
                escapeHtml(r.lehrerEmail || '') +
                '" aria-label="Lehrer E-Mail"></td>' +
                '<td><input type="text" class="utk-input--wide" data-field="teamName" value="' +
                escapeHtml(r.teamName || '') +
                '" aria-label="Anzeigename"></td>' +
                '<td><input type="text" data-field="gruppenmail" value="' +
                escapeHtml(r.gruppenmail || '') +
                '" aria-label="Mail-Nickname"></td>' +
                '<td class="utk-m365-cell">' +
                m365StatusCellHtml(r) +
                '</td>' +
                m365ActionCellHtml(linked, r.graphGroupId) +
                '</tr>'
            );
        })
        .join('');

    host.innerHTML =
        '<div class="utk-table-wrap"><table class="utk-table"><thead><tr>' +
        '<th>Klasse</th><th>Fach</th><th>LK</th><th>Name</th><th>Lehrer E-Mail</th><th>Anzeigename</th><th>Mail-Nickname</th><th>M365</th><th>Aktion</th>' +
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
        tr.querySelector('.utk-m365-link')?.addEventListener('click', () => linkOneRowByKey(key));
        tr.querySelector('.utk-m365-unlink')?.addEventListener('click', () => {
            if (!window.confirm('M365-Verknüpfung für diese Zeile entfernen?')) return;
            unlinkRowByKey(key);
        });
    });
}

function render() {
    readFiltersFromUi();
    fillFilterSelects();
    const filtered = filterUnterrichtsteamRows(allRows, filters);
    const api = window.ms365AppDataV2;
    const schoolYear =
        api && typeof api.getContainer === 'function'
            ? String(api.getContainer()?.years?.current || '').trim()
            : '';
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
            (schoolYear ? '<span>App-Schuljahr: <strong>' + escapeHtml(schoolYear) + '</strong></span>' : '') +
            (metaYearPrefix ? '<span>Team-Präfix: <strong>' + escapeHtml(metaYearPrefix) + '</strong></span>' : '');
    }
    const empty = $('utkEmpty');
    if (empty) {
        empty.hidden = allRows.length > 0;
        if (!allRows.length) {
            let otherHint = '';
            if (api && typeof api.getContainer === 'function') {
                const other = findUnterrichtsbelegungInOtherYears(api.getContainer(), schoolYear);
                if (other) {
                    otherHint =
                        ' Im Schuljahr <strong>' +
                        escapeHtml(other.year) +
                        '</strong> sind <strong>' +
                        other.count +
                        '</strong> Teams gespeichert – ggf. im Dashboard das aktive Schuljahr wechseln.';
                }
            }
            empty.innerHTML =
                'Noch keine Teams in den App-Daten für das aktuelle Schuljahr.' +
                otherHint +
                ' In <a href="kursteams.html">Kursteams</a> Team-Namen generieren oder unter ' +
                '<a href="kursteams.html#kursteamGraphImportPanel">Aus M365 importieren</a> laden – die Auswahl wird in die App-Daten geschrieben.';
        }
    }
    const hint = $('utkDirtyHint');
    if (hint) hint.hidden = !dirty;
    const yp = $('utkYearPrefix');
    if (yp && metaYearPrefix && !yp.value) yp.value = metaYearPrefix;
    renderTable(filtered);
    renderAbgleichPanel();
}

function pullFromWizard() {
    const state = loadWizardMergeFromBrowser();
    const api = window.ms365AppDataV2;
    const existing = api && typeof api.getUnterrichtsbelegung === 'function' ? api.getUnterrichtsbelegung() : null;
    const { ok, snapshot, imported, source } = hydrateUnterrichtsbelegungFromKursteamState(existing, state);
    if (!ok || !imported) {
        toast('Kein Kursteam-/WebUntis-Stand im Browser – zuerst in Kursteams importieren oder Team-Namen generieren.');
        return;
    }
    if (snapshot && snapshot.rows) {
        allRows = sortUnterrichtsteamRows(snapshot.rows);
        metaYearPrefix = snapshot.yearPrefix || metaYearPrefix;
        if (api && typeof api.setUnterrichtsbelegung === 'function') {
            api.setUnterrichtsbelegung(snapshot);
            dirty = false;
            toast(imported + ' Einträge aus Assistent übernommen und gespeichert (' + source + ').');
        } else {
            dirty = true;
            toast(imported + ' Einträge übernommen (' + source + '). App-Daten fehlen – bitte „Speichern“ nach Neuladen.');
        }
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
    $('utkEnrichTeachers')?.addEventListener('click', () => {
        enrichAllRowsFromTenant({ persist: false });
        render();
    });
    $('utkAddRow')?.addEventListener('click', () => addEmptyRow());
    $('utkM365Load')?.addEventListener('click', () => loadM365Teams());
    $('utkM365LinkAll')?.addEventListener('click', () => linkAllFromAbgleich());
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

async function init() {
    wireUi();
    await waitForAppDataV2();
    loadFromApp();
    const hydrated = tryHydrateFromKursteamBrowserState();
    if (hydrated.ok && hydrated.count) {
        toast(
            hydrated.count +
                ' Teams aus dem Kursteam-Zwischenstand in die App-Daten übernommen (' +
                (hydrated.source || 'kursteams') +
                ').'
        );
    }
    enrichAllRowsFromTenant({ quiet: true, persist: true });
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
