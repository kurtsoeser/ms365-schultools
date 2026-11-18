/**
 * Raumplan Admin-UI – Übersicht, Liste, Import, Plan, Technik, SPO Live Sync.
 */

import {
    filterRaeume,
    raeumeToCsv,
    raumStats,
    technikMeldungUrl,
    typLabel
} from './raumplan-logic.js';
import { RAUM_TYP, RAUM_TYP_LABEL } from './raumplan-model.js';
import {
    clearAll,
    clearDemoSeed,
    clearPlan,
    importRaeumeCsv,
    loadDemoData,
    loadState,
    removeRaum,
    setPlanImage,
    setPlanPdfName,
    upsertRaum
} from './raumplan-store.js';
import {
    getRaumplanAccount,
    getRaumplanAdminAccessCache,
    isRaumplanAdmin,
    resolveRaumplanAdmin
} from './raumplan-entra-role.js';
import { resolveRaumplanSiteUrl } from './raumplan-settings.js';
import {
    getRaumplanSyncMeta,
    pullRaumplanFromSharePoint,
    pushRaumplanDelete,
    pushRaumplanRaum
} from './raumplan-sync.js';

const VIEWS = [
    { id: 'uebersicht', label: 'Übersicht', icon: 'bi-speedometer2' },
    { id: 'liste', label: 'Raumliste', icon: 'bi-door-open' },
    { id: 'import', label: 'Import', icon: 'bi-upload' },
    { id: 'plan', label: 'Plan', icon: 'bi-map' },
    { id: 'technik', label: 'Technik', icon: 'bi-wrench' },
    { id: 'system', label: 'System', icon: 'bi-gear' }
];

const WRITE_VIEWS = new Set(['import', 'system']);

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function formatDeDate(iso) {
    const s = String(iso || '').slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return s || '–';
    const [y, m, d] = s.split('-');
    return d + '.' + m + '.' + y;
}

/**
 * @param {HTMLElement} root
 */
export function mountApp(root) {
    if (!root) return;

    /** @type {{ view: string, message: {type:string,text:string}|null, filters: object, editId: string }} */
    const ui = {
        view: 'uebersicht',
        message: null,
        filters: { q: '', typ: '', gebaeude: '' },
        editId: ''
    };

    function setMessage(type, text) {
        ui.message = text ? { type, text } : null;
    }

    function state() {
        return loadState();
    }

    function admin() {
        return isRaumplanAdmin(getRaumplanAccount().email);
    }

    async function afterLocalWrite(item, okMsg) {
        const push = await pushRaumplanRaum(item);
        if (push.ok) {
            setMessage('ok', okMsg || 'In SharePoint gespeichert.');
        } else if (push.skipped) {
            setMessage('ok', okMsg || 'Gespeichert (Demo lokal).');
        } else {
            setMessage(
                'warn',
                (okMsg ? okMsg + ' ' : '') +
                    (push.error || 'SharePoint-Sync fehlgeschlagen – lokal gespeichert.')
            );
        }
        render();
        return push;
    }

    function render() {
        if (!admin() && WRITE_VIEWS.has(ui.view)) ui.view = 'uebersicht';
        root.innerHTML =
            '<div class="ag-shell">' +
            renderNav() +
            '<div class="ag-main">' +
            renderMessage() +
            renderSyncHint() +
            renderView() +
            '</div></div>';
        bind();
    }

    function renderNav() {
        const isAdm = admin();
        const role = getRaumplanAdminAccessCache();
        const buttons = VIEWS.filter((v) => isAdm || !WRITE_VIEWS.has(v.id))
            .map(
                (v) =>
                    '<button type="button" class="ag-nav__btn' +
                    (ui.view === v.id ? ' is-active' : '') +
                    '" data-rp-view="' +
                    esc(v.id) +
                    '"><i class="bi ' +
                    esc(v.icon) +
                    '" aria-hidden="true"></i>' +
                    esc(v.label) +
                    '</button>'
            )
            .join('');
        return (
            '<nav class="ag-nav" aria-label="Raumplan">' +
            '<div class="ag-nav__brand">' +
            '<i class="bi bi-map" aria-hidden="true"></i>' +
            '<div><strong>Raumplan</strong><span>Liste &amp; Plan · SPO Live</span></div>' +
            '</div>' +
            '<div class="ag-nav__list">' +
            buttons +
            '</div>' +
            '<p class="ag-nav__hint">Lehrkräfte lesen/suchen; Import &amp; Pflege nur Admin (Schulleitung/IT).' +
            (isAdm
                ? ' · Admin' +
                  (role.demo ? ' (Demo)' : role.sources.length ? ' (' + esc(role.sources.join(', ')) + ')' : '')
                : ' · nur Lesen') +
            '</p>' +
            '</nav>'
        );
    }

    function renderMessage() {
        if (!ui.message) return '';
        const cls =
            ui.message.type === 'ok'
                ? 'ag-msg--ok'
                : ui.message.type === 'warn'
                  ? 'ag-msg--warn'
                  : 'ag-msg--err';
        return '<p class="ag-msg ' + cls + '" role="status">' + esc(ui.message.text) + '</p>';
    }

    function renderCalmBanner(opts) {
        const eyebrow = opts.eyebrow || '';
        const text = opts.text || '';
        const kind = opts.kind === 'ops' ? 'tm-calm-banner--ops' : 'tm-calm-banner--action';
        return (
            '<section class="tm-calm-banner ' +
            kind +
            '" role="status">' +
            '<div class="tm-calm-banner__text">' +
            (eyebrow ? '<p class="tm-calm-banner__eyebrow">' + esc(eyebrow) + '</p>' : '') +
            '<p>' +
            esc(text) +
            '</p></div></section>'
        );
    }

    function renderSyncHint() {
        if (ui.view === 'system') return '';
        const sync = getRaumplanSyncMeta();
        if (sync.source === 'live' && sync.ok) return '';
        if (!sync.siteUrl && sync.source !== 'demo') {
            return renderCalmBanner({
                kind: 'ops',
                eyebrow: 'SharePoint',
                text: 'Noch keine Site – IT richtet die Liste ein. Bis dahin lokal/Demo. Plan-Bild bleibt immer lokal.'
            });
        }
        if (sync.source === 'demo' || sync.source === 'local') {
            return renderCalmBanner({
                kind: 'ops',
                eyebrow: 'Lokal / Demo',
                text:
                    sync.error ||
                    'SharePoint nicht verbunden – Räume bleiben lokal, bis die Liste greift.'
            });
        }
        return '';
    }

    function renderView() {
        if (ui.view === 'liste') return renderListe();
        if (ui.view === 'import') return renderImport();
        if (ui.view === 'plan') return renderPlan();
        if (ui.view === 'technik') return renderTechnik();
        if (ui.view === 'system') return renderSystem();
        return renderUebersicht();
    }

    function card(label, value) {
        return (
            '<div class="ag-card"><p class="ag-card__label">' +
            esc(label) +
            '</p><p class="ag-card__value">' +
            esc(value) +
            '</p></div>'
        );
    }

    function typOptions(selected) {
        return Object.entries(RAUM_TYP_LABEL)
            .map(
                ([id, label]) =>
                    '<option value="' +
                    esc(id) +
                    '"' +
                    (id === selected ? ' selected' : '') +
                    '>' +
                    esc(label) +
                    '</option>'
            )
            .join('');
    }

    function gebaeudeOptions(selected) {
        const set = new Set();
        for (const r of state().raeume) {
            if (r.gebaeude) set.add(r.gebaeude);
        }
        return [...set]
            .sort((a, b) => a.localeCompare(b, 'de'))
            .map(
                (g) =>
                    '<option value="' +
                    esc(g) +
                    '"' +
                    (g === selected ? ' selected' : '') +
                    '>' +
                    esc(g) +
                    '</option>'
            )
            .join('');
    }

    function renderUebersicht() {
        const s = state();
        const st = raumStats(s.raeume);
        const hits = filterRaeume(s.raeume, { q: ui.filters.q }).slice(0, 8);
        const hitHtml = hits.length
            ? hits
                  .map(
                      (r) =>
                          '<article class="ag-item"><h3 class="ag-item__title">' +
                          esc(r.code) +
                          ' · ' +
                          esc(r.name) +
                          '</h3><p class="ag-item__meta">' +
                          esc(typLabel(r.typ)) +
                          ' · ' +
                          esc(r.gebaeude || '–') +
                          ' · Stock ' +
                          esc(r.stockwerk || '–') +
                          ' · Kap. ' +
                          esc(r.kapazitaet) +
                          '</p></article>'
                  )
                  .join('')
            : '<p class="ag-item__meta">Keine Treffer.</p>';

        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Übersicht</h2>' +
            '<div class="ag-cards">' +
            card('Räume', st.total) +
            card('Gebäude', st.gebaeude) +
            card('Fachräume', st.fachraeume) +
            card('Kapazität Σ', st.kapazitaet) +
            '</div>' +
            '<div class="ag-filters">' +
            '<input id="rpSearchQ" placeholder="Raum suchen (Code, Name …)" value="' +
            esc(ui.filters.q) +
            '">' +
            '<button type="button" class="btn" id="rpSearchApply">Suchen</button>' +
            '</div>' +
            '<div class="ag-list">' +
            hitHtml +
            '</div>' +
            '<div class="ag-toolbar" style="margin-top:14px;">' +
            '<button type="button" class="btn btn-success" data-rp-go="liste">Raumliste</button>' +
            '<button type="button" class="btn" data-rp-go="plan">Plan anzeigen</button>' +
            (admin()
                ? '<button type="button" class="btn" data-rp-go="import">CSV-Import</button>'
                : '') +
            '</div>'
        );
    }

    function renderListe() {
        const s = state();
        const list = filterRaeume(s.raeume, ui.filters);
        const isAdm = admin();
        if (!isAdm) ui.editId = '';
        const rows = list
            .map(
                (r) =>
                    '<tr>' +
                    '<td><strong>' +
                    esc(r.code) +
                    '</strong></td>' +
                    '<td>' +
                    esc(r.name) +
                    '</td>' +
                    '<td>' +
                    esc(r.gebaeude || '–') +
                    '</td>' +
                    '<td>' +
                    esc(r.stockwerk || '–') +
                    '</td>' +
                    '<td>' +
                    esc(r.kapazitaet) +
                    '</td>' +
                    '<td><span class="ag-pill">' +
                    esc(typLabel(r.typ)) +
                    '</span></td>' +
                    '<td>' +
                    (isAdm
                        ? '<button type="button" class="btn" data-rp-edit="' +
                          esc(r.id) +
                          '">Bearbeiten</button> ' +
                          '<button type="button" class="btn" data-rp-del="' +
                          esc(r.id) +
                          '">Löschen</button> '
                        : '') +
                    '<a class="btn" href="' +
                    esc(technikMeldungUrl(r.code)) +
                    '">Technik</a>' +
                    '</td></tr>'
            )
            .join('');

        const edit =
            isAdm && ui.editId
                ? ui.editId === 'new'
                    ? {
                          id: '',
                          code: '',
                          name: '',
                          gebaeude: '',
                          stockwerk: '',
                          kapazitaet: 0,
                          typ: RAUM_TYP.KLASSE,
                          notiz: ''
                      }
                    : s.raeume.find((r) => r.id === ui.editId)
                : null;

        const form = edit
            ? '<form class="ag-form" id="rpRaumForm" style="margin-bottom:14px;">' +
              '<input type="hidden" name="id" value="' +
              esc(edit.id) +
              '">' +
              '<div class="ag-form__row">' +
              '<label>Code<input name="code" required value="' +
              esc(edit.code) +
              '"></label>' +
              '<label>Name<input name="name" required value="' +
              esc(edit.name) +
              '"></label></div>' +
              '<div class="ag-form__row">' +
              '<label>Gebäude<input name="gebaeude" value="' +
              esc(edit.gebaeude) +
              '"></label>' +
              '<label>Stockwerk<input name="stockwerk" value="' +
              esc(edit.stockwerk) +
              '"></label></div>' +
              '<div class="ag-form__row">' +
              '<label>Kapazität<input name="kapazitaet" type="number" min="0" value="' +
              esc(edit.kapazitaet) +
              '"></label>' +
              '<label>Typ<select name="typ">' +
              typOptions(edit.typ) +
              '</select></label></div>' +
              '<label>Notiz<textarea name="notiz" rows="2">' +
              esc(edit.notiz) +
              '</textarea></label>' +
              '<div class="ag-form__actions">' +
              '<button type="submit" class="btn btn-success">Speichern</button>' +
              '<button type="button" class="btn" data-rp-cancel-edit>Abbrechen</button>' +
              '</div></form>'
            : '';

        return (
            '<h2 style="margin:0 0 10px;font-size:1.15rem;">Raumliste</h2>' +
            '<div class="ag-filters">' +
            '<select id="rpFilterTyp"><option value="">Alle Typen</option>' +
            typOptions(ui.filters.typ) +
            '</select>' +
            '<select id="rpFilterGeb"><option value="">Alle Gebäude</option>' +
            gebaeudeOptions(ui.filters.gebaeude) +
            '</select>' +
            '<input id="rpFilterQ" placeholder="Suche" value="' +
            esc(ui.filters.q) +
            '">' +
            '<button type="button" class="btn" id="rpFilterApply">Filtern</button>' +
            (isAdm
                ? '<button type="button" class="btn btn-success" data-rp-new>Neu</button>'
                : '') +
            '</div>' +
            form +
            '<div style="overflow:auto;">' +
            '<table class="ag-table" style="width:100%;border-collapse:collapse;font-size:0.9rem;">' +
            '<thead><tr>' +
            '<th style="text-align:left;padding:8px 6px;border-bottom:1px solid var(--border);">Code</th>' +
            '<th style="text-align:left;padding:8px 6px;border-bottom:1px solid var(--border);">Name</th>' +
            '<th style="text-align:left;padding:8px 6px;border-bottom:1px solid var(--border);">Gebäude</th>' +
            '<th style="text-align:left;padding:8px 6px;border-bottom:1px solid var(--border);">Stock</th>' +
            '<th style="text-align:left;padding:8px 6px;border-bottom:1px solid var(--border);">Kap.</th>' +
            '<th style="text-align:left;padding:8px 6px;border-bottom:1px solid var(--border);">Typ</th>' +
            '<th style="text-align:left;padding:8px 6px;border-bottom:1px solid var(--border);">Aktionen</th>' +
            '</tr></thead><tbody>' +
            (rows ||
                '<tr><td colspan="7" style="padding:12px;">Keine Räume.</td></tr>') +
            '</tbody></table></div>'
        );
    }

    function renderImport() {
        if (!admin()) {
            return (
                '<div class="ag-head"><h2>CSV-Import</h2><p>Nur Admin (Schulleitung/IT).</p></div>'
            );
        }
        return (
            '<h2 style="margin:0 0 10px;font-size:1.15rem;">CSV-Import</h2>' +
            '<p class="ag-item__meta">Format: <code>Code;Name;Gebaeude;Stockwerk;Kapazitaet;Typ</code> – Typ: klasse, fachraum, buero, sporthalle, sonstiges. Import schreibt nach SharePoint (kein Demo-Seed).</p>' +
            '<form class="ag-form" id="rpImportForm">' +
            '<label>CSV einfügen<textarea name="csv" rows="10" placeholder="Code;Name;Gebaeude;Stockwerk;Kapazitaet;Typ&#10;INF1;Informatik 1;Trakt B;1;28;fachraum"></textarea></label>' +
            '<div class="ag-form__actions"><button type="submit" class="btn btn-success">Importieren</button></div>' +
            '</form>' +
            '<div class="ag-toolbar" style="margin-top:16px;">' +
            '<button type="button" class="btn" data-rp-export>CSV exportieren</button>' +
            '</div>'
        );
    }

    function renderSystem() {
        if (!admin()) {
            return '<div class="ag-head"><h2>System</h2><p>Nur Admin.</p></div>';
        }
        const sync = getRaumplanSyncMeta();
        const site = resolveRaumplanSiteUrl() || sync.siteUrl || '– nicht gesetzt –';
        const s = state();
        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">System</h2>' +
            '<div class="tm-system-panel">' +
            '<h3>SharePoint (Live)</h3>' +
            '<p><code>' +
            esc(site) +
            '</code></p>' +
            '<p>Sync: <strong>' +
            esc(sync.source) +
            '</strong>' +
            (sync.listTitle ? ' · Liste „' + esc(sync.listTitle) + '“' : '') +
            (sync.lastPullAt ? ' · zuletzt ' + esc(formatDeDate(sync.lastPullAt)) : '') +
            (sync.connected ? ' · verbunden' : ' · nicht verbunden') +
            (sync.error ? ' · ' + esc(sync.error) : '') +
            '</p>' +
            '<p>Lokal: ' +
            (s.raeume || []).length +
            ' Räume' +
            (s.localDemoOnly ? ' · Demo local-first' : '') +
            ' · mit spoItemId: ' +
            (s.raeume || []).filter((r) => r.spoItemId).length +
            '</p>' +
            '<p class="ag-item__meta">Plan-Bild/PDF bleibt browser-lokal (kein SPO-SoR).</p>' +
            '<div class="ag-toolbar">' +
            '<button type="button" class="btn btn-success" data-rp-pull>Von SharePoint laden</button>' +
            '<button type="button" class="btn" data-rp-demo>Demo laden</button>' +
            '<button type="button" class="btn" data-rp-clear-seed>Demo-Seed entfernen</button>' +
            '<button type="button" class="btn" data-rp-clear>Lokal leeren</button>' +
            '<a class="btn" href="sharepoint-liste-raumplan.html"><i class="bi bi-list-ul"></i> Listen einrichten</a>' +
            '</div></div>'
        );
    }

    function renderPlan() {
        const s = state();
        let preview = '';
        if (s.planImageDataUrl) {
            preview =
                '<div style="margin-top:12px;"><img src="' +
                esc(s.planImageDataUrl) +
                '" alt="Raumplan" style="max-width:100%;border:1px solid var(--border);border-radius:12px;background:#fff;"></div>';
        } else if (s.planPdfName) {
            preview =
                '<p class="ag-item__meta" style="margin-top:12px;">PDF hinterlegt: <strong>' +
                esc(s.planPdfName) +
                '</strong> (nur Hinweis – PDF wird lokal nicht als Bild gerendert).</p>';
        } else {
            preview =
                '<p class="ag-item__meta" style="margin-top:12px;">Noch kein Plan. Bild (PNG/JPG/WebP) hochladen oder PDF-Dateiname notieren.</p>';
        }

        const isAdm = admin();
        return (
            '<h2 style="margin:0 0 10px;font-size:1.15rem;">Plan-Viewer</h2>' +
            (isAdm
                ? '<div class="ag-form">' +
                  '<label>Bild hochladen (wird als Data-URL lokal gespeichert)' +
                  '<input type="file" id="rpPlanImage" accept="image/png,image/jpeg,image/webp,image/gif">' +
                  '</label>' +
                  '<label>PDF (nur Dateiname / Hinweis)<input type="file" id="rpPlanPdf" accept="application/pdf"></label>' +
                  '<div class="ag-form__actions">' +
                  '<button type="button" class="btn" data-rp-clear-plan>Plan entfernen</button>' +
                  '</div></div>'
                : '<p class="ag-item__meta">Plan ansehen – Upload nur für Admin.</p>') +
            preview
        );
    }

    function renderTechnik() {
        const s = state();
        const list = filterRaeume(s.raeume, { q: ui.filters.q }).slice(0, 20);
        const items = list
            .map(
                (r) =>
                    '<article class="ag-item">' +
                    '<h3 class="ag-item__title">' +
                    esc(r.code) +
                    ' · ' +
                    esc(r.name) +
                    '</h3>' +
                    '<p class="ag-item__meta">' +
                    esc(r.gebaeude || '–') +
                    ' · ' +
                    esc(typLabel(r.typ)) +
                    '</p>' +
                    '<div class="ag-item__actions">' +
                    '<a class="btn btn-success" href="' +
                    esc(technikMeldungUrl(r.code)) +
                    '"><i class="bi bi-wrench"></i> Technik melden</a>' +
                    '</div></article>'
            )
            .join('');

        return (
            '<h2 style="margin:0 0 10px;font-size:1.15rem;">Technik-Meldungen</h2>' +
            '<p class="ag-item__meta">Deep-Link: <code>technik-meldungen.html?raum=CODE</code> – Raumcode wird vorausgefüllt.</p>' +
            '<div class="ag-filters">' +
            '<input id="rpTechQ" placeholder="Raum filtern" value="' +
            esc(ui.filters.q) +
            '">' +
            '<button type="button" class="btn" id="rpTechApply">Filtern</button>' +
            '</div>' +
            '<div class="ag-list">' +
            (items || '<p class="ag-item__meta">Keine Räume.</p>') +
            '</div>'
        );
    }

    function downloadText(filename, text, mime) {
        const blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        a.click();
        URL.revokeObjectURL(a.href);
    }

    function readFileAsDataUrl(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(String(reader.result || ''));
            reader.onerror = () => reject(reader.error || new Error('Lesen fehlgeschlagen'));
            reader.readAsDataURL(file);
        });
    }

    function bind() {
        root.querySelectorAll('[data-rp-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = btn.getAttribute('data-rp-view') || 'uebersicht';
                ui.message = null;
                ui.editId = '';
                render();
            });
        });
        root.querySelectorAll('[data-rp-go]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = btn.getAttribute('data-rp-go') || 'uebersicht';
                render();
            });
        });

        root.querySelector('#rpSearchApply')?.addEventListener('click', () => {
            ui.filters.q = root.querySelector('#rpSearchQ')?.value || '';
            render();
        });
        root.querySelector('#rpSearchQ')?.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter') {
                ui.filters.q = root.querySelector('#rpSearchQ')?.value || '';
                render();
            }
        });

        root.querySelector('#rpFilterApply')?.addEventListener('click', () => {
            ui.filters.typ = root.querySelector('#rpFilterTyp')?.value || '';
            ui.filters.gebaeude = root.querySelector('#rpFilterGeb')?.value || '';
            ui.filters.q = root.querySelector('#rpFilterQ')?.value || '';
            render();
        });

        root.querySelector('#rpTechApply')?.addEventListener('click', () => {
            ui.filters.q = root.querySelector('#rpTechQ')?.value || '';
            render();
        });

        root.querySelector('[data-rp-new]')?.addEventListener('click', () => {
            if (!admin()) return;
            ui.editId = 'new';
            render();
        });
        root.querySelector('[data-rp-cancel-edit]')?.addEventListener('click', () => {
            ui.editId = '';
            render();
        });
        root.querySelectorAll('[data-rp-edit]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!admin()) return;
                ui.editId = btn.getAttribute('data-rp-edit') || '';
                render();
            });
        });
        root.querySelectorAll('[data-rp-del]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!admin()) return;
                if (!window.confirm('Raum löschen?')) return;
                const id = btn.getAttribute('data-rp-del') || '';
                const res = removeRaum(id);
                if (!res.ok) {
                    setMessage('err', (res.errors || []).join(' '));
                    render();
                    return;
                }
                void pushRaumplanDelete(res.item).then((del) => {
                    if (del.ok) setMessage('ok', 'Gelöscht.');
                    else
                        setMessage(
                            'warn',
                            'Lokal gelöscht. ' + (del.error || 'SharePoint-Löschen fehlgeschlagen.')
                        );
                    render();
                });
            });
        });

        root.querySelector('#rpRaumForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            if (!admin()) return;
            const fd = new FormData(ev.target);
            const res = upsertRaum({
                id: String(fd.get('id') || '') || undefined,
                code: String(fd.get('code') || ''),
                name: String(fd.get('name') || ''),
                gebaeude: String(fd.get('gebaeude') || ''),
                stockwerk: String(fd.get('stockwerk') || ''),
                kapazitaet: Number(fd.get('kapazitaet') || 0),
                typ: String(fd.get('typ') || RAUM_TYP.SONSTIGES),
                notiz: String(fd.get('notiz') || '')
            });
            if (!res.ok) {
                setMessage('err', (res.errors || []).join(' '));
                render();
                return;
            }
            ui.editId = '';
            void afterLocalWrite(res.item, 'Gespeichert.');
        });

        root.querySelector('#rpImportForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            if (!admin()) return;
            const fd = new FormData(ev.target);
            const res = importRaeumeCsv(String(fd.get('csv') || ''));
            if (!(res.added || res.updated)) {
                setMessage('err', res.errors.slice(0, 3).join(' ') || 'Kein Import.');
                render();
                return;
            }
            ui.view = 'liste';
            void (async () => {
                let pushFail = 0;
                for (const item of res.changed || []) {
                    const push = await pushRaumplanRaum(item);
                    if (!push.ok && !push.skipped) pushFail += 1;
                }
                setMessage(
                    pushFail ? 'warn' : 'ok',
                    res.added +
                        ' neu, ' +
                        res.updated +
                        ' aktualisiert.' +
                        (pushFail
                            ? ' ' + pushFail + ' SharePoint-Push fehlgeschlagen – lokal gespeichert.'
                            : ' In SharePoint übernommen.') +
                        (res.errors.length ? ' ' + res.errors.slice(0, 3).join(' ') : '')
                );
                render();
            })();
        });

        root.querySelector('[data-rp-export]')?.addEventListener('click', () => {
            downloadText('raeume.csv', raeumeToCsv(state().raeume), 'text/csv;charset=utf-8');
        });
        root.querySelector('[data-rp-demo]')?.addEventListener('click', () => {
            if (!admin()) return;
            loadDemoData();
            setMessage('ok', 'Demo geladen (nur lokal, kein SPO-Push).');
            ui.view = 'liste';
            render();
        });
        root.querySelector('[data-rp-clear-seed]')?.addEventListener('click', () => {
            if (!admin()) return;
            clearDemoSeed();
            setMessage('ok', 'Demo-Seed entfernt.');
            render();
        });
        root.querySelector('[data-rp-clear]')?.addEventListener('click', () => {
            if (!admin()) return;
            if (!window.confirm('Lokale Raumplan-Daten leeren?')) return;
            clearAll();
            setMessage('ok', 'Geleert.');
            render();
        });
        root.querySelector('[data-rp-pull]')?.addEventListener('click', () => {
            if (!admin()) return;
            void pullRaumplanFromSharePoint({ forceDemoIfEmpty: false }).then((r) => {
                setMessage(
                    r.ok ? 'ok' : 'warn',
                    r.ok
                        ? 'Von SharePoint geladen.'
                        : r.softError || r.meta.error || 'Pull fehlgeschlagen.'
                );
                render();
            });
        });

        root.querySelector('#rpPlanImage')?.addEventListener('change', async (ev) => {
            if (!admin()) return;
            const file = ev.target.files && ev.target.files[0];
            if (!file) return;
            if (!String(file.type || '').startsWith('image/')) {
                setMessage('err', 'Nur Bilddateien (PNG/JPG/WebP).');
                render();
                return;
            }
            try {
                const dataUrl = await readFileAsDataUrl(file);
                setPlanImage(dataUrl);
                setMessage('ok', 'Plan-Bild lokal gespeichert.');
            } catch {
                setMessage('err', 'Bild konnte nicht gelesen werden.');
            }
            render();
        });

        root.querySelector('#rpPlanPdf')?.addEventListener('change', (ev) => {
            if (!admin()) return;
            const file = ev.target.files && ev.target.files[0];
            if (!file) return;
            setPlanPdfName(file.name || 'plan.pdf');
            setMessage('warn', 'PDF als Hinweis gespeichert (kein Inline-Viewer).');
            render();
        });

        root.querySelector('[data-rp-clear-plan]')?.addEventListener('click', () => {
            if (!admin()) return;
            clearPlan();
            setMessage('ok', 'Plan entfernt.');
            render();
        });
    }

    render();
    void resolveRaumplanAdmin(getRaumplanAccount().email).then(() => render());
    void pullRaumplanFromSharePoint().then((r) => {
        if (r.softError) setMessage('warn', r.softError);
        else if (!r.ok && r.meta && r.meta.error && r.meta.source === 'none') {
            /* calm banner covers missing site */
        }
        render();
    });
}
