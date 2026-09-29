/**
 * UI-Rendering Schulaktivitäten-Planer.
 */
import {
    computeDashboardKpis,
    monthGridDates,
    itemCoversDay,
    validateAktivitaet,
    DEFAULT_RULES
} from './schulaktivitaeten-planer-logic.js';
import {
    viewsForRole,
    filterItems,
    labelMaps,
    formatDeDate,
    statusLabel,
    typLabel,
    roleLabel,
    emptyForm,
    canEditItem,
    canDecide,
    scopeFromState,
    typChoices
} from './schulaktivitaeten-planer-state.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function statusBadge(status) {
    const s = String(status || 'beantragt').toLowerCase();
    const cls =
        s === 'genehmigt' ? 'akt-badge akt-badge--ok' : s === 'abgelehnt' ? 'akt-badge akt-badge--bad' : 'akt-badge akt-badge--warn';
    return `<span class="${cls}">${esc(statusLabel(s))}</span>`;
}

function optionList(items, selected, emptyLabel) {
    const opts = [`<option value="">${esc(emptyLabel)}</option>`];
    (items || []).forEach((it) => {
        const v = it.code || it.value;
        const label = it.name || it.label || v;
        opts.push(
            `<option value="${esc(v)}"${String(selected) === String(v) ? ' selected' : ''}>${esc(label)}</option>`
        );
    });
    return opts.join('');
}

export function renderApp(state, root) {
    if (!root) return;
    const views = viewsForRole(state.role);
    root.innerHTML = `
    <div class="akt-shell">
      <aside class="akt-nav" aria-label="Navigation">
        <div class="akt-nav__brand">
          <i class="bi bi-bus-front-fill" aria-hidden="true"></i>
          <div>
            <strong>Schulaktivitäten</strong>
            <span>Exkursionen &amp; Anträge</span>
          </div>
        </div>
        <nav class="akt-nav__list">
          ${views
              .map(
                  (v) => `
            <button type="button" class="akt-nav__btn${state.view === v.id ? ' is-active' : ''}" data-akt-view="${esc(v.id)}">
              <i class="bi ${v.icon}" aria-hidden="true"></i>${esc(v.label)}
            </button>`
              )
              .join('')}
        </nav>
        <div class="akt-nav__role">
          <span class="akt-nav__role-label">Rolle (Demo)</span>
          <div class="akt-nav__role-btns">
            <button type="button" class="akt-chip${state.role === 'lehrer' ? ' is-active' : ''}" data-akt-role="lehrer">Lehrer</button>
            <button type="button" class="akt-chip${state.role === 'admin' ? ' is-active' : ''}" data-akt-role="admin">Admin</button>
          </div>
          <p class="akt-nav__hint">Später Entra-Gruppen. Demo-Umschalter nur zur Vorschau.</p>
        </div>
      </aside>
      <div class="akt-main">
        <header class="akt-top">
          <div class="akt-top__site">
            <label for="aktSiteUrl">SharePoint-Site</label>
            <div class="akt-top__row">
              <input type="url" id="aktSiteUrl" value="${esc(state.siteUrl)}" placeholder="https://…sharepoint.com/sites/Intranet" spellcheck="false">
              <button type="button" class="btn" id="aktBtnLoad"><i class="bi bi-arrow-repeat"></i>Laden</button>
              <a class="btn" href="sharepoint-liste-schulaktivitaeten.html" style="text-decoration:none"><i class="bi bi-list-ul"></i>Listen</a>
              <button type="button" class="btn" id="aktBtnIcs"><i class="bi bi-download"></i>ICS</button>
              <button type="button" class="btn" id="aktBtnDemo" title="SJ 2026/27 Demo laden (lokal, optional SharePoint)"><i class="bi bi-database"></i>Demo</button>
              <button type="button" class="btn" id="aktBtnDemoReset" title="Demo-Einträge löschen / zurücksetzen"><i class="bi bi-trash"></i>Demo reset</button>
            </div>
          </div>
          <div class="akt-top__user">
            <span>Angemeldet als <span class="akt-badge akt-badge--info">${esc(roleLabel(state.role))}</span></span>
            ${
                state.localDemoOnly
                    ? '<span class="akt-badge akt-badge--warn">Demo lokal</span>'
                    : ''
            }
            ${
                state.accountEmail
                    ? `<small>${esc(state.accountName || state.accountEmail)}</small>`
                    : '<small class="muted">Bitte oben rechts anmelden</small>'
            }
          </div>
        </header>
        ${
            state.localDemoOnly
                ? '<div class="akt-alert akt-alert--info" role="status">Lokaler Demo-Modus (SJ 2026/27) – Anzeige ohne SharePoint. Mit „Demo“ erneut starten und SharePoint-Schreiben bestätigen, oder Site laden.</div>'
                : ''
        }
        ${state.error ? `<div class="akt-alert akt-alert--bad" role="alert">${esc(state.error)}</div>` : ''}
        ${state.loading ? '<div class="akt-skel" aria-busy="true"><div class="akt-skel__card"></div><div class="akt-skel__card"></div></div>' : ''}
        <div id="aktView" class="akt-view${state.loading ? ' is-loading' : ''}">
          ${state.loading ? '' : renderView(state)}
        </div>
      </div>
    </div>
    ${state.detailId ? renderDetailModal(state) : ''}
  `;
}

function filterBar(state, opts) {
    const sd = state.stammdaten;
    const scope = scopeFromState(state, opts || {});
    const count = filterItems(state.items, state.filters, scope).length;
    return `
    <div class="akt-filters" data-akt-filters>
      <select id="aktFilterKlasse" aria-label="Klasse">${optionList(sd.classes, state.filters.klasse, 'Alle Klassen')}</select>
      <select id="aktFilterTyp" aria-label="Typ">${optionList(typChoices(), state.filters.typ, 'Alle Typen')}</select>
      <select id="aktFilterStatus" aria-label="Status">
        <option value="">Alle Status</option>
        <option value="beantragt"${state.filters.status === 'beantragt' ? ' selected' : ''}>Beantragt</option>
        <option value="genehmigt"${state.filters.status === 'genehmigt' ? ' selected' : ''}>Genehmigt</option>
        <option value="abgelehnt"${state.filters.status === 'abgelehnt' ? ' selected' : ''}>Abgelehnt</option>
      </select>
      <select id="aktFilterLehrer" aria-label="Lehrer">${optionList(sd.teachers, state.filters.lehrer, 'Alle Lehrkräfte')}</select>
      <button type="button" class="btn" id="aktFilterReset"><i class="bi bi-arrow-counterclockwise"></i>Reset</button>
      <span class="akt-filters__meta">${count} Einträge</span>
    </div>`;
}

function renderView(state) {
    switch (state.view) {
        case 'liste':
            return renderListe(state);
        case 'kalender':
            return renderKalender(state);
        case 'antrag':
            return renderAntrag(state);
        case 'freigabe':
            return renderFreigabe(state);
        case 'regeln':
            return renderRegeln(state);
        default:
            return renderDashboard(state);
    }
}

function renderDashboard(state) {
    const scope = scopeFromState(state, { scopeAll: state.role === 'admin' });
    const items = filterItems(state.items, {}, scope);
    const kpi = computeDashboardKpis(items);
    const offen = filterItems(state.items, { status: 'beantragt' }, { ...scope, onlyOpen: false }).filter(
        (i) => String(i.status).toLowerCase() === 'beantragt'
    );
    return `
    <section class="akt-panel">
      <h2>Übersicht</h2>
      <div class="akt-kpi-row">
        <div class="akt-kpi"><strong>${kpi.offen}</strong><span>Offene Anträge</span></div>
        <div class="akt-kpi"><strong>${kpi.genehmigt}</strong><span>Genehmigt</span></div>
        <div class="akt-kpi"><strong>${kpi.demnaechst}</strong><span>In 14 Tagen</span></div>
        <div class="akt-kpi"><strong>${kpi.gesamt}</strong><span>Gesamt (Filter)</span></div>
      </div>
    </section>
    <section class="akt-panel">
      <h3>Offene Anträge</h3>
      ${offen.length ? renderTable(offen, state) : '<p class="muted">Keine offenen Anträge.</p>'}
      <div class="akt-actions" style="margin-top:12px">
        <button type="button" class="btn btn-success" data-akt-view-jump="antrag"><i class="bi bi-plus-lg"></i>Neuer Antrag</button>
        ${canDecide(state) ? '<button type="button" class="btn" data-akt-view-jump="freigabe"><i class="bi bi-check2-square"></i>Zur Freigabe</button>' : ''}
      </div>
    </section>`;
}

function renderListe(state) {
    const scope = scopeFromState(state, { scopeAll: state.role === 'admin' });
    const items = filterItems(state.items, state.filters, scope);
    return `
    <section class="akt-panel">
      <h2>Liste</h2>
      ${filterBar(state, { scopeAll: state.role === 'admin' })}
      ${items.length ? renderTable(items, state) : '<p class="muted">Keine Einträge.</p>'}
    </section>`;
}

function renderTable(items, state) {
    const labels = labelMaps(state.stammdaten);
    return `
    <div class="akt-table-wrap">
      <table class="akt-table">
        <thead>
          <tr>
            <th>Zeitraum</th><th>Titel</th><th>Typ</th><th>Klasse</th><th>Ort</th><th>Status</th><th></th>
          </tr>
        </thead>
        <tbody>
          ${items
              .map((it) => {
                  const zeit =
                      formatDeDate(it.startdatum) +
                      (it.enddatum && it.enddatum !== it.startdatum ? ' – ' + formatDeDate(it.enddatum) : '');
                  return `<tr>
              <td>${esc(zeit)}</td>
              <td>${esc(it.titel)}</td>
              <td>${esc(typLabel(it.typ))}</td>
              <td>${esc(labels.klasse[it.klasseCode] || it.klasseCode)}</td>
              <td>${esc(it.ort || '–')}</td>
              <td>${statusBadge(it.status)}</td>
              <td>
                <button type="button" class="btn btn-sm" data-akt-detail="${esc(it.aktivitaetId || it.itemId)}">Details</button>
              </td>
            </tr>`;
              })
              .join('')}
        </tbody>
      </table>
    </div>`;
}

function renderKalender(state) {
    const y = state.calYear;
    const m = state.calMonth;
    const days = monthGridDates(y, m);
    const scope = scopeFromState(state, { scopeAll: true });
    const items = filterItems(state.items, state.filters, scope).filter(
        (i) => String(i.status).toLowerCase() !== 'abgelehnt'
    );
    const monthName = new Date(y, m - 1, 1).toLocaleDateString('de-AT', { month: 'long', year: 'numeric' });
    const cells = days
        .map((day) => {
            const inMonth = Number(day.slice(5, 7)) === m;
            const dayItems = items.filter((it) => itemCoversDay(it, day));
            return `<div class="akt-cal__day${!inMonth ? ' is-out' : ''}">
          <div class="akt-cal__num">${Number(day.slice(8))}</div>
          ${dayItems
              .slice(0, 3)
              .map(
                  (it) =>
                      `<button type="button" class="akt-cal__chip akt-cal__chip--${esc(String(it.status || 'beantragt'))}" data-akt-detail="${esc(
                          it.aktivitaetId || it.itemId
                      )}" title="${esc(it.titel)}">${esc(it.titel)}</button>`
              )
              .join('')}
          ${dayItems.length > 3 ? `<span class="akt-cal__more">+${dayItems.length - 3}</span>` : ''}
        </div>`;
        })
        .join('');
    return `
    <section class="akt-panel">
      <div class="akt-cal__head">
        <h2>Kalender</h2>
        <div class="akt-top__row">
          <button type="button" class="btn" id="aktCalPrev"><i class="bi bi-chevron-left"></i></button>
          <strong>${esc(monthName)}</strong>
          <button type="button" class="btn" id="aktCalNext"><i class="bi bi-chevron-right"></i></button>
        </div>
      </div>
      ${filterBar(state, { scopeAll: true })}
      <div class="akt-cal__weekdays"><span>Mo</span><span>Di</span><span>Mi</span><span>Do</span><span>Fr</span><span>Sa</span><span>So</span></div>
      <div class="akt-cal__grid">${cells}</div>
    </section>`;
}

function renderAntrag(state) {
    const f = state.form || emptyForm();
    const sd = state.stammdaten;
    const rules = state.rules || DEFAULT_RULES;
    const preview = validateAktivitaet({
        draft: f,
        existing: state.items,
        rules: {
            minVorlaufTage: rules.minVorlaufTage,
            maxGleichzeitigProKlasse: rules.maxGleichzeitigProKlasse
        }
    });
    return `
    <section class="akt-panel">
      <h2>${state.editingItemId ? 'Antrag bearbeiten' : 'Neuer Antrag'}</h2>
      <form id="aktForm" class="akt-form" autocomplete="off">
        <div class="akt-form-grid">
          <label>Titel<input name="titel" required value="${esc(f.titel)}"></label>
          <label>Typ<select name="typ">${optionList(typChoices(), f.typ, '–')}</select></label>
          <label>Klasse<select name="klasseCode">${optionList(sd.classes, f.klasseCode, '– wählen –')}</select></label>
          <label>Lehrkraft<select name="lehrerCode">${optionList(sd.teachers, f.lehrerCode, '– wählen –')}</select></label>
          <label>Ort / Ziel<input name="ort" value="${esc(f.ort)}"></label>
          <label>Verkehrsmittel<input name="verkehrsmittel" value="${esc(f.verkehrsmittel)}"></label>
          <label>Startdatum<input type="date" name="startdatum" required value="${esc(f.startdatum)}"></label>
          <label>Enddatum<input type="date" name="enddatum" value="${esc(f.enddatum || f.startdatum)}"></label>
          <label>Startzeit<input name="startZeit" placeholder="08:00" value="${esc(f.startZeit)}"></label>
          <label>Endzeit<input name="endZeit" placeholder="16:00" value="${esc(f.endZeit)}"></label>
          <label class="akt-span-2">Begleitung<textarea name="begleitung" rows="2">${esc(f.begleitung)}</textarea></label>
          <label class="akt-span-2">Kostenhinweis<input name="kostenHinweis" value="${esc(f.kostenHinweis)}"></label>
          <label class="akt-span-2">Notiz / Begründung<textarea name="notiz" rows="3">${esc(f.notiz)}</textarea></label>
        </div>
        ${
            preview.errors.length
                ? `<div class="akt-alert akt-alert--bad"><ul>${preview.errors.map((e) => `<li>${esc(e)}</li>`).join('')}</ul></div>`
                : ''
        }
        ${
            preview.warnings.length
                ? `<div class="akt-alert akt-alert--info"><ul>${preview.warnings.map((e) => `<li>${esc(e)}</li>`).join('')}</ul></div>`
                : ''
        }
        <div class="akt-actions">
          <button type="submit" class="btn btn-success" ${preview.errors.length ? 'disabled' : ''}><i class="bi bi-save"></i>Speichern</button>
          <button type="button" class="btn" id="aktFormReset">Zurücksetzen</button>
        </div>
      </form>
    </section>`;
}

function renderFreigabe(state) {
    if (!canDecide(state)) return '<p class="muted">Nur für Admin/Direktion.</p>';
    const open = filterItems(state.items, { status: 'beantragt' }, { role: 'admin', onlyOpen: true });
    const labels = labelMaps(state.stammdaten);
    return `
    <section class="akt-panel">
      <h2>Freigabe</h2>
      <p class="muted">Offene Anträge prüfen, genehmigen oder ablehnen.</p>
      ${
          open.length
              ? `<div class="akt-card-list">${open
                    .map(
                        (it) => `
        <article class="akt-card">
          <header>
            <strong>${esc(it.titel)}</strong>
            ${statusBadge(it.status)}
          </header>
          <p>${esc(typLabel(it.typ))} · ${esc(labels.klasse[it.klasseCode] || it.klasseCode)} · ${esc(
                            formatDeDate(it.startdatum)
                        )}${it.enddatum && it.enddatum !== it.startdatum ? ' – ' + esc(formatDeDate(it.enddatum)) : ''}</p>
          <p class="muted">${esc(it.ort || '')}${it.notiz ? ' — ' + esc(it.notiz) : ''}</p>
          <div class="akt-actions">
            <button type="button" class="btn btn-success" data-akt-approve="${esc(it.itemId)}"><i class="bi bi-check-lg"></i>Genehmigen</button>
            <button type="button" class="btn" data-akt-reject="${esc(it.itemId)}"><i class="bi bi-x-lg"></i>Ablehnen</button>
            <button type="button" class="btn" data-akt-detail="${esc(it.aktivitaetId || it.itemId)}">Details</button>
          </div>
        </article>`
                    )
                    .join('')}</div>`
              : '<p class="muted">Keine offenen Anträge.</p>'
      }
    </section>`;
}

function renderRegeln(state) {
    const r = state.rules || { ...DEFAULT_RULES, itemId: '', title: '' };
    return `
    <section class="akt-panel">
      <h2>Regelwerk</h2>
      <form id="aktRulesForm" class="akt-form">
        <div class="akt-form-grid">
          <label>Min. Vorlauf (Tage)<input type="number" min="0" name="minVorlaufTage" value="${esc(r.minVorlaufTage)}"></label>
          <label>Max. gleichzeitig pro Klasse<input type="number" min="1" name="maxGleichzeitigProKlasse" value="${esc(
              r.maxGleichzeitigProKlasse
          )}"></label>
        </div>
        <div class="akt-actions">
          <button type="submit" class="btn btn-success" ${!r.itemId && !state.localDemoOnly ? 'disabled' : ''}><i class="bi bi-save"></i>Speichern</button>
        </div>
        ${r.itemId || state.localDemoOnly ? '' : '<p class="muted">Kein Regelwerk-Eintrag geladen – Listen-Setup ausführen.</p>'}
      </form>
    </section>
    <section class="akt-panel">
      <h3>Demo-Daten (Schuljahr 2026/27)</h3>
      <p class="muted" style="margin:0 0 10px">Ca. 36 Exkursionen/Aktivitäten zum Vorzeigen. Einspielen überschreibt bestehende Demo-IDs; Reset löscht nur Demo-Zeilen.</p>
      <div class="akt-actions">
        <button type="button" class="btn" id="aktBtnDemoPanel"><i class="bi bi-database"></i>Demo laden</button>
        <button type="button" class="btn" id="aktBtnDemoResetPanel"><i class="bi bi-trash"></i>Demo zurücksetzen</button>
        <label class="btn" for="aktImportJson"><i class="bi bi-upload"></i>JSON importieren</label>
        <input type="file" id="aktImportJson" accept=".json,application/json" hidden>
      </div>
      <p class="muted" style="margin:10px 0 0;font-size:0.88em;">Datei: <code>docs/demo-data/schulaktivitaeten-2026-27.json</code></p>
    </section>`;
}

function renderDetailModal(state) {
    const id = state.detailId;
    const it = state.items.find((x) => x.aktivitaetId === id || x.itemId === id);
    if (!it) return '';
    const labels = labelMaps(state.stammdaten);
    const editable = canEditItem(it, state);
    return `
    <div class="akt-modal" role="dialog" aria-modal="true">
      <div class="akt-modal__card">
        <header class="akt-modal__head">
          <h3>${esc(it.titel)}</h3>
          <button type="button" class="btn" id="aktDetailClose" aria-label="Schließen"><i class="bi bi-x-lg"></i></button>
        </header>
        <div class="akt-modal__body">
          <p>${statusBadge(it.status)} · ${esc(typLabel(it.typ))}</p>
          <dl class="akt-dl">
            <dt>Zeitraum</dt><dd>${esc(formatDeDate(it.startdatum))}${
                it.enddatum && it.enddatum !== it.startdatum ? ' – ' + esc(formatDeDate(it.enddatum)) : ''
            }${it.startZeit ? ' · ' + esc(it.startZeit) : ''}${it.endZeit ? '–' + esc(it.endZeit) : ''}</dd>
            <dt>Klasse</dt><dd>${esc(labels.klasse[it.klasseCode] || it.klasseCode)}</dd>
            <dt>Lehrkraft</dt><dd>${esc(labels.lehrer[it.lehrerCode] || it.lehrerCode || it.lehrerEmail)}</dd>
            <dt>Ort</dt><dd>${esc(it.ort || '–')}</dd>
            <dt>Begleitung</dt><dd>${esc(it.begleitung || '–')}</dd>
            <dt>Notiz</dt><dd>${esc(it.notiz || '–')}</dd>
            ${it.ablehnungsGrund ? `<dt>Ablehnung</dt><dd>${esc(it.ablehnungsGrund)}</dd>` : ''}
          </dl>
        </div>
        <footer class="akt-actions">
          ${
              editable
                  ? `<button type="button" class="btn" data-akt-edit="${esc(it.itemId)}"><i class="bi bi-pencil"></i>Bearbeiten</button>
                     <button type="button" class="btn" data-akt-delete="${esc(it.itemId)}"><i class="bi bi-trash"></i>Löschen</button>`
                  : ''
          }
          ${
              canDecide(state) && String(it.status).toLowerCase() === 'beantragt'
                  ? `<button type="button" class="btn btn-success" data-akt-approve="${esc(it.itemId)}">Genehmigen</button>
                     <button type="button" class="btn" data-akt-reject="${esc(it.itemId)}">Ablehnen</button>`
                  : ''
          }
        </footer>
      </div>
    </div>`;
}

export function readFormFromDom() {
    const form = document.getElementById('aktForm');
    if (!form) return emptyForm();
    const fd = new FormData(form);
    const get = (k) => String(fd.get(k) || '').trim();
    return {
        titel: get('titel'),
        typ: get('typ') || 'Exkursion',
        klasseCode: get('klasseCode'),
        lehrerCode: get('lehrerCode'),
        ort: get('ort'),
        startdatum: get('startdatum'),
        enddatum: get('enddatum') || get('startdatum'),
        startZeit: get('startZeit'),
        endZeit: get('endZeit'),
        begleitung: get('begleitung'),
        verkehrsmittel: get('verkehrsmittel'),
        kostenHinweis: get('kostenHinweis'),
        notiz: get('notiz'),
        status: 'beantragt'
    };
}

export function readFiltersFromDom() {
    const val = (id) => {
        const el = document.getElementById(id);
        return el ? String(el.value || '').trim() : '';
    };
    return {
        klasse: val('aktFilterKlasse'),
        typ: val('aktFilterTyp'),
        status: val('aktFilterStatus'),
        lehrer: val('aktFilterLehrer')
    };
}

export function readRulesForm() {
    const form = document.getElementById('aktRulesForm');
    if (!form) return null;
    const fd = new FormData(form);
    return {
        minVorlaufTage: Number(fd.get('minVorlaufTage')) || 7,
        maxGleichzeitigProKlasse: Number(fd.get('maxGleichzeitigProKlasse')) || 1
    };
}
