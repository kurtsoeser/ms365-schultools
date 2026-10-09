/**
 * UI: Freistellungen Lehrkräfte.
 */
import { viewsForRole, scopeFromState, canDecide } from './lfr-state.js';
import { KATEGORIE_CHOICES as KAT } from './lfr-schema.js';
import {
    filterItems,
    formatDeDateTimeRange,
    monthGridDates,
    itemCoversDay,
    normalizeStatus
} from './lfr-logic.js';

function esc(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function statusBadge(status) {
    const s = normalizeStatus(status);
    const cls =
        s === 'Genehmigt' ? 'fr-badge--ok' : s === 'Abgelehnt' ? 'fr-badge--bad' : 'fr-badge--warn';
    return `<span class="fr-badge ${cls}">${esc(s)}</span>`;
}

function roleSubtitle(role) {
    return role === 'direktion' ? 'Verwaltung' : 'Lehrkraft';
}

function topTitle(state) {
    if (state.view === 'einrichtung') return 'Administration';
    return state.role === 'direktion'
        ? 'Freistellungen Lehrkräfte – Verwaltung'
        : 'Freistellungen Lehrkräfte';
}

function renderNavSession(state) {
    return `
        <div class="fr-nav__session" aria-label="Anmeldung und Rolle">
          <span class="fr-nav__role-label">Anmeldung</span>
          ${
              state.accountEmail
                  ? `<p class="fr-nav__account-name">${esc(state.accountName || state.accountEmail)}</p>`
                  : '<p class="fr-nav__hint">Mit Schul-Konto anmelden (Menü oben rechts).</p>'
          }
          <div class="fr-nav__session-role">
            <span class="fr-nav__session-role-label">Rolle wählen</span>
            <div class="fr-nav__role-btns" role="group" aria-label="Rolle wählen">
              <button type="button" class="fr-chip${state.role === 'lehrer' ? ' is-active' : ''}" data-lfr-role="lehrer">Lehrkraft</button>
              <button type="button" class="fr-chip${state.role === 'direktion' ? ' is-active' : ''}" data-lfr-role="direktion">Direktion</button>
            </div>
          </div>
          <p class="fr-nav__hint">Genehmigung über Microsoft Approvals (Power Automate).</p>
        </div>`;
}

function renderMainTop(state) {
    const onAdmin = state.view === 'einrichtung';
    return `
        <header class="fr-top fr-top--student">
          <div>
            <h2 class="fr-top__student-title">${esc(topTitle(state))}</h2>
            <p class="fr-nav__hint">${
                onAdmin
                    ? 'SharePoint-Verbindung und Links zur IT-Einrichtung.'
                    : state.role === 'direktion'
                      ? 'Übersicht, Anträge und Freigabe. Technik: Menü „Einrichtung“.'
                      : 'Anträge stellen und Status verfolgen.'
            }</p>
          </div>
          ${
              state.accountEmail && !onAdmin
                  ? `<button type="button" class="btn" id="lfrBtnRefresh" title="Aktualisieren"><i class="bi bi-arrow-repeat"></i>Aktualisieren</button>`
                  : ''
          }
        </header>`;
}

function computeLfrKpis(items) {
    const now = new Date();
    const in14 = new Date(now.getTime());
    in14.setDate(in14.getDate() + 14);
    let ausstehend = 0;
    let genehmigt = 0;
    let mehrtage = 0;
    let demnaechst = 0;
    (items || []).forEach((it) => {
        const st = normalizeStatus(it.status);
        if (st === 'Ausstehend') ausstehend++;
        if (st === 'Genehmigt') genehmigt++;
        const b = it.beginn ? new Date(it.beginn) : null;
        const e = it.ende ? new Date(it.ende) : null;
        if (b && e && !isNaN(b) && !isNaN(e) && e.getTime() - b.getTime() > 86400000) mehrtage++;
        if (b && !isNaN(b) && b >= now && b <= in14) demnaechst++;
    });
    return { ausstehend, genehmigt, mehrtage, demnaechst };
}

function dayCount(beginn, ende) {
    const b = beginn ? new Date(beginn) : null;
    const e = ende ? new Date(ende) : null;
    if (!b || !e || isNaN(b) || isNaN(e)) return null;
    const ms = e.getTime() - b.getTime();
    return Math.max(1, Math.round(ms / 86400000) + 1);
}

/**
 * @param {object} state
 * @param {HTMLElement} root
 */
export function renderApp(state, root) {
    const views = viewsForRole(state.role);
    root.innerHTML = `
    <div class="fr-shell fr-shell--student">
      <aside class="fr-nav" aria-label="Navigation">
        <div class="fr-nav__brand">
          <i class="bi bi-briefcase" aria-hidden="true"></i>
          <div>
            <strong>Freistellungen</strong>
            <span>${esc(roleSubtitle(state.role))}</span>
          </div>
        </div>
        <nav class="fr-nav__list">
          ${views
              .map(
                  (v) =>
                      `<button type="button" class="fr-nav__btn${state.view === v.id ? ' is-active' : ''}" data-lfr-view="${esc(
                          v.id
                      )}"><i class="bi ${esc(v.icon)}" aria-hidden="true"></i>${esc(v.label)}</button>`
              )
              .join('')}
        </nav>
        ${renderNavSession(state)}
      </aside>
      <div class="fr-main">
        ${renderMainTop(state)}
        ${
            state.localDemoOnly
                ? '<div class="fr-alert fr-alert--info" role="status">Lokaler Demo-Modus – unter „Einrichtung“ oder IT-Setup SharePoint verbinden.</div>'
                : ''
        }
        ${state.error ? `<div class="fr-alert fr-alert--bad" role="alert">${esc(state.error)}</div>` : ''}
        ${state.loading ? '<div class="fr-skel" aria-busy="true"><div class="fr-skel__card"></div><div class="fr-skel__card"></div></div>' : ''}
        <div class="fr-view${state.loading ? ' is-loading' : ''}">
          ${state.loading ? '' : renderView(state)}
        </div>
      </div>
    </div>`;
}

function renderView(state) {
    switch (state.view) {
        case 'meine':
            return renderListe(state, { mine: true });
        case 'antrag':
            return renderAntrag(state);
        case 'liste':
            return renderListe(state, { mine: false });
        case 'freigabe':
            return renderFreigabe(state);
        case 'kalender':
            return renderKalender(state);
        case 'export':
            return renderExport(state);
        case 'einrichtung':
            return renderEinrichtung(state);
        default:
            return renderDashboard(state);
    }
}

function renderDashboard(state) {
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, {}, scope);
    const kpi = computeLfrKpis(items);
    const offen = items.filter((i) => normalizeStatus(i.status) === 'Ausstehend');
    return `
    <section class="fr-panel">
      <h2>Übersicht</h2>
      <p class="muted">Anträge in der SharePoint-Liste; Power Automate startet Microsoft Approvals (Direktion).</p>
      <div class="fr-kpi-row">
        <div class="fr-kpi"><strong>${kpi.ausstehend}</strong><span>Ausstehend</span></div>
        <div class="fr-kpi"><strong>${kpi.genehmigt}</strong><span>Genehmigt</span></div>
        <div class="fr-kpi"><strong>${kpi.mehrtage}</strong><span>Mehrtägig</span></div>
        <div class="fr-kpi"><strong>${kpi.demnaechst}</strong><span>In 14 Tagen</span></div>
      </div>
    </section>
    <section class="fr-panel">
      <h3>Offene Anträge</h3>
      ${offen.length ? renderTable(offen) : '<p class="muted">Keine offenen Anträge.</p>'}
      <div class="fr-actions" style="margin-top:12px">
        <button type="button" class="btn btn-success" data-lfr-view-jump="antrag"><i class="bi bi-plus-lg"></i>Neuer Antrag</button>
        <button type="button" class="btn" data-lfr-view-jump="kalender"><i class="bi bi-calendar3"></i>Kalender</button>
        ${canDecide(state) ? '<button type="button" class="btn" data-lfr-view-jump="freigabe"><i class="bi bi-check2-square"></i>Offene Genehmigungen</button>' : ''}
      </div>
    </section>`;
}

function filterBar(state) {
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const count = filterItems(state.items, state.filters, scope).length;
    const katOpts = KAT.map(
        (k) => `<option value="${esc(k)}"${state.filters.kategorie === k ? ' selected' : ''}>${esc(k)}</option>`
    ).join('');
    return `
    <div class="fr-filters">
      <select id="lfrFilterStatus" aria-label="Status">
        <option value="">Alle Status</option>
        <option value="Ausstehend"${state.filters.status === 'Ausstehend' ? ' selected' : ''}>Ausstehend</option>
        <option value="Genehmigt"${state.filters.status === 'Genehmigt' ? ' selected' : ''}>Genehmigt</option>
        <option value="Abgelehnt"${state.filters.status === 'Abgelehnt' ? ' selected' : ''}>Abgelehnt</option>
      </select>
      <select id="lfrFilterKat" aria-label="Kategorie">
        <option value="">Alle Kategorien</option>${katOpts}
      </select>
      <input type="search" id="lfrFilterQ" placeholder="Suchen …" value="${esc(state.filters.q || '')}">
      <span class="fr-filters__meta muted">${count} Einträge</span>
    </div>`;
}

function renderListe(state, opts) {
    const mine = opts && opts.mine;
    const scope = mine
        ? { scopeAll: false, accountEmail: state.accountEmail }
        : scopeFromState(state, { scopeAll: true });
    const items = filterItems(state.items, state.filters, scope);
    return `
    <section class="fr-panel">
      <h2>${mine ? 'Meine Anträge' : 'Alle Anträge'}</h2>
      ${filterBar(state)}
      ${items.length ? renderTable(items) : '<p class="muted">Keine Einträge.</p>'}
    </section>`;
}

function renderTable(items) {
    const rows = items
        .map((it) => {
            const days = dayCount(it.beginn, it.ende);
            const id = it.antragId || it.itemId;
            return `<tr>
          <td><button type="button" class="fr-link" data-lfr-detail="${esc(id)}">${esc(it.lehrerName || it.titel || it.lehrerEmail)}</button></td>
          <td>${esc(formatDeDateTimeRange(it.beginn, it.ende))}</td>
          <td>${days != null ? days : '–'}${days != null && days > 1 ? ' <span class="fr-badge fr-badge--info">mehr</span>' : ''}</td>
          <td>${statusBadge(it.status)}</td>
          <td>${esc(it.kategorie)}</td>
          <td>${esc(it.titel)}</td>
          <td class="fr-td-actions">
            <button type="button" class="btn btn-sm" data-lfr-detail="${esc(id)}" title="Details"><i class="bi bi-eye"></i></button>
          </td>
        </tr>`;
        })
        .join('');
    return `
    <div class="fr-table-wrap">
      <table class="fr-table">
        <thead>
          <tr>
            <th>Lehrkraft</th><th>Zeitraum</th><th>Tage</th><th>Status</th><th>Kategorie</th><th>Titel</th><th></th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

function renderKalender(state) {
    const y = state.calYear;
    const m = state.calMonth;
    const days = monthGridDates(y, m);
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, state.filters, scope).filter(
        (i) => normalizeStatus(i.status) !== 'Abgelehnt'
    );
    const monthName = new Date(y, m - 1, 1).toLocaleDateString('de-AT', { month: 'long', year: 'numeric' });
    const cells = days
        .map((day) => {
            const inMonth = Number(day.slice(5, 7)) === m;
            const dayItems = items.filter((it) => itemCoversDay(it, day));
            return `<div class="fr-cal__day${!inMonth ? ' is-out' : ''}">
          <div class="fr-cal__num">${Number(day.slice(8))}</div>
          ${dayItems
              .slice(0, 4)
              .map(
                  (it) => {
                      const c =
                          normalizeStatus(it.status) === 'Genehmigt'
                              ? 'fr-cal__chip--ok'
                              : normalizeStatus(it.status) === 'Abgelehnt'
                                ? 'fr-cal__chip--bad'
                                : 'fr-cal__chip--warn';
                      return `<button type="button" class="fr-cal__chip ${c}" data-lfr-detail="${esc(
                          it.antragId || it.itemId
                      )}" title="${esc(it.titel)}">${esc(it.lehrerName || it.titel)}</button>`;
                  }
              )
              .join('')}
        </div>`;
        })
        .join('');
    return `
    <section class="fr-panel">
      <div class="fr-cal__head">
        <h2>Kalender</h2>
        <div class="fr-cal__nav">
          <button type="button" class="btn" id="lfrCalPrev"><i class="bi bi-chevron-left"></i></button>
          <strong>${esc(monthName)}</strong>
          <button type="button" class="btn" id="lfrCalNext"><i class="bi bi-chevron-right"></i></button>
        </div>
      </div>
      ${filterBar(state)}
      <div class="fr-cal__weekdays"><span>Mo</span><span>Di</span><span>Mi</span><span>Do</span><span>Fr</span><span>Sa</span><span>So</span></div>
      <div class="fr-cal__grid">${cells}</div>
    </section>`;
}

function renderAntrag(state) {
    const f = state.form;
    const katOpts = KAT.map(
        (k) => `<option value="${esc(k)}"${f.kategorie === k ? ' selected' : ''}>${esc(k)}</option>`
    ).join('');
    return `
    <section class="fr-panel">
      <h2>Freistellung beantragen</h2>
      <p class="muted">Nach dem Speichern startet (mit Flow) die Genehmigung durch die Direktion – Status erscheint hier und in der Liste.</p>
      <div class="fr-form-grid">
        <label>Titel<input type="text" id="lfrFormTitel" value="${esc(f.titel)}" maxlength="200"></label>
        <label>Beginn<input type="datetime-local" id="lfrFormBeginn" value="${esc(f.beginn)}"></label>
        <label>Ende<input type="datetime-local" id="lfrFormEnde" value="${esc(f.ende)}"></label>
        <label>Kategorie<select id="lfrFormKat">${katOpts}</select></label>
        <label class="fr-form-span2">Beschreibung<textarea id="lfrFormBesch" rows="4">${esc(f.beschreibung)}</textarea></label>
        <label>Lehrkraft<input type="text" id="lfrFormName" value="${esc(f.lehrerName)}" readonly></label>
        <label>E-Mail<input type="email" id="lfrFormMail" value="${esc(f.lehrerEmail)}" readonly></label>
      </div>
      <div class="fr-actions">
        <button type="button" class="btn btn-success" id="lfrFormSubmit"><i class="bi bi-send"></i>Antrag absenden</button>
        <button type="button" class="btn" id="lfrFormReset">Zurücksetzen</button>
      </div>
    </section>`;
}

function renderFreigabe(state) {
    const items = filterItems(state.items, { status: 'Ausstehend' }, { scopeAll: true });
    return `
    <section class="fr-panel">
      <h2>Freigabe (Direktion)</h2>
      <p class="muted">Manuelle Freigabe für Tests ohne Flow. In Produktion übernimmt Power Automate Approvals und schreibt Status zurück.</p>
      ${items.length ? renderFreigabeTable(items) : '<p class="muted">Keine offenen Anträge.</p>'}
    </section>`;
}

function renderFreigabeTable(items) {
    return `
    <div class="fr-table-wrap">
      <table class="fr-table">
        <thead><tr><th>Zeitraum</th><th>Titel</th><th>Lehrkraft</th><th></th></tr></thead>
        <tbody>
          ${items
              .map(
                  (it) => `<tr>
              <td>${esc(formatDeDateTimeRange(it.beginn, it.ende))}</td>
              <td>${esc(it.titel)}</td>
              <td>${esc(it.lehrerName || it.lehrerEmail)}</td>
              <td class="fr-td-actions">
                <button type="button" class="btn btn-sm btn-success" data-lfr-approve="${esc(it.antragId || it.itemId)}">Genehmigen</button>
                <button type="button" class="btn btn-sm alt" data-lfr-reject="${esc(it.antragId || it.itemId)}">Ablehnen</button>
              </td>
            </tr>`
              )
              .join('')}
        </tbody>
      </table>
    </div>`;
}

function renderExport(state) {
    const approved = filterItems(state.items, { status: 'Genehmigt' }, { scopeAll: true });
    const calUser = esc(state.outlookCalendarUser || '');
    const calHint = calUser
        ? `Zielkalender: <code>${calUser}</code>${state.outlookCalendarId ? ' (eigene Kalender-ID)' : ''}`
        : 'Kalender-Ziel noch nicht konfiguriert – im <a href="lehrer-freistellung-setup.html">IT-Setup</a> Schritt 4.';
    return `
    <section class="fr-panel">
      <h2>Kalender-Export</h2>
      <p class="muted">Genehmigte Freistellungen als Datei oder per Microsoft Graph in einen freigegebenen Outlook-Kalender.</p>
      <p><strong>${approved.length}</strong> genehmigte Einträge.</p>
      <div class="fr-actions">
        <button type="button" class="btn btn-success" id="lfrBtnIcs"><i class="bi bi-download"></i>iCal herunterladen</button>
        <button type="button" class="btn" id="lfrBtnOutlookSync"><i class="bi bi-calendar-plus"></i>In Outlook-Kalender synchronisieren</button>
      </div>
      <p class="muted" style="margin-top:14px;line-height:1.45;">${calHint}</p>
      <p class="muted" style="font-size:0.9em;">Graph-Berechtigung: Calendars.ReadWrite (angemeldetes Konto mit Zugriff auf den Zielkalender).</p>
    </section>`;
}

function renderEinrichtung(state) {
    return `
    <section class="fr-panel">
      <h2>Einrichtung</h2>
      <p class="muted">SharePoint-Liste für Lehrer-Freistellungen (getrennt von der Schüler-Liste). Power-Automate-Flow: ein Genehmigungsschritt Direktion.</p>
      <nav class="fr-admin-nav" aria-label="Administration">
        <a class="fr-admin-nav__item" href="lehrer-freistellung-setup.html"><i class="bi bi-wrench" aria-hidden="true"></i><span>IT-Setup</span></a>
        <a class="fr-admin-nav__item" href="freistellung-planer.html"><i class="bi bi-mortarboard" aria-hidden="true"></i><span>Schüler-Planer</span></a>
        <a class="fr-admin-nav__item" href="../tenant.html"><i class="bi bi-gear" aria-hidden="true"></i><span>Stammdaten</span></a>
        <a class="fr-admin-nav__item" href="../index.html"><i class="bi bi-arrow-left" aria-hidden="true"></i><span>Dashboard</span></a>
      </nav>
      <div class="fr-form-grid" style="margin-top:14px">
        <label class="fr-form-span2">SharePoint-Site-URL
          <input type="url" id="lfrSetupSite" value="${esc(state.siteUrl)}" placeholder="https://…sharepoint.com/sites/…">
        </label>
        <label>Listentitel
          <input type="text" id="lfrSetupListName" value="${esc(state.listName)}">
        </label>
        <label>Listen-ID (optional)
          <input type="text" id="lfrSetupListId" value="${esc(state.listId)}">
        </label>
      </div>
      <div class="fr-actions">
        <button type="button" class="btn" id="lfrBtnSaveSetup"><i class="bi bi-save"></i>Speichern</button>
        <button type="button" class="btn btn-success" id="lfrBtnEnsureList"><i class="bi bi-list-plus"></i>Liste anlegen / prüfen</button>
        <button type="button" class="btn" id="lfrBtnReload"><i class="bi bi-arrow-clockwise"></i>Daten laden</button>
      </div>
      <p class="muted" style="margin-top:12px;">
        Vollständige IT-Einrichtung (Liste, Berechtigungen, Flow, Kalender): <a href="lehrer-freistellung-setup.html">Lehrer-Freistellungen Setup</a>.
        Schüler: <a href="freistellung-setup.html">Freistellungen Setup</a> · <a href="freistellung-planer.html">Schüler-Planer</a>
      </p>
    </section>`;
}

export function readFormFromDom() {
    const $ = (id) => document.getElementById(id);
    return {
        titel: String(($('lfrFormTitel') && $('lfrFormTitel').value) || '').trim(),
        beginn: String(($('lfrFormBeginn') && $('lfrFormBeginn').value) || '').trim(),
        ende: String(($('lfrFormEnde') && $('lfrFormEnde').value) || '').trim(),
        kategorie: String(($('lfrFormKat') && $('lfrFormKat').value) || '').trim(),
        beschreibung: String(($('lfrFormBesch') && $('lfrFormBesch').value) || '').trim(),
        lehrerName: String(($('lfrFormName') && $('lfrFormName').value) || '').trim(),
        lehrerEmail: String(($('lfrFormMail') && $('lfrFormMail').value) || '').trim()
    };
}

export function readFiltersFromDom(state) {
    const $ = (id) => document.getElementById(id);
    return {
        ...state.filters,
        status: String(($('lfrFilterStatus') && $('lfrFilterStatus').value) || '').trim(),
        kategorie: String(($('lfrFilterKat') && $('lfrFilterKat').value) || '').trim(),
        q: String(($('lfrFilterQ') && $('lfrFilterQ').value) || '').trim()
    };
}

export function readSetupFromDom(state) {
    const $ = (id) => document.getElementById(id);
    return {
        siteUrl: String(($('lfrSetupSite') && $('lfrSetupSite').value) || '').trim(),
        listName: String(($('lfrSetupListName') && $('lfrSetupListName').value) || '').trim(),
        listId: String(($('lfrSetupListId') && $('lfrSetupListId').value) || '').trim()
    };
}
