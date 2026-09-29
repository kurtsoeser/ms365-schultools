/**
 * UI-Rendering Freistellungs-Planer.
 */
import {
    computeDashboardKpis,
    validateFreistellung,
    approvalPath,
    inclusiveDayCount,
    STATUS_CHOICES,
    KATEGORIE_CHOICES
} from './freistellung-planer-logic.js';
import {
    viewsForRole,
    filterItems,
    formatDeDate,
    statusLabel,
    roleLabel,
    scopeFromState,
    canDecide,
    resolveKvForClass
} from './freistellung-planer-state.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function statusBadge(status) {
    const s = String(status || 'Ausstehend').trim();
    const cls =
        s === 'Genehmigt'
            ? 'fr-badge fr-badge--ok'
            : s === 'Abgelehnt'
              ? 'fr-badge fr-badge--bad'
              : 'fr-badge fr-badge--warn';
    return `<span class="${cls}">${esc(statusLabel(s))}</span>`;
}

function optionList(items, selected, emptyLabel, valueKey, labelKey) {
    const vk = valueKey || 'code';
    const lk = labelKey || 'name';
    const opts = [`<option value="">${esc(emptyLabel)}</option>`];
    (items || []).forEach((it) => {
        const v = typeof it === 'string' ? it : it[vk] || it.value;
        const label = typeof it === 'string' ? it : it[lk] || it.label || v;
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
    <div class="fr-shell">
      <aside class="fr-nav" aria-label="Navigation">
        <div class="fr-nav__brand">
          <i class="bi bi-calendar2-check" aria-hidden="true"></i>
          <div>
            <strong>Freistellungen</strong>
            <span>Antrag &amp; Genehmigung</span>
          </div>
        </div>
        <nav class="fr-nav__list">
          ${views
              .map(
                  (v) => `
            <button type="button" class="fr-nav__btn${state.view === v.id ? ' is-active' : ''}" data-fr-view="${esc(v.id)}">
              <i class="bi ${v.icon}" aria-hidden="true"></i>${esc(v.label)}
            </button>`
              )
              .join('')}
        </nav>
        <div class="fr-nav__role">
          <span class="fr-nav__role-label">Rolle (Demo)</span>
          <div class="fr-nav__role-btns">
            <button type="button" class="fr-chip${state.role === 'schueler' ? ' is-active' : ''}" data-fr-role="schueler">Schüler</button>
            <button type="button" class="fr-chip${state.role === 'kv' ? ' is-active' : ''}" data-fr-role="kv">KV</button>
            <button type="button" class="fr-chip${state.role === 'direktion' ? ' is-active' : ''}" data-fr-role="direktion">Direktion</button>
          </div>
          <p class="fr-nav__hint">Genehmigung läuft über Microsoft Approvals (Power Automate). Später Entra-Rollen.</p>
        </div>
      </aside>
      <div class="fr-main">
        <header class="fr-top">
          <div class="fr-top__site">
            <label for="frSiteUrl">SharePoint-Site</label>
            <div class="fr-top__row">
              <input type="url" id="frSiteUrl" value="${esc(state.siteUrl)}" placeholder="https://…sharepoint.com/sites/Administration" spellcheck="false">
              <button type="button" class="btn" id="frBtnLoad"><i class="bi bi-arrow-repeat"></i>Laden</button>
              <a class="btn" href="freistellung-setup.html" style="text-decoration:none"><i class="bi bi-gear"></i>Setup</a>
              <button type="button" class="btn" id="frBtnCsv"><i class="bi bi-download"></i>CSV</button>
              <button type="button" class="btn" id="frBtnDemo" title="SJ 2026/27 Demo laden (lokal, optional SharePoint)"><i class="bi bi-database"></i>Demo</button>
              <button type="button" class="btn" id="frBtnDemoReset" title="Demo-Einträge löschen / zurücksetzen"><i class="bi bi-trash"></i>Demo reset</button>
            </div>
          </div>
          <div class="fr-top__user">
            <span>Rolle <span class="fr-badge fr-badge--info">${esc(roleLabel(state.role))}</span></span>
            ${state.localDemoOnly ? '<span class="fr-badge fr-badge--warn">Demo lokal</span>' : ''}
            ${
                state.accountEmail
                    ? `<small>${esc(state.accountName || state.accountEmail)}</small>`
                    : '<small class="muted">Bitte oben rechts anmelden</small>'
            }
          </div>
        </header>
        ${
            state.localDemoOnly
                ? '<div class="fr-alert fr-alert--info" role="status">Lokaler Demo-Modus – Anzeige ohne SharePoint. Mit „Laden“ echte Liste verbinden.</div>'
                : ''
        }
        ${
            state.info
                ? `<div class="fr-alert fr-alert--info" role="status">${esc(state.info)}</div>`
                : ''
        }
        ${state.error ? `<div class="fr-alert fr-alert--bad" role="alert">${esc(state.error)}</div>` : ''}
        ${state.loading ? '<div class="fr-skel" aria-busy="true"><div class="fr-skel__card"></div><div class="fr-skel__card"></div></div>' : ''}
        <div id="frView" class="fr-view${state.loading ? ' is-loading' : ''}">
          ${state.loading ? '' : renderView(state)}
        </div>
      </div>
    </div>
    ${state.detailId ? renderDetailModal(state) : ''}
  `;
}

function filterBar(state) {
    const sd = state.stammdaten;
    const scope = scopeFromState(state);
    const count = filterItems(state.items, state.filters, scope).length;
    return `
    <div class="fr-filters" data-fr-filters>
      <input type="search" id="frFilterQ" placeholder="Suche…" value="${esc(state.filters.q || '')}" aria-label="Suche">
      <select id="frFilterKlasse" aria-label="Klasse">${optionList(sd.classes, state.filters.klasse, 'Alle Klassen')}</select>
      <select id="frFilterStatus" aria-label="Status">
        <option value="">Alle Status</option>
        ${STATUS_CHOICES.map(
            (s) =>
                `<option value="${esc(s)}"${state.filters.status === s ? ' selected' : ''}>${esc(s)}</option>`
        ).join('')}
      </select>
      <select id="frFilterKat" aria-label="Kategorie">
        <option value="">Alle Kategorien</option>
        ${KATEGORIE_CHOICES.map(
            (k) =>
                `<option value="${esc(k)}"${state.filters.kategorie === k ? ' selected' : ''}>${esc(k)}</option>`
        ).join('')}
      </select>
      <select id="frFilterMulti" aria-label="Dauer">
        <option value="">Alle Dauern</option>
        <option value="0"${state.filters.multiDay === '0' ? ' selected' : ''}>1 Tag</option>
        <option value="1"${state.filters.multiDay === '1' ? ' selected' : ''}>Mehrtägig</option>
      </select>
      <button type="button" class="btn" id="frFilterReset"><i class="bi bi-arrow-counterclockwise"></i>Reset</button>
      <span class="fr-filters__meta">${count} Einträge</span>
    </div>`;
}

function renderView(state) {
    switch (state.view) {
        case 'liste':
            return renderListe(state);
        case 'antrag':
            return renderAntrag(state);
        case 'meine':
            return renderMeine(state);
        case 'freigabe':
            return renderFreigabe(state);
        case 'bericht':
            return renderBericht(state);
        default:
            return renderDashboard(state);
    }
}

function renderDashboard(state) {
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, {}, scope);
    const kpi = computeDashboardKpis(items);
    const offen = items.filter((i) => String(i.status) === 'Ausstehend');
    return `
    <section class="fr-panel">
      <h2>Übersicht</h2>
      <p class="muted">Anträge landen in der SharePoint-Liste; der Power-Automate-Flow startet Microsoft Approvals (KV, bei Mehrtagen sequentiell Direktion).</p>
      <div class="fr-kpi-row">
        <div class="fr-kpi"><strong>${kpi.ausstehend}</strong><span>Ausstehend</span></div>
        <div class="fr-kpi"><strong>${kpi.genehmigt}</strong><span>Genehmigt</span></div>
        <div class="fr-kpi"><strong>${kpi.mehrtage}</strong><span>Mehrtägig</span></div>
        <div class="fr-kpi"><strong>${kpi.demnaechst}</strong><span>In 14 Tagen</span></div>
      </div>
    </section>
    <section class="fr-panel">
      <h3>Offene Anträge</h3>
      ${offen.length ? renderTable(offen, state) : '<p class="muted">Keine offenen Anträge.</p>'}
      <div class="fr-actions" style="margin-top:12px">
        <button type="button" class="btn btn-success" data-fr-view-jump="antrag"><i class="bi bi-plus-lg"></i>Neuer Antrag</button>
        ${canDecide(state) ? '<button type="button" class="btn" data-fr-view-jump="freigabe"><i class="bi bi-check2-square"></i>Offene Genehmigungen</button>' : ''}
      </div>
    </section>`;
}

function renderListe(state) {
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, state.filters, scope);
    return `
    <section class="fr-panel">
      <h2>Liste</h2>
      ${filterBar(state)}
      ${items.length ? renderTable(items, state) : '<p class="muted">Keine Einträge.</p>'}
    </section>`;
}

function renderMeine(state) {
    const items = filterItems(state.items, state.filters, {
        onlyMine: true,
        accountEmail: state.accountEmail
    });
    return `
    <section class="fr-panel">
      <h2>Meine Anträge</h2>
      ${filterBar(state)}
      ${items.length ? renderTable(items, state) : '<p class="muted">Noch keine eigenen Anträge.</p>'}
    </section>`;
}

function renderFreigabe(state) {
    const scope =
        state.role === 'direktion'
            ? { accountEmail: state.accountEmail }
            : { onlyKv: true, accountEmail: state.accountEmail };
    const items = filterItems(state.items, { status: 'Ausstehend' }, scope);
    return `
    <section class="fr-panel">
      <h2>Offene Genehmigungen</h2>
      <p class="muted">Die eigentliche Freigabe erfolgt in <strong>Microsoft Approvals</strong> (Teams/Outlook). Hier sehen Sie den Listenstand und können bei Bedarf den Status manuell nachziehen (Fallback ohne Flow).</p>
      ${items.length ? renderTable(items, state, { showDecide: true }) : '<p class="muted">Keine ausstehenden Anträge.</p>'}
    </section>`;
}

function renderBericht(state) {
    const items = filterItems(state.items, state.filters, {});
    const kpi = computeDashboardKpis(items);
    const byKlasse = {};
    items.forEach((it) => {
        const k = it.klasse || '–';
        byKlasse[k] = (byKlasse[k] || 0) + 1;
    });
    const rows = Object.keys(byKlasse)
        .sort()
        .map((k) => `<tr><td>${esc(k)}</td><td>${byKlasse[k]}</td></tr>`)
        .join('');
    return `
    <section class="fr-panel">
      <h2>Berichte</h2>
      ${filterBar(state)}
      <div class="fr-kpi-row">
        <div class="fr-kpi"><strong>${kpi.gesamt}</strong><span>Gesamt</span></div>
        <div class="fr-kpi"><strong>${kpi.genehmigt}</strong><span>Genehmigt</span></div>
        <div class="fr-kpi"><strong>${kpi.abgelehnt}</strong><span>Abgelehnt</span></div>
        <div class="fr-kpi"><strong>${kpi.mehrtage}</strong><span>Mehrtägig</span></div>
      </div>
      <h3>Nach Klasse</h3>
      <table class="fr-table">
        <thead><tr><th>Klasse</th><th>Anzahl</th></tr></thead>
        <tbody>${rows || '<tr><td colspan="2" class="muted">Keine Daten</td></tr>'}</tbody>
      </table>
      <div class="fr-actions" style="margin-top:12px">
        <button type="button" class="btn" id="frBtnCsvInline"><i class="bi bi-download"></i>CSV exportieren</button>
      </div>
    </section>`;
}

function renderAntrag(state) {
    const f = state.form || {};
    const path = approvalPath(f.beginn, f.ende);
    const check = validateFreistellung({ draft: f });
    const classOpts = optionList(state.stammdaten.classes, f.klasse, 'Klasse wählen…');
    return `
    <section class="fr-panel">
      <h2>Neuer Freistellungsantrag</h2>
      <p class="muted">Nach dem Speichern schreibt die App in die SharePoint-Liste (Status <em>Ausstehend</em>). Der eingerichtete Flow startet dann den Genehmigungsprozess.</p>
      <form id="frAntragForm" class="fr-form" autocomplete="on">
        <div class="fr-form-grid">
          <label>Schüler/in *
            <input type="text" id="frFormName" required value="${esc(f.schuelerName)}" placeholder="Vor- und Nachname">
          </label>
          <label>Klasse *
            <select id="frFormKlasse" required>${classOpts}</select>
          </label>
          <label>Beginn *
            <input type="date" id="frFormBeginn" required value="${esc(f.beginn)}">
          </label>
          <label>Ende *
            <input type="date" id="frFormEnde" required value="${esc(f.ende)}">
          </label>
          <label>Kategorie *
            <select id="frFormKat" required>
              ${KATEGORIE_CHOICES.map(
                  (k) =>
                      `<option value="${esc(k)}"${f.kategorie === k ? ' selected' : ''}>${esc(k)}</option>`
              ).join('')}
            </select>
          </label>
          <label>Klassenvorstand (E-Mail) *
            <input type="email" id="frFormKvEmail" required value="${esc(f.kvEmail)}" placeholder="kv@schule.at">
          </label>
          <label class="fr-form-span2">Begründung
            <textarea id="frFormBeschreibung" rows="3" placeholder="Kurz den Grund beschreiben…">${esc(f.beschreibung)}</textarea>
          </label>
        </div>
        <div class="fr-path ${path.multiDay ? 'is-multi' : ''}">
          <strong>Genehmigungspfad:</strong> ${esc(path.label)}
          ${path.dayCount != null ? ` <span class="muted">(${path.dayCount} Tag${path.dayCount === 1 ? '' : 'e'})</span>` : ''}
        </div>
        ${
            check.warnings.length
                ? `<ul class="fr-hints">${check.warnings.map((w) => `<li>${esc(w)}</li>`).join('')}</ul>`
                : ''
        }
        <div class="fr-actions">
          <button type="submit" class="btn btn-success" id="frBtnSubmit"><i class="bi bi-send"></i>Antrag einreichen</button>
          <button type="button" class="btn" id="frBtnFormReset">Zurücksetzen</button>
        </div>
      </form>
    </section>`;
}

function renderTable(items, state, opts) {
    const o = opts || {};
    const rows = items
        .map((it) => {
            const days = inclusiveDayCount(it.beginn, it.ende);
            const decide =
                o.showDecide && String(it.status) === 'Ausstehend'
                    ? `<button type="button" class="btn btn-sm" data-fr-approve="${esc(it.itemId)}" title="Status Genehmigt (Fallback)">✓</button>
                       <button type="button" class="btn btn-sm" data-fr-reject="${esc(it.itemId)}" title="Status Abgelehnt (Fallback)">✗</button>`
                    : '';
            return `<tr>
          <td><button type="button" class="fr-link" data-fr-detail="${esc(it.itemId)}">${esc(it.schuelerName || it.titel)}</button></td>
          <td>${esc(it.klasse)}</td>
          <td>${esc(formatDeDate(it.beginn))}${it.ende && it.ende !== it.beginn ? ' – ' + esc(formatDeDate(it.ende)) : ''}</td>
          <td>${days != null ? days : '–'}${it.multiDay ? ' <span class="fr-badge fr-badge--info">mehr</span>' : ''}</td>
          <td>${statusBadge(it.status)}</td>
          <td>${esc(it.kategorie)}</td>
          <td>${esc(it.kvName || it.kvEmail || '–')}</td>
          <td class="fr-td-actions">${decide}
            <button type="button" class="btn btn-sm" data-fr-detail="${esc(it.itemId)}"><i class="bi bi-eye"></i></button>
          </td>
        </tr>`;
        })
        .join('');
    return `
    <div class="fr-table-wrap">
      <table class="fr-table">
        <thead>
          <tr>
            <th>Schüler/in</th><th>Klasse</th><th>Zeitraum</th><th>Tage</th>
            <th>Status</th><th>Kategorie</th><th>KV</th><th></th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

function renderDetailModal(state) {
    const it = (state.items || []).find((x) => String(x.itemId) === String(state.detailId));
    if (!it) return '';
    return `
    <div class="fr-modal" role="dialog" aria-modal="true" aria-labelledby="frDetailTitle">
      <div class="fr-modal__card">
        <header class="fr-modal__head">
          <h2 id="frDetailTitle">${esc(it.schuelerName || it.titel)}</h2>
          <button type="button" class="btn" data-fr-close-detail aria-label="Schließen"><i class="bi bi-x-lg"></i></button>
        </header>
        <div class="fr-modal__body">
          <dl class="fr-dl">
            <dt>Status</dt><dd>${statusBadge(it.status)}</dd>
            <dt>Klasse</dt><dd>${esc(it.klasse)}</dd>
            <dt>Zeitraum</dt><dd>${esc(formatDeDate(it.beginn))} – ${esc(formatDeDate(it.ende))} (${it.dayCount ?? '–'} Tage)</dd>
            <dt>Genehmigung</dt><dd>${esc(it.approvalLabel)}</dd>
            <dt>Kategorie</dt><dd>${esc(it.kategorie)}</dd>
            <dt>KV</dt><dd>${esc(it.kvName || '–')} &lt;${esc(it.kvEmail || '')}&gt;</dd>
            <dt>Beschreibung</dt><dd>${esc(it.beschreibung || '–')}</dd>
            <dt>Bemerkungen</dt><dd>${esc(it.bemerkungen || '–')}</dd>
            <dt>Beantragt von</dt><dd>${esc(it.authorEmail || '–')}</dd>
          </dl>
        </div>
      </div>
    </div>`;
}

export function readFormFromDom(root) {
    const $ = (id) => (root || document).querySelector('#' + id);
    return {
        schuelerName: String(($('frFormName') && $('frFormName').value) || '').trim(),
        klasse: String(($('frFormKlasse') && $('frFormKlasse').value) || '').trim(),
        beginn: String(($('frFormBeginn') && $('frFormBeginn').value) || '').trim(),
        ende: String(($('frFormEnde') && $('frFormEnde').value) || '').trim(),
        kategorie: String(($('frFormKat') && $('frFormKat').value) || '').trim(),
        kvEmail: String(($('frFormKvEmail') && $('frFormKvEmail').value) || '')
            .trim()
            .toLowerCase(),
        beschreibung: String(($('frFormBeschreibung') && $('frFormBeschreibung').value) || '').trim(),
        status: 'Ausstehend'
    };
}

export function readFiltersFromDom(root) {
    const $ = (id) => (root || document).querySelector('#' + id);
    return {
        q: String(($('frFilterQ') && $('frFilterQ').value) || '').trim(),
        klasse: String(($('frFilterKlasse') && $('frFilterKlasse').value) || ''),
        status: String(($('frFilterStatus') && $('frFilterStatus').value) || ''),
        kategorie: String(($('frFilterKat') && $('frFilterKat').value) || ''),
        multiDay: String(($('frFilterMulti') && $('frFilterMulti').value) || '')
    };
}

export function applyKvFromClass(state) {
    const kv = resolveKvForClass(state.stammdaten.classes, state.form.klasse);
    if (kv) {
        state.form.kvEmail = kv.email;
        state.form.kvName = kv.name;
    }
}
