/**
 * UI-Rendering für den Schularbeiten-Planer.
 */
import {
    validateSchularbeit,
    computeDashboardKpis,
    buildWeeklyDistribution,
    toIsoDateOnly,
    dateInWindow,
    mondayOfWeekContaining,
    buildSchoolWeekDays,
    addDays,
    schularbeitDisplayTitle,
    formatSchularbeitZeitspanne,
    formatSchularbeitZeitKurz,
    normalizeBeginnUhrzeit,
    sortSchularbeitenByBeginn
} from './schularbeiten-planer-logic.js';
import {
    viewsForRole,
    filterSchularbeiten,
    EMPTY_SCHULARBEIT_FILTERS,
    labelMaps,
    formatDeDate,
    statusLabel,
    roleLabel,
    emptyForm,
    canEditSchularbeit,
    canDeleteSchularbeit,
    canAdminDecide,
    fachMetaForCode,
    planerPublicUrl,
    buildEmbedSnippet,
    scopeFromState,
    resolveStudentKlasseCode,
    schoolYearOptions,
    filterItemsForCalendarView
} from './schularbeiten-planer-state.js';
import { printTableCaption } from './schularbeiten-planer-export.js';
import { describeClassGroupLink, countPersonalCalendarLinked } from './schularbeiten-planer-calendar-sync.js';
import { loadPermissionsConfig } from './schularbeiten-planer-permissions.js';
import { htmlSchularbeitenEntraPermGrid } from './schularbeiten-permissions-ui.js';
import {
    isPlanerDemoRoleUiEnabled,
    canUsePlanerRoleSwitcher,
    roleSourceLabel,
    entraGroupsConfigured,
    sortPlanerRoles,
    PLANER_ROLE_ORDER
} from './schularbeiten-planer-entra-role.js';
import {
    LIST_TITLES,
    LIST_KEYS,
    LIST_DESCRIPTIONS,
    DEFAULT_FACH_META_STANDARD_DAUER,
    DEFAULT_FACH_META_PRO_SEMESTER,
    FACH_META_COLOR_PALETTE
} from './schularbeiten-planer-schema.js';

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
        s === 'fixiert' ? 'sa-badge sa-badge--ok' : s === 'abgelehnt' ? 'sa-badge sa-badge--bad' : 'sa-badge sa-badge--warn';
    return `<span class="${cls}">${esc(statusLabel(s))}</span>`;
}

function renderNavSession(state, ctx) {
    const { demoRoleUi, isSchueler, klasseCode, klasseLabel } = ctx;
    const switchable = demoRoleUi ? PLANER_ROLE_ORDER : sortPlanerRoles(state.planerRoles || []);
    const showRoleSwitcher = canUsePlanerRoleSwitcher(state.planerRoles || [], demoRoleUi);
    const activeSource = roleSourceLabel(
        (state.planerRoleSources && state.planerRoleSources[state.role]) || state.roleSource
    );

    const roleBlock = showRoleSwitcher
        ? `<div class="sa-nav__role-btns" role="group" aria-label="Rolle wählen">
            ${switchable
                .map(
                    (r) =>
                        `<button type="button" class="sa-chip${state.role === r ? ' is-active' : ''}" data-sa-role="${esc(r)}">${esc(
                            roleLabel(r)
                        )}</button>`
                )
                .join('')}
          </div>`
        : switchable.length === 1
          ? `<p class="sa-nav__role-fixed"><span class="sa-badge sa-badge--nav">${esc(roleLabel(state.role))}</span>${
                activeSource ? ` <small class="sa-nav__account-src">(${esc(activeSource)})</small>` : ''
            }</p>`
          : state.accountEmail
            ? `<p class="sa-nav__account-hint muted">Keine Planer-Rolle zugewiesen.</p>`
            : '';

    return `
        <div class="sa-nav__session" aria-label="Anmeldung und Rolle">
          <span class="sa-nav__role-label">Anmeldung</span>
          ${
              state.accountEmail
                  ? `<p class="sa-nav__account-name">${esc(state.accountName || state.accountEmail)}</p>`
                  : '<p class="sa-nav__account-hint">Mit Schul-Konto anmelden:</p>'
          }
          ${
              state.accountEmail || switchable.length
                  ? `<div class="sa-nav__session-role">
              <span class="sa-nav__session-role-label">Rolle${showRoleSwitcher ? ' wählen' : ''}</span>
              ${roleBlock}
              ${
                  isSchueler && klasseCode
                      ? `<span class="sa-badge sa-nav__account-klasse">${esc(klasseLabel)}</span>`
                      : ''
              }
            </div>`
                  : ''
          }
          ${
              isSchueler
                  ? `<label class="sa-nav__klasse" for="saDemoKlasse">Klasse
              <select id="saDemoKlasse" ${state.studentMatch ? 'disabled' : ''}>
                <option value="">– wählen –</option>
                ${(state.stammdaten.classes || [])
                    .map(
                        (c) =>
                            `<option value="${esc(c.code)}"${klasseCode === c.code ? ' selected' : ''}>${esc(
                                c.name || c.code
                            )}</option>`
                    )
                    .join('')}
              </select>
            </label>
            <p class="sa-nav__hint">${
                state.studentMatch
                    ? 'Klasse aus Stammdaten (E-Mail).'
                    : demoRoleUi
                      ? 'Demo: Klasse wählen, wenn keine Schüler-E-Mail in den Stammdaten.'
                      : state.planerRoleSources && state.planerRoleSources.schueler === 'admin-schueler' && !state.studentMatch
                        ? 'Als Admin: Klasse wählen, um die Schüler-Ansicht zu prüfen.'
                        : 'Klasse aus Stammdaten (Schüler-Entra-Gruppe oder E-Mail-Zuordnung).'
            }</p>`
                  : showRoleSwitcher && !demoRoleUi
                    ? '<p class="sa-nav__hint">Mehrere Berechtigungen – hier umschalten.</p>'
                    : demoRoleUi
                      ? '<p class="sa-nav__hint">Demo-Umschalter zur Vorschau.</p>'
                      : ''
          }
        </div>`;
}

function optionList(items, selected, emptyLabel) {
    const opts = [`<option value="">${esc(emptyLabel)}</option>`];
    (items || []).forEach((it) => {
        const v = it.code || it.value;
        const label = it.name || it.label || v;
        opts.push(
            `<option value="${esc(v)}"${String(selected) === String(v) ? ' selected' : ''}>${esc(label)}${
                it.code && it.name && it.code !== it.name ? ' (' + esc(it.code) + ')' : ''
            }</option>`
        );
    });
    return opts.join('');
}

/**
 * @param {object} state
 * @param {HTMLElement} root
 */
export function renderApp(state, root) {
    if (!root) return;
    const isSchueler = state.role === 'schueler';
    const views = viewsForRole(state.role);
    const demoRoleUi = isPlanerDemoRoleUiEnabled(
        state.entraGroupsConfigured || entraGroupsConfigured(loadPermissionsConfig()),
        state.demoRoleOverride
    );
    const klasseCode = resolveStudentKlasseCode(state);
    const klasseLabel =
        (state.stammdaten.classes || []).find((c) => c.code === klasseCode)?.name || klasseCode;

    root.innerHTML = `
    <div class="sa-shell">
      <aside class="sa-nav" aria-label="Planer-Navigation">
        <div class="sa-nav__brand">
          <i class="bi bi-mortarboard-fill" aria-hidden="true"></i>
          <div>
            <strong>Schularbeiten</strong>
            <span>HAK Planer</span>
          </div>
        </div>
        <nav class="sa-nav__list">
          ${views
              .map(
                  (v) => `
            <button type="button" class="sa-nav__btn${state.view === v.id ? ' is-active' : ''}" data-sa-view="${esc(v.id)}">
              <i class="bi ${v.icon}" aria-hidden="true"></i>${esc(v.label)}
            </button>`
              )
              .join('')}
        </nav>
        <div class="sa-nav__bottom">
        ${renderNavSession(state, { demoRoleUi, isSchueler, klasseCode, klasseLabel })}
        </div>
      </aside>
      <div class="sa-main">
        ${state.error ? `<div class="sa-alert sa-alert--bad" role="alert">${esc(state.error)}</div>` : ''}
        ${
            state.roleHint
                ? `<div class="sa-alert sa-alert--info" role="status">${esc(state.roleHint)}
            ${
                state.role !== 'schueler' && demoRoleUi
                    ? '<button type="button" class="btn btn-sm" id="saBtnRoleAdmin" style="margin-left:8px;">Als Admin anzeigen</button>'
                    : ''
            }</div>`
                : ''
        }
        ${state.loading ? skeletonBlock() : ''}
        <div id="saView" class="sa-view${state.loading ? ' is-loading' : ''}">
          ${state.loading ? '' : renderView(state)}
        </div>
      </div>
    </div>
    ${state.detailId ? renderDetailModal(state) : ''}
  `;
}

function skeletonBlock() {
    return `<div class="sa-skel" aria-busy="true" aria-live="polite">
      <div class="sa-skel__card"></div><div class="sa-skel__card"></div><div class="sa-skel__card"></div>
    </div>`;
}

function mergeStammdatenOptions(stammdatenRows, codesFromData) {
    const map = new Map();
    (stammdatenRows || []).forEach((row) => {
        const code = String((row && row.code) || '').trim();
        if (!code) return;
        map.set(code, { code, name: String(row.name || code).trim() || code });
    });
    (codesFromData || []).forEach((code) => {
        const c = String(code || '').trim();
        if (c && !map.has(c)) map.set(c, { code: c, name: c });
    });
    return Array.from(map.values()).sort((a, b) =>
        String(a.name).localeCompare(String(b.name), 'de')
    );
}

function itemsInFilterScope(state, scope) {
    return filterSchularbeiten(state.items, EMPTY_SCHULARBEIT_FILTERS, scope);
}

function subjectCodesForFilterPills(state, scope) {
    const labels = labelMaps(state.stammdaten);
    const seen = new Set();
    itemsInFilterScope(state, scope).forEach((sa) => {
        const c = sa && sa.fachCode;
        if (c) seen.add(String(c).trim());
    });
    (state.fachMeta || []).forEach((m) => {
        const c = m && m.fachCode;
        if (c) seen.add(String(c).trim());
    });
    return Array.from(seen).sort((a, b) =>
        String(labels.fach[a] || a).localeCompare(String(labels.fach[b] || b), 'de')
    );
}

function renderFilterSelect(id, label, icon, items, selected, emptyLabel) {
    return `<div class="sa-filter-field">
      <label class="sa-filter-field__label" for="${esc(id)}"><i class="bi ${icon}" aria-hidden="true"></i>${esc(label)}</label>
      <div class="sa-filter-field__control">
        <select id="${esc(id)}" class="sa-filter-select" aria-label="${esc(label)}">${optionList(items, selected, emptyLabel)}</select>
        <i class="bi bi-chevron-down sa-filter-field__chev" aria-hidden="true"></i>
      </div>
    </div>`;
}

function renderStatusPills(activeStatus) {
    const cur = String(activeStatus || '').toLowerCase();
    const opts = [
        { value: '', label: 'Alle' },
        { value: 'beantragt', label: 'Beantragt' },
        { value: 'fixiert', label: 'Fixiert' },
        { value: 'abgelehnt', label: 'Abgelehnt' }
    ];
    return `<div class="sa-filter-pills" role="group" aria-label="Status filtern">
      ${opts
          .map((o) => {
              const on = o.value === '' ? !cur : cur === o.value;
              const cls =
                  'sa-filter-pill sa-filter-pill--status' +
                  (on ? ' is-active' : '') +
                  (o.value === 'fixiert' && on ? ' sa-filter-pill--ok' : '') +
                  (o.value === 'beantragt' && on ? ' sa-filter-pill--warn' : '') +
                  (o.value === 'abgelehnt' && on ? ' sa-filter-pill--bad' : '');
              return `<button type="button" class="${cls}" data-sa-filter-pill data-sa-filter-key="status" data-sa-filter-value="${esc(
                  o.value
              )}" aria-pressed="${on ? 'true' : 'false'}">${esc(o.label)}</button>`;
          })
          .join('')}
    </div>`;
}

function renderSchuljahrFilterField(state) {
    const disabled = !state.bootstrapped && !state.localDemoOnly;
    const years = schoolYearOptions(state);
    const sel = String(state.schuljahr || '');
    const opts = years
        .map((y) => `<option value="${esc(y)}"${sel === String(y) ? ' selected' : ''}>${esc(y)}</option>`)
        .join('');
    return `<div class="sa-filter-field">
      <label class="sa-filter-field__label" for="saSchuljahr"><i class="bi bi-calendar3" aria-hidden="true"></i>Schuljahr</label>
      <div class="sa-filter-field__control">
        <select id="saSchuljahr" class="sa-filter-select" aria-label="Schuljahr"${disabled ? ' disabled' : ''}>${opts}</select>
        <i class="bi bi-chevron-down sa-filter-field__chev" aria-hidden="true"></i>
      </div>
    </div>`;
}

function renderFachFilterPills(codes, labels, fachMeta, activeFach) {
    const cur = String(activeFach || '').trim();
    const allOn = !cur;
    let html = `<button type="button" class="sa-filter-pill sa-filter-pill--fach${allOn ? ' is-active' : ''}" data-sa-filter-pill data-sa-filter-key="fach" data-sa-filter-value="" aria-pressed="${allOn ? 'true' : 'false'}">Alle Fächer</button>`;
    codes.forEach((code, i) => {
        const on = cur === code;
        const color = fachColor(code, i, fachMeta);
        const label = (labels.fach && labels.fach[code]) || code;
        html += `<button type="button" class="sa-filter-pill sa-filter-pill--fach${on ? ' is-active' : ''}" style="--sa-fach:${esc(
            color
        )}" data-sa-filter-pill data-sa-filter-key="fach" data-sa-filter-value="${esc(code)}" aria-pressed="${
            on ? 'true' : 'false'
        }" title="${esc(label)}"><i class="sa-filter-pill__dot" aria-hidden="true"></i>${esc(label)}</button>`;
    });
    return `<div class="sa-filter-pills sa-filter-pills--fach" role="group" aria-label="Fach filtern">${html}</div>`;
}

function filterBar(state, opts) {
    const sd = state.stammdaten;
    const labels = labelMaps(sd);
    const scopeAll = typeof opts === 'boolean' ? opts : !!(opts && opts.scopeAll);
    const onlyMine = typeof opts === 'object' && opts ? !!opts.onlyMine : false;
    const isSchueler = state.role === 'schueler';
    const scope = scopeFromState(state, { scopeAll: scopeAll && !isSchueler, onlyMine: onlyMine && !isSchueler });
    const scopeItems = itemsInFilterScope(state, scope);
    const filteredCount = filterSchularbeiten(state.items, state.filters, scope).length;
    const scopeTotal = scopeItems.length;
    const loadedTotal = (state.items || []).length;
    const fachCodes = subjectCodesForFilterPills(state, scope);
    const klasseOptions = mergeStammdatenOptions(
        sd.classes,
        scopeItems.map((sa) => sa.klasseCode).filter(Boolean)
    );
    const lehrerOptions = mergeStammdatenOptions(
        sd.teachers,
        scopeItems.map((sa) => sa.lehrerCode).filter(Boolean)
    );
    const metaLabel = isSchueler
        ? `${filteredCount} fixierte Termine`
        : `${filteredCount} Treffer · ${scopeTotal} im Schuljahr` +
          (loadedTotal > scopeTotal ? ` · ${loadedTotal} geladen` : '');

    if (isSchueler) {
        return `
    <div class="sa-filters sa-filters--modern sa-filters--compact" data-sa-filters>
      <div class="sa-filters__grid">
        <div class="sa-filters__group">
          <span class="sa-filters__group-title">Schuljahr &amp; Fach</span>
          <div class="sa-filters__group-cols">
            ${renderSchuljahrFilterField(state)}
          </div>
          ${renderFachFilterPills(fachCodes, labels, state.fachMeta, state.filters.fach)}
        </div>
      </div>
      <div class="sa-filters__foot">
        <button type="button" class="btn btn-sm" id="saFilterReset"><i class="bi bi-arrow-counterclockwise"></i>Zurücksetzen</button>
        <span class="sa-filters__meta">${metaLabel}</span>
      </div>
    </div>`;
    }

    return `
    <div class="sa-filters sa-filters--modern sa-filters--compact" data-sa-filters>
      <div class="sa-filters__grid">
        <div class="sa-filters__group sa-filters__group--scope">
          <span class="sa-filters__group-title">Zuordnung</span>
          <div class="sa-filters__group-cols">
            ${renderSchuljahrFilterField(state)}
            ${renderFilterSelect('saFilterKlasse', 'Klasse', 'bi-people', klasseOptions, state.filters.klasse, 'Alle Klassen')}
            ${renderFilterSelect('saFilterLehrer', 'Lehrkraft', 'bi-person-badge', lehrerOptions, state.filters.lehrer, 'Alle Lehrer:innen')}
          </div>
        </div>
        <div class="sa-filters__group sa-filters__group--status">
          <span class="sa-filters__group-title">Status</span>
          ${renderStatusPills(state.filters.status)}
        </div>
        <div class="sa-filters__group sa-filters__group--fach">
          <span class="sa-filters__group-title">Fach</span>
          ${renderFachFilterPills(fachCodes, labels, state.fachMeta, state.filters.fach)}
        </div>
      </div>
      <div class="sa-filters__foot">
        <button type="button" class="btn btn-sm" id="saFilterReset"><i class="bi bi-arrow-counterclockwise"></i>Zurücksetzen</button>
        <span class="sa-filters__meta">${metaLabel}</span>
      </div>
    </div>`;
}

function renderView(state) {
    if (state.planerAccessDenied) {
        return `<section class="tm-panel">
      <h2 class="sa-h2">Kein Zugriff auf den Schularbeiten-Planer</h2>
      <p>Ihr Microsoft-Konto ist keiner der konfigurierten Entra-Gruppen zugeordnet und nicht als Einzelperson eingetragen (Verwaltung, Lehrkräfte oder Schüler).</p>
      <p>Bitte wenden Sie sich an die Schul-IT. Gruppen und Einzelpersonen legen Sie unter <strong>Administration → SharePoint-Berechtigungen</strong> fest.</p>
    </section>`;
    }
    if (!state.bootstrapped && !state.ctx && !state.localDemoOnly) {
        if (state.role === 'admin') {
            return `<section class="tm-panel"><p>SharePoint noch nicht verbunden. Unter <strong>Administration</strong> die Site-URL eintragen und <strong>Laden</strong>.</p>
          <p><a href="sharepoint-liste-schularbeiten.html">→ Schularbeiten-Listen anlegen</a></p></section>`;
        }
        return `<section class="tm-panel"><p>Daten werden geladen … Falls nichts erscheint, bitte die Schul-IT (SharePoint-Verbindung in der Administration).</p></section>`;
    }
    if (state.role === 'schueler' && ['neu', 'meine', 'admin', 'regeln'].includes(state.view)) {
        return renderSchuelerDashboard(state);
    }
    switch (state.view) {
        case 'kalender':
            return renderKalender(state);
        case 'neu':
            return renderNeu(state);
        case 'meine':
            return renderMeine(state);
        case 'admin':
            return renderAdmin(state);
        case 'regeln':
            return renderRegeln(state);
        case 'export':
            return renderExport(state);
        case 'liste':
            return renderListe(state);
        case 'dashboard':
        default:
            return renderDashboard(state);
    }
}

function scopedItems(state, onlyMine) {
    return filterSchularbeiten(
        state.items,
        state.filters,
        scopeFromState(state, { onlyMine: !!onlyMine })
    );
}

function renderDashboard(state) {
    if (state.role === 'schueler') return renderSchuelerDashboard(state);

    const items =
        state.role === 'admin'
            ? filterSchularbeiten(state.items, state.filters, scopeFromState(state, { scopeAll: true }))
            : scopedItems(state, true);
    const labels = labelMaps(state.stammdaten);
    const weekMonday =
        state.dashboardWeekMonday || mondayOfWeekContaining(toIsoDateOnly(new Date()));
    const todayIso = toIsoDateOnly(new Date());

    return `
    ${filterBar(state, { scopeAll: state.role === 'admin' })}
    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-calendar-week"></i>
        <div><h3>Unterrichtswoche</h3>
        <p>Mo–Fr · ${esc(formatDeDate(weekMonday))} – ${esc(formatDeDate(addDays(weekMonday, 4) || weekMonday))}</p></div>
        <div class="sa-week-nav">
          <button type="button" class="btn btn-sm" data-sa-week-prev title="Vorherige Woche"><i class="bi bi-chevron-left"></i></button>
          <button type="button" class="btn btn-sm" data-sa-week-today title="Aktuelle Woche">Heute</button>
          <button type="button" class="btn btn-sm" data-sa-week-next title="Nächste Woche"><i class="bi bi-chevron-right"></i></button>
        </div>
      </div>
      ${renderSchoolWeekBoard(items, labels, state.fachMeta, weekMonday, todayIso)}
      <p class="muted" style="margin:12px 0 0;font-size:0.88em;">
        Alle Einträge tabellarisch: <button type="button" class="sa-link" data-sa-view="liste">Liste</button>
      </p>
    </section>`;
}

function renderListe(state) {
    if (state.role === 'schueler') {
        const klasseCode = resolveStudentKlasseCode(state);
        if (!klasseCode) {
            return `<section class="tm-panel"><p class="muted">Bitte Klasse wählen (Demo) oder Stammdaten mit Schüler-E-Mail.</p></section>`;
        }
    }
    const items =
        state.role === 'admin'
            ? filterSchularbeiten(state.items, state.filters, scopeFromState(state, { scopeAll: true }))
            : state.role === 'schueler'
              ? scopedItems(state, false)
              : scopedItems(state, true);
    const labels = labelMaps(state.stammdaten);
    const listItems = items.slice().sort((a, b) => String(a.datum).localeCompare(String(b.datum)));

    return `
    ${filterBar(state, { scopeAll: state.role === 'admin' })}
    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-list-ul"></i>
        <div><h3>Listenansicht</h3><p>${listItems.length} Einträge nach aktuellem Filter</p></div>
      </div>
      ${tableSchularbeiten(listItems, labels, {
          showActions: state.role === 'admin' || state.role === 'lehrer',
          fachMeta: state.fachMeta,
          scope: scopeFromState(state)
      })}
    </section>`;
}

function renderSchoolWeekBoard(items, labels, fachMeta, weekMonday, todayIso) {
    const week = buildSchoolWeekDays(weekMonday);
    const byDay = new Map();
    week.days.forEach((d) => byDay.set(d.iso, []));
    (items || []).forEach((sa) => {
        const d = toIsoDateOnly(sa.datum);
        if (!d || !byDay.has(d)) return;
        byDay.get(d).push(sa);
    });
    week.days.forEach((d) => {
        byDay.set(d.iso, sortSchularbeitenByBeginn(byDay.get(d.iso)));
    });

    const cols = week.days
        .map((d) => {
            const dayItems = byDay.get(d.iso) || [];
            const isToday = d.iso === todayIso;
            const cards = dayItems.length
                ? dayItems
                      .map((sa) => {
                          const klasse = labels.klasse[sa.klasseCode] || sa.klasseCode;
                          const zeitKurz = formatSchularbeitZeitKurz(sa);
                          return `<article class="sa-week-card" style="--sa-fach:${esc(fachColor(sa.fachCode, -1, fachMeta))}">
              <button type="button" class="sa-week-card__main sa-link" data-sa-detail="${esc(sa.itemId)}">
                <span class="sa-week-card__fach">${fachChip(sa.fachCode, labels, fachMeta)}</span>
                <strong class="sa-week-card__thema">${esc(schularbeitDisplayTitle(sa))}</strong>
                <span class="sa-week-card__meta"><span class="sa-week-card__time">${esc(zeitKurz || '–')}</span> · ${esc(klasse)} · ${esc(labels.lehrer[sa.lehrerCode] || sa.lehrerCode || '–')}</span>
              </button>
              ${statusBadge(sa.status)}
            </article>`;
                      })
                      .join('')
                : '<p class="sa-week-empty muted">Keine Schularbeiten</p>';
            return `<div class="sa-week-day${isToday ? ' is-today' : ''}">
        <header class="sa-week-day__head">
          <span class="sa-week-day__wd">${esc(d.weekday)}</span>
          <span class="sa-week-day__date">${esc(formatDeDate(d.iso))}</span>
        </header>
        <div class="sa-week-day__body">${cards}</div>
      </div>`;
        })
        .join('');

    return `<div class="sa-week-board" role="region" aria-label="Schularbeiten Unterrichtswoche">${cols}</div>`;
}

function renderSchuelerDashboard(state) {
    const klasseCode = resolveStudentKlasseCode(state);
    const labels = labelMaps(state.stammdaten);
    const items = scopedItems(state, false)
        .slice()
        .sort((a, b) => String(a.datum).localeCompare(String(b.datum)));
    const weekMonday =
        state.dashboardWeekMonday || mondayOfWeekContaining(toIsoDateOnly(new Date()));
    const todayIso = toIsoDateOnly(new Date());

    if (!klasseCode) {
        return `
    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-people"></i>
        <div><h3>Meine Klasse</h3><p>Bitte Klasse wählen oder Stammdaten mit Schüler-E-Mail pflegen.</p></div>
      </div>
      <p class="muted">Ohne Klassen-Zuordnung können keine Schularbeiten angezeigt werden.</p>
    </section>`;
    }

    return `
    ${filterBar(state, {})}
    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-calendar-week"></i>
        <div><h3>Unterrichtswoche</h3>
        <p>Mo–Fr · ${esc(formatDeDate(weekMonday))} – ${esc(formatDeDate(addDays(weekMonday, 4) || weekMonday))}</p></div>
        <div class="sa-week-nav">
          <button type="button" class="btn btn-sm" data-sa-week-prev title="Vorherige Woche"><i class="bi bi-chevron-left"></i></button>
          <button type="button" class="btn btn-sm" data-sa-week-today title="Aktuelle Woche">Heute</button>
          <button type="button" class="btn btn-sm" data-sa-week-next title="Nächste Woche"><i class="bi bi-chevron-right"></i></button>
        </div>
      </div>
      ${renderSchoolWeekBoard(items, labels, state.fachMeta, weekMonday, todayIso)}
      <p class="muted" style="margin:12px 0 0;font-size:0.88em;">
        Alle Termine: <button type="button" class="sa-link" data-sa-view="liste">Liste</button>
      </p>
    </section>`;
}

/** Feste Palette für Fächer (ohne SharePoint-Meta). */
const FACH_COLORS = [
    '#6366f1',
    '#0ea5e9',
    '#10b981',
    '#f59e0b',
    '#ec4899',
    '#8b5cf6',
    '#14b8a6',
    '#f97316',
    '#64748b',
    '#ef4444'
];

function normalizeHexColor(raw) {
    const s = String(raw || '').trim();
    if (!s) return '';
    const withHash = s.charAt(0) === '#' ? s : '#' + s;
    if (!/^#[0-9a-fA-F]{3}$|^#[0-9a-fA-F]{6}$|^#[0-9a-fA-F]{8}$/.test(withHash)) return '';
    if (withHash.length === 4) {
        return (
            '#' +
            withHash[1] +
            withHash[1] +
            withHash[2] +
            withHash[2] +
            withHash[3] +
            withHash[3]
        );
    }
    return withHash.slice(0, 7);
}

function fachColor(code, index, fachMeta) {
    const meta = fachMetaForCode(fachMeta, code);
    const fromMeta = normalizeHexColor(meta && meta.farbe);
    if (fromMeta) return fromMeta;
    let h = 0;
    const s = String(code || '');
    for (let i = 0; i < s.length; i++) h = (h + s.charCodeAt(i) * (i + 1)) % FACH_COLORS.length;
    return FACH_COLORS[(index >= 0 ? index : h) % FACH_COLORS.length];
}

/** Hell/dunkel Text für Lesbarkeit auf Fachfarbe. */
function contrastOn(hex) {
    const h = normalizeHexColor(hex) || '#64748b';
    const r = parseInt(h.slice(1, 3), 16);
    const g = parseInt(h.slice(3, 5), 16);
    const b = parseInt(h.slice(5, 7), 16);
    const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return lum > 0.62 ? '#0f172a' : '#ffffff';
}

function fachChip(code, labels, fachMeta) {
    const color = fachColor(code, -1, fachMeta);
    const label = (labels && labels.fach && labels.fach[code]) || code || '–';
    return `<span class="sa-fach" style="--sa-fach:${esc(color)};--sa-fach-fg:${esc(contrastOn(color))}" title="${esc(
        label
    )}"><i aria-hidden="true"></i>${esc(label)}</span>`;
}

/**
 * @param {{ weeks: object[], faecher: string[] }} dist
 * @param {{ fach: Record<string, string> }} labels
 * @param {object[]} [fachMeta]
 */
function renderWeeklyChart(dist, labels, fachMeta) {
    const weeks = (dist && dist.weeks) || [];
    const faecher = (dist && dist.faecher) || [];
    if (!weeks.length) {
        return '<p class="muted">Keine Daten für die Wochenverteilung.</p>';
    }
    const max = Math.max(1, ...weeks.map((w) => w.total || 0));
    const W = 520;
    const H = 220;
    const padL = 36;
    const padB = 36;
    const padT = 12;
    const padR = 12;
    const chartW = W - padL - padR;
    const chartH = H - padT - padB;
    const gap = 8;
    const barW = Math.max(12, (chartW - gap * (weeks.length - 1)) / weeks.length);

    const bars = weeks
        .map((w, i) => {
            const x = padL + i * (barW + gap);
            let y = padT + chartH;
            const stacks = [];
            faecher.forEach((fach, fi) => {
                const n = (w.byFach && w.byFach[fach]) || 0;
                if (!n) return;
                const h = (n / max) * chartH;
                y -= h;
                stacks.push(
                    `<rect class="sa-chart__seg" x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(
                        1
                    )}" height="${Math.max(1, h).toFixed(1)}" fill="${fachColor(fach, fi, fachMeta)}"><title>${esc(
                        (labels.fach[fach] || fach) + ' · ' + w.label + ': ' + n
                    )}</title></rect>`
                );
            });
            if (!w.total) {
                stacks.push(
                    `<rect x="${x.toFixed(1)}" y="${(padT + chartH - 2).toFixed(1)}" width="${barW.toFixed(
                        1
                    )}" height="2" fill="#cbd5e1" opacity="0.5"></rect>`
                );
            }
            return (
                stacks.join('') +
                `<text class="sa-chart__xlabel" x="${(x + barW / 2).toFixed(1)}" y="${H - 10}" text-anchor="middle">${esc(
                    w.label
                )}</text>`
            );
        })
        .join('');

    const yTicks = [];
    const tickCount = Math.min(4, max);
    for (let t = 0; t <= tickCount; t++) {
        const val = Math.round((max * t) / tickCount);
        const y = padT + chartH - (val / max) * chartH;
        yTicks.push(
            `<line x1="${padL}" x2="${W - padR}" y1="${y.toFixed(1)}" y2="${y.toFixed(
                1
            )}" class="sa-chart__grid"/>` +
                `<text class="sa-chart__ylabel" x="${padL - 6}" y="${(y + 3).toFixed(1)}" text-anchor="end">${val}</text>`
        );
    }

    const legend = faecher.length
        ? `<div class="sa-chart__legend">${faecher
              .map(
                  (f, i) =>
                      `<span><i style="background:${fachColor(f, i, fachMeta)}"></i>${esc(labels.fach[f] || f)}</span>`
              )
              .join('')}</div>`
        : '<p class="muted" style="margin:8px 0 0;font-size:0.88em;">Noch keine Termine in den nächsten Wochen.</p>';

    return `
    <div class="sa-chart">
      <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Verteilung Schularbeiten pro Kalenderwoche">
        ${yTicks.join('')}
        ${bars}
      </svg>
      ${legend}
    </div>`;
}

function renderIntranetPanel(state) {
    const url = planerPublicUrl();
    const snippet = buildEmbedSnippet(url);
    return `
    <section class="tm-panel sa-intranet">
      <div class="tm-panel__head"><i class="bi bi-house-door"></i>
        <div><h3>Im Intranet verlinken</h3>
        <p>Quicklink oder Einbettung auf der SharePoint-Kommunikationssite</p></div>
      </div>
      <ol class="sa-intranet__steps">
        <li>Intranet bearbeiten → <strong>Quicklink</strong> / Link-Webpart „Schularbeiten planen“.</li>
        <li>Ziel-URL: Planer-Adresse (unten kopieren).</li>
        <li>Optional: HTML-Webpart mit dem Snippet (Link-Karte; iframe auskommentiert).</li>
      </ol>
      <div class="sa-intranet__copy">
        <input type="text" id="saPlanerUrl" readonly value="${esc(url)}" aria-label="Planer-URL">
        <button type="button" class="btn" id="saBtnCopyUrl"><i class="bi bi-clipboard"></i>URL kopieren</button>
        <button type="button" class="btn" id="saBtnCopyEmbed"><i class="bi bi-code-slash"></i>Snippet kopieren</button>
      </div>
      <details class="tm-tech" style="margin-top:10px;">
        <summary>Embed-Snippet</summary>
        <pre id="saEmbedSnippet" class="sa-embed-pre">${esc(snippet)}</pre>
      </details>
      <p class="muted" style="margin:10px 0 0;font-size:0.88em;">
        Status-Mails: <a href="pa-schularbeiten-mail.html">Power-Automate-Rezept</a> ·
        Listen: <a href="sharepoint-liste-schularbeiten.html">Paket erneut ausführen</a>
        ${state.listsMissingFachMeta ? ' · <strong>SAP-FachMeta fehlt noch</strong> – Setup erneut starten.' : ''}
      </p>
    </section>`;
}

function renderNeu(state) {
    if (state.role === 'schueler') {
        return `<section class="tm-panel"><p>Schüler:innen können keine Schularbeiten beantragen.</p></section>`;
    }
    const f = state.form || emptyForm();
    const sd = state.stammdaten;
    const rules = {
        maxProTag: state.rules.maxProTag,
        maxProWoche: state.rules.maxProWoche,
        ankuendigungsfristTage: state.rules.ankuendigungsfristTage,
        sperreVorNotenkonferenzTage: state.rules.sperreVorNotenkonferenzTage
    };
    const draft = {
        schularbeitId: f.schularbeitId || 'sa-draft',
        fachCode: f.fachCode,
        klasseCode: f.klasseCode,
        lehrerCode: f.lehrerCode,
        datum: f.datum,
        beginnUhrzeit: f.beginnUhrzeit,
        dauerMinuten: Number(f.dauerMinuten) || 100,
        semester: f.semester,
        status: 'beantragt'
    };
    const result = validateSchularbeit({
        draft,
        existing: state.items,
        rules,
        windows: state.windows,
        fachMeta: (() => {
            const m = fachMetaForCode(state.fachMeta, f.fachCode);
            if (!m) return undefined;
            return {
                proSemester: m.proSemester,
                minDauer: 50,
                maxDauer: Math.max(150, m.standardDauer || DEFAULT_FACH_META_STANDARD_DAUER)
            };
        })()
    });
    const editing = !!state.editingItemId;

    return `
    <section class="sa-split">
      <div class="tm-panel">
        <div class="tm-panel__head"><i class="bi bi-plus-circle"></i>
          <div><h3>${editing ? 'Antrag bearbeiten' : 'Neue Schularbeit'}</h3>
          <p>Live-Prüfung gegen Regelwerk &amp; Sperrzeiten</p></div>
        </div>
        <div class="sa-form">
          <div class="tm-field"><label for="saTitel">Titel</label>
            <input id="saTitel" type="text" value="${esc(f.titel)}" placeholder="z. B. 1. Schularbeit Deutsch" maxlength="250"></div>
          <div class="tm-field"><label for="saThema">Thema <span class="muted">(optional)</span></label>
            <input id="saThema" type="text" value="${esc(f.thema)}" placeholder="z. B. Erörterung: Digitalisierung" maxlength="250"></div>
          <div class="sa-form__row">
            <div class="tm-field"><label for="saFach">Fach</label>
              <select id="saFach">${optionList(sd.subjects, f.fachCode, 'Bitte wählen')}</select></div>
            <div class="tm-field"><label for="saKlasse">Klasse</label>
              <select id="saKlasse">${optionList(sd.classes, f.klasseCode, 'Bitte wählen')}</select></div>
          </div>
          <div class="sa-form__row">
            <div class="tm-field"><label for="saDatum">Wunschtermin</label>
              <input id="saDatum" type="date" value="${esc(f.datum)}"></div>
            <div class="tm-field"><label for="saBeginn">Beginn (Uhrzeit)</label>
              <input id="saBeginn" type="time" value="${esc(f.beginnUhrzeit || '08:00')}" step="300"></div>
            <div class="tm-field"><label for="saDauer">Dauer (Min.)</label>
              <input id="saDauer" type="number" min="50" max="300" step="25" value="${esc(f.dauerMinuten)}"></div>
          </div>
          <p class="muted" style="margin:-4px 0 8px;font-size:0.85em;">Ende der Schularbeit = Beginn + Dauer (wird für Kalender-Export verwendet).</p>
          <div class="sa-form__row">
            <div class="tm-field"><label for="saSemester">Semester</label>
              <select id="saSemester">
                <option value="WS"${f.semester === 'WS' ? ' selected' : ''}>Wintersemester</option>
                <option value="SS"${f.semester === 'SS' ? ' selected' : ''}>Sommersemester</option>
              </select></div>
            <div class="tm-field"><label for="saLehrer">Lehrer:in</label>
              <select id="saLehrer">${optionList(sd.teachers, f.lehrerCode, 'Bitte wählen')}</select></div>
          </div>
          <div class="tm-field"><label for="saNotiz">Notiz</label>
            <textarea id="saNotiz" rows="3">${esc(f.notiz)}</textarea></div>
          <div class="tm-actions">
            <button type="button" class="btn btn-success" id="saBtnSubmit">
              <i class="bi bi-send"></i>${editing ? 'Änderungen speichern' : 'Antrag einreichen'}
            </button>
            <button type="button" class="btn" id="saBtnCancelForm">Abbrechen</button>
          </div>
        </div>
      </div>
      <div class="tm-panel sa-rules-live">
        <div class="tm-panel__head"><i class="bi bi-shield-check"></i>
          <div><h3>Live-Regelprüfung</h3><p>${esc(state.rules.name || 'Regelwerk')}</p></div>
        </div>
        <div class="sa-alert ${result.canSubmit ? 'sa-alert--ok' : 'sa-alert--bad'}">
          ${
              result.canSubmit
                  ? 'Alle Pflichtregeln erfüllt. Der Antrag kann eingereicht werden.'
                  : 'Es gibt blockierende Regelverstöße.'
          }
        </div>
        ${
            result.errors.length
                ? `<ul class="sa-rule-list sa-rule-list--err">${result.errors.map((e) => `<li>${esc(e)}</li>`).join('')}</ul>`
                : ''
        }
        ${
            result.warnings.length
                ? `<ul class="sa-rule-list sa-rule-list--warn">${result.warnings.map((w) => `<li>${esc(w)}</li>`).join('')}</ul>`
                : ''
        }
        <p class="muted" style="margin-top:12px;font-size:0.88em;">
          Max. ${esc(rules.maxProTag)} / Tag · Max. ${esc(rules.maxProWoche)} / Woche ·
          Ankündigung ${esc(rules.ankuendigungsfristTage)} Tage
        </p>
      </div>
    </section>`;
}

function renderMeine(state) {
    const labels = labelMaps(state.stammdaten);
    const items = scopedItems(state, true).slice().sort((a, b) => String(a.datum).localeCompare(String(b.datum)));
    const ownSyncable = personalCalendarSyncItems(state);
    let emptyHint = '';
    if (!items.length && state.items.length) {
        emptyHint = state.teacherMatch
            ? `<p class="muted">Keine eigenen Einträge für ${esc(state.teacherMatch.name || state.teacherMatch.code)}.</p>`
            : `<p class="muted">Keine Einträge zu Ihrer Anmeldung (${esc(state.accountEmail || '—')}).
            „Meine Schularbeiten“ zeigt nur Anträge Ihrer Lehrkraft-Zuordnung.
            Alle Termine: <button type="button" class="sa-link" data-sa-view="dashboard">Dashboard</button>
            ${state.role === 'admin' ? ' bzw. Administration.' : ''}
            Für die Lehrer-Ansicht: Rolle „Lehrer“ und E-Mail in den Stammdaten pflegen.</p>`;
    } else if (!items.length) {
        emptyHint = '<p class="muted">Keine Einträge.</p>';
    }
    return `
    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-card-checklist"></i>
        <div><h3>Meine Schularbeiten</h3><p>Eigene Anträge und Termine</p></div>
      </div>
      ${filterBar(state, { onlyMine: true, scopeAll: false })}
      ${
          ownSyncable.length
              ? `<div class="sa-meine-cal-bar">
        <button type="button" class="btn btn-sm" data-sa-sync-my-cal>
          <i class="bi bi-calendar-plus"></i>In meinen Outlook-Kalender (${ownSyncable.length})
        </button>
        <span class="muted">Fixierte und beantragte eigene Termine · Filter oben beachten</span>
      </div>`
              : ''
      }
      ${items.length ? tableSchularbeiten(items, labels, { showActions: true, scope: scopeFromState(state), fachMeta: state.fachMeta }) : emptyHint}
    </section>`;
}

function tableSchularbeiten(items, labels, opts) {
    if (!items.length) return '<p class="muted">Keine Einträge.</p>';
    const showActions = opts && opts.showActions;
    const fachMeta = (opts && opts.fachMeta) || [];
    const scope = opts && opts.scope;
    return `<div class="sa-table-wrap"><table class="sa-table"><thead><tr>
      <th>Datum</th><th>Fach</th><th>Klasse</th><th>Titel</th><th>Dauer</th><th>Sem.</th><th>Status</th>
      ${showActions ? '<th>Aktionen</th>' : ''}
    </tr></thead><tbody>
    ${items
        .map((sa) => {
            const editable = scope ? canEditSchularbeit(sa, scope) : schularbeitStatus(sa) === 'beantragt';
            const deletable = scope ? canDeleteSchularbeit(sa, scope) : editable;
            const color = fachColor(sa.fachCode, -1, fachMeta);
            return `<tr class="sa-row" style="--sa-fach:${esc(color)}">
        <td><button type="button" class="sa-link" data-sa-detail="${esc(sa.itemId)}">${esc(formatDeDate(sa.datum))}</button></td>
        <td>${fachChip(sa.fachCode, labels, fachMeta)}</td>
        <td>${esc(labels.klasse[sa.klasseCode] || sa.klasseCode)}</td>
        <td>${esc(schularbeitDisplayTitle(sa))}</td>
        <td>${esc(sa.dauerMinuten)}'</td>
        <td>${esc(sa.semester)}</td>
        <td>${statusBadge(sa.status)}</td>
        ${
            showActions
                ? `<td class="sa-actions">
            ${
                editable || deletable
                    ? `${editable ? `<button type="button" class="btn btn-sm" data-sa-edit="${esc(sa.itemId)}" title="Bearbeiten"><i class="bi bi-pencil"></i></button>` : ''}
               ${deletable ? `<button type="button" class="btn btn-sm" data-sa-del="${esc(sa.itemId)}" title="Löschen"><i class="bi bi-trash"></i></button>` : ''}`
                    : '–'
            }
          </td>`
                : ''
        }
      </tr>`;
        })
        .join('')}
    </tbody></table></div>`;
}

function schularbeitStatus(sa) {
    return String((sa && sa.status) || '').toLowerCase();
}

function renderAdminKpiGrid(state) {
    const items = filterSchularbeiten(state.items, state.filters, scopeFromState(state, { scopeAll: true }));
    const kpis = computeDashboardKpis(items, undefined, state.rules);
    return `
    <div class="sa-kpi-grid sa-admin-kpis" role="region" aria-label="Kennzahlen">
      <div class="sa-kpi"><span>Offene Anträge</span><strong>${kpis.offen}</strong></div>
      <div class="sa-kpi"><span>Fixiert (2 Wochen)</span><strong>${kpis.fixiertNaechste2Wochen}</strong></div>
      <div class="sa-kpi${kpis.konflikte ? ' sa-kpi--warn' : ''}"><span>Konflikte</span><strong>${kpis.konflikte}</strong></div>
      <div class="sa-kpi"><span>Diese Woche</span><strong>${kpis.dieseWoche}</strong></div>
    </div>`;
}

function fachMetaPaletteColor(index) {
    const palette = FACH_META_COLOR_PALETTE;
    return palette[((index % palette.length) + palette.length) % palette.length];
}

function renderFachMetaAdminTable(state) {
    const labels = labelMaps(state.stammdaten);
    const subjects = state.stammdaten.subjects || [];
    const saved = state.fachMeta || [];
    const drafts = state.fachMetaDrafts || [];
    const usedCodes = new Set(saved.map((m) => m.fachCode));
    drafts.forEach((d) => {
        if (d.fachCode) usedCodes.add(d.fachCode);
    });

    const savedRows = saved
        .map((m) => {
            const code = m.fachCode || '';
            const farbe = normalizeHexColor(m.farbe) || '#6366f1';
            const pro = m.proSemester != null ? m.proSemester : DEFAULT_FACH_META_PRO_SEMESTER;
            const dauer = m.standardDauer != null ? m.standardDauer : DEFAULT_FACH_META_STANDARD_DAUER;
            return `<tr data-sa-fm-row data-sa-fm-id="${esc(m.itemId)}" data-sa-fm-code="${esc(code)}">
          <td class="sa-fm-fach">${fachChip(code, labels, state.fachMeta)}</td>
          <td>
            <div class="sa-color-input sa-fm-color-input">
              <input type="color" data-sa-fm-field="farbePicker" value="${esc(farbe)}" aria-label="Farbe">
              <input type="text" data-sa-fm-field="farbe" value="${esc(farbe)}" maxlength="20" class="sa-fm-hex">
            </div>
          </td>
          <td><input type="number" class="sa-fm-num" data-sa-fm-field="pro" min="0" max="6" value="${esc(pro)}" aria-label="Anzahl pro Semester"></td>
          <td><input type="number" class="sa-fm-num" data-sa-fm-field="dauer" min="50" max="300" step="50" value="${esc(dauer)}" aria-label="Standarddauer Minuten"></td>
          <td class="sa-fm-actions">
            <button type="button" class="btn btn-sm" data-sa-del-fachmeta="${esc(m.itemId)}" title="Zeile löschen"><i class="bi bi-trash"></i></button>
          </td>
        </tr>`;
        })
        .join('');

    const draftRows = drafts
        .map((d, di) => {
            const farbe = normalizeHexColor(d.farbe) || fachMetaPaletteColor(saved.length + di);
            const pro = d.proSemester != null ? d.proSemester : DEFAULT_FACH_META_PRO_SEMESTER;
            const dauer = d.standardDauer != null ? d.standardDauer : DEFAULT_FACH_META_STANDARD_DAUER;
            const options = subjects
                .filter((s) => s.code && (!usedCodes.has(s.code) || s.code === d.fachCode))
                .map(
                    (s) =>
                        `<option value="${esc(s.code)}"${s.code === d.fachCode ? ' selected' : ''}>${esc(
                            s.name || s.code
                        )} (${esc(s.code)})</option>`
                )
                .join('');
            return `<tr data-sa-fm-row data-sa-fm-draft="${esc(d.draftId)}">
          <td>
            <select data-sa-fm-field="code" class="sa-fm-select" aria-label="Fach wählen">
              <option value="">Fach wählen …</option>
              ${options}
            </select>
          </td>
          <td>
            <div class="sa-color-input sa-fm-color-input">
              <input type="color" data-sa-fm-field="farbePicker" value="${esc(farbe)}" aria-label="Farbe">
              <input type="text" data-sa-fm-field="farbe" value="${esc(farbe)}" maxlength="20" class="sa-fm-hex">
            </div>
          </td>
          <td><input type="number" class="sa-fm-num" data-sa-fm-field="pro" min="0" max="6" value="${esc(pro)}"></td>
          <td><input type="number" class="sa-fm-num" data-sa-fm-field="dauer" min="50" max="300" step="50" value="${esc(dauer)}"></td>
          <td class="sa-fm-actions">
            <button type="button" class="btn btn-sm alt" data-sa-fm-cancel-draft="${esc(d.draftId)}" title="Entfernen"><i class="bi bi-x-lg"></i></button>
          </td>
        </tr>`;
        })
        .join('');

    const body =
        savedRows || draftRows
            ? savedRows + draftRows
            : `<tr><td colspan="5" class="muted">Noch keine Fächer – mit <strong>+</strong> hinzufügen.</td></tr>`;

    return `
        <div class="sa-fm-table-toolbar">
          <button type="button" class="btn btn-sm" id="saBtnFachMetaAdd" title="Weiteres Fach aus Stammdaten">
            <i class="bi bi-plus-lg"></i> Fach hinzufügen
          </button>
        </div>
        <div class="sa-table-wrap">
          <table class="sa-table sa-fm-table">
            <thead>
              <tr>
                <th>Fach</th>
                <th>Farbe</th>
                <th>Anzahl / Semester</th>
                <th>Standarddauer (Min.)</th>
                <th></th>
              </tr>
            </thead>
            <tbody>${body}</tbody>
          </table>
        </div>
        <p class="muted" style="margin:8px 0 0;font-size:0.88em;">Änderungen werden beim Verlassen des Feldes gespeichert. Neue Zeilen: Fach wählen, dann Werte eintragen.</p>`;
}

function renderAdminListResetPanel(state) {
    const hasCtx = !!(state.ctx && state.ctx.lists);
    const rows = LIST_KEYS.map((key) => {
        const title = LIST_TITLES[key] || key;
        const desc = LIST_DESCRIPTIONS[key] || '';
        return `
        <label class="sa-reset-row">
          <input type="checkbox" class="sa-reset-check" name="saClearList" value="${esc(key)}" checked ${hasCtx ? '' : 'disabled'}>
          <span class="sa-reset-row__text">
            <strong>${esc(title)}</strong>
            <span class="muted">${esc(desc)}</span>
          </span>
          <button type="button" class="btn btn-sm alt" data-sa-clear-one="${esc(key)}" ${hasCtx ? '' : 'disabled'} title="Nur diese Liste leeren">Leeren</button>
        </label>`;
    }).join('');

    return `
    <section class="tm-panel sa-reset-panel">
      <div class="tm-panel__head"><i class="bi bi-trash3"></i>
        <div><h3>Listen zurücksetzen</h3><p>Alle Einträge löschen – die Listen und Spalten bleiben erhalten</p></div>
      </div>
      ${
          hasCtx
              ? ''
              : '<p class="muted">Bitte oben die SharePoint-Site laden, bevor Listen geleert werden können.</p>'
      }
      <div class="sa-reset-grid">${rows}</div>
      <div class="sa-reset-actions">
        <button type="button" class="btn alt" id="saBtnClearSelected" ${hasCtx ? '' : 'disabled'}>
          <i class="bi bi-check2-square"></i>Ausgewählte Listen leeren
        </button>
        <button type="button" class="btn btn-danger" id="saBtnClearAllPlaner" ${hasCtx ? '' : 'disabled'}>
          <i class="bi bi-exclamation-triangle"></i>Alle Planer-Listen leeren
        </button>
      </div>
      ${
          state.localDemoOnly
              ? '<button type="button" class="btn btn-sm alt" id="saBtnClearLocalDemo" style="margin-top:10px;"><i class="bi bi-display"></i>Lokale Demo-Anzeige zurücksetzen</button>'
              : ''
      }
      <pre id="saResetLog" class="sa-reset-log" hidden aria-live="polite"></pre>
      <p class="muted sa-reset-hint">
        Schultermine, Outlook- und Teams-Kalendereinträge werden dabei nicht automatisch entfernt.
        Für einen Neustart: Listen leeren, danach JSON erneut importieren.
      </p>
    </section>`;
}

function renderAdminSitePanel(state) {
    const connected = !!(state.bootstrapped && state.ctx);
    return `
    <section class="tm-panel sa-admin-site">
      <div class="tm-panel__head"><i class="bi bi-cloud-arrow-up"></i>
        <div><h3>SharePoint-Verbindung</h3>
        <p>Datenquelle für Lehrer:innen und Schüler:innen (URL wird lokal im Browser gespeichert).</p></div>
      </div>
      <div class="sa-admin-site__row">
        <div class="tm-field sa-admin-site__url">
          <label for="saSiteUrl">Website-URL</label>
          <input type="url" id="saSiteUrl" value="${esc(state.siteUrl)}" placeholder="https://…sharepoint.com/sites/administration" spellcheck="false" autocomplete="off">
        </div>
        <div class="sa-admin-site__actions">
          <button type="button" class="btn" id="saBtnLoad"><i class="bi bi-arrow-repeat"></i>Laden</button>
          <label class="btn" for="saImportJson"><i class="bi bi-filetype-json"></i>JSON importieren</label>
          <input type="file" id="saImportJson" accept=".json,application/json" hidden>
          <a class="btn" href="sharepoint-liste-schularbeiten.html" style="text-decoration:none"><i class="bi bi-list-ul"></i>Listen</a>
        </div>
      </div>
      ${
          connected
              ? `<p class="sa-admin-site__status ok"><i class="bi bi-check-circle"></i> Verbunden: <code>${esc(state.siteUrl)}</code></p>`
              : '<p class="sa-admin-site__status muted">Noch nicht geladen – bitte URL prüfen und „Laden“ wählen.</p>'
      }
      ${
          state.localDemoOnly
              ? '<p class="sa-import-hint">Lokaler Demo-Import aktiv – Anzeige ohne SharePoint. Zum Schreiben: Site laden + erneut importieren und „Auf SharePoint schreiben“ wählen.</p>'
              : ''
      }
    </section>`;
}

function renderAdmin(state) {
    if (state.role !== 'admin') {
        return `<section class="tm-panel"><p>Nur für Administrator:innen.</p></section>`;
    }
    const labels = labelMaps(state.stammdaten);
    const open = state.items.filter((s) => s.status === 'beantragt');
    const rw = state.rules;

    return `
    ${renderAdminSitePanel(state)}
    ${renderAdminKpiGrid(state)}
    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-inbox"></i>
        <div><h3>Offene Anträge</h3><p>Fixieren oder ablehnen</p></div>
      </div>
      ${
          open.length
              ? `<div class="sa-table-wrap"><table class="sa-table"><thead><tr>
          <th>Datum</th><th>Fach</th><th>Klasse</th><th>Lehrer:in</th><th>Titel</th><th>Status</th><th>Aktionen</th>
        </tr></thead><tbody>
        ${open
            .map(
                (sa) => `<tr class="sa-row" style="--sa-fach:${esc(fachColor(sa.fachCode, -1, state.fachMeta))}">
          <td>${esc(formatDeDate(sa.datum))}</td>
          <td>${fachChip(sa.fachCode, labels, state.fachMeta)}</td>
          <td>${esc(labels.klasse[sa.klasseCode] || sa.klasseCode)}</td>
          <td>${esc(labels.lehrer[sa.lehrerCode] || sa.lehrerCode)}</td>
          <td>${esc(schularbeitDisplayTitle(sa))}</td>
          <td>${statusBadge(sa.status)}</td>
          <td class="sa-actions">
            <button type="button" class="btn btn-sm btn-success" data-sa-fix="${esc(sa.itemId)}" title="Fixieren"><i class="bi bi-check-lg"></i></button>
            <button type="button" class="btn btn-sm" data-sa-reject="${esc(sa.itemId)}" title="Ablehnen"><i class="bi bi-x-lg"></i></button>
          </td>
        </tr>`
            )
            .join('')}
        </tbody></table></div>`
              : '<p class="muted">Keine offenen Anträge.</p>'
      }
    </section>

    <div class="sa-split">
      <section class="tm-panel">
        <div class="tm-panel__head"><i class="bi bi-slash-circle"></i>
          <div><h3>Sperrzeiten (Terminfenster)</h3><p>Ferien, Sportwoche, …</p></div>
        </div>
        <div class="sa-form">
          <div class="tm-field"><label for="saTfTitel">Titel</label><input id="saTfTitel" type="text" placeholder="z. B. Herbstferien"></div>
          <div class="sa-form__row">
            <div class="tm-field"><label for="saTfVon">Von</label><input id="saTfVon" type="date"></div>
            <div class="tm-field"><label for="saTfBis">Bis</label><input id="saTfBis" type="date"></div>
          </div>
          <div class="tm-field"><label for="saTfDesc">Beschreibung</label><textarea id="saTfDesc" rows="2"></textarea></div>
          <button type="button" class="btn" id="saBtnAddFenster"><i class="bi bi-plus"></i>Sperrzeit hinzufügen</button>
        </div>
        <ul class="sa-fenster-list">
          ${
              state.windows.length
                  ? state.windows
                        .map(
                            (w) => `<li>
              <div><strong>${esc(w.titel)}</strong>
              <span>${esc(formatDeDate(w.startdatum))} – ${esc(formatDeDate(w.enddatum))} · ${esc(w.typ)}</span></div>
              <button type="button" class="btn btn-sm" data-sa-del-fenster="${esc(w.itemId)}" title="Löschen"><i class="bi bi-trash"></i></button>
            </li>`
                        )
                        .join('')
                  : '<li class="muted">Noch keine Terminfenster.</li>'
          }
        </ul>
      </section>

      <section class="tm-panel">
        <div class="tm-panel__head"><i class="bi bi-sliders"></i>
          <div><h3>Regelwerk</h3><p>Basis: SchUG § 17 / LBVO § 7</p></div>
        </div>
        <div class="sa-form">
          <div class="tm-field"><label for="saRwName">Name</label><input id="saRwName" type="text" value="${esc(rw.name || '')}"></div>
          <div class="sa-form__row">
            <div class="tm-field"><label for="saRwTag">Max. pro Tag</label><input id="saRwTag" type="number" min="1" max="3" value="${esc(rw.maxProTag)}"></div>
            <div class="tm-field"><label for="saRwWoche">Max. pro Woche</label><input id="saRwWoche" type="number" min="1" max="5" value="${esc(rw.maxProWoche)}"></div>
          </div>
          <div class="sa-form__row">
            <div class="tm-field"><label for="saRwFrist">Ankündigungsfrist (Tage)</label><input id="saRwFrist" type="number" min="1" max="30" value="${esc(rw.ankuendigungsfristTage)}"></div>
            <div class="tm-field"><label for="saRwSperre">Sperre vor Notenkonferenz</label><input id="saRwSperre" type="number" min="0" max="21" value="${esc(rw.sperreVorNotenkonferenzTage)}"></div>
          </div>
          <button type="button" class="btn" id="saBtnSaveRules"><i class="bi bi-save"></i>Speichern</button>
        </div>
      </section>
    </div>

    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-palette"></i>
        <div><h3>SAP-FachMeta</h3><p>Farbe, Kontingent und Standarddauer (optional, je Schuljahr)</p></div>
      </div>
      ${
          state.ctx && state.ctx.lists && state.ctx.lists.fachMeta
              ? renderFachMetaAdminTable(state)
              : '<p class="muted">Liste SAP-FachMeta fehlt. <a href="sharepoint-liste-schularbeiten.html">Listen-Paket erneut ausführen</a>.</p>'
      }
    </section>

    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-people"></i>
        <div><h3>SharePoint-Berechtigungen</h3>
          <p>Entra-Gruppen auf allen Planer-Listen (Vererbung wird gebrochen)</p></div>
      </div>
      ${(() => {
          const p = loadPermissionsConfig();
          return `
      <p class="muted" style="font-size:0.88em;line-height:1.45;margin:0 0 10px;">
        Empfohlen auf einer <strong>Administrations-Site</strong> (z. B. <code>/sites/administration</code>).
        Schularbeiten: Lehrer <em>Beitragen</em>, Schüler <em>Lesen</em>, Verwaltung <em>Vollzugriff</em>.
        Regelwerk/Terminfenster/FachMeta: Lehrer Lesen, Schüler kein Zugriff.
        Voraussetzung: <code>Sites.FullControl.All</code> und <code>Group.Read.All</code>.
      </p>
      ${htmlSchularbeitenEntraPermGrid(p, { delegated: true, mode: 'planer' })}
      <div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:10px;">
        <button type="button" class="btn" id="saBtnSavePerms"><i class="bi bi-save"></i>Gruppen speichern</button>
        <button type="button" class="btn btn-primary" id="saBtnApplyPerms"><i class="bi bi-shield-lock"></i>Berechtigungen anwenden</button>
      </div>
      <p class="muted" style="margin:10px 0 0;font-size:0.88em;">
        Gleiche Einstellungen wie im Tool
        <a href="sharepoint-liste-schularbeiten.html">Schularbeiten-Listen</a>.
        Site muss unter „SharePoint-Verbindung“ geladen sein (${esc(state.siteUrl || '–')}).
      </p>`;
      })()}
    </section>

    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-sliders2"></i>
        <div><h3>Automatisierung</h3><p>Lokal im Browser gespeichert</p></div>
      </div>
      <label class="sa-check">
        <input type="checkbox" id="saSyncTermine"${state.settings && state.settings.syncSchultermine ? ' checked' : ''}>
        Bei Fixierung Eintrag in Schultermine schreiben (Kategorie Prüfung)
      </label>
      <label class="sa-check" style="margin-top:8px;">
        <input type="checkbox" id="saSyncClassCalendar"${state.settings && state.settings.syncClassCalendar ? ' checked' : ''}>
        Bei Fixierung Termin in den Klassen-Teams-Kalender schreiben
      </label>
      <div class="tm-field" style="margin-top:10px;max-width:320px;">
        <label for="saSyncList">Schultermine-Listenname</label>
        <input id="saSyncList" type="text" value="${esc((state.settings && state.settings.schultermineList) || 'Schultermine')}">
      </div>
      <div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:10px;">
        <button type="button" class="btn" id="saBtnSaveSettings"><i class="bi bi-save"></i>Einstellungen speichern</button>
        <button type="button" class="btn" id="saBtnSyncClassCal" title="Alle fixierten Schularbeiten in die jeweiligen Klassenkalender schreiben">
          <i class="bi bi-calendar-plus"></i>Fixierte in Klassenkalender schreiben
        </button>
      </div>
      <p class="muted" style="margin:10px 0 0;font-size:0.88em;">
        Klassenkalender braucht verknüpfte Teams-Gruppen in den
        <a href="../tenant.html">Stammdaten</a> (Klassen-Teams / Gruppenabgleich)
        und Graph-Recht <code>Group.ReadWrite.All</code>. Spalte
        <code>TeamsCalendarEventId</code> ggf. über
        <a href="sharepoint-liste-schularbeiten.html">Listen einrichten</a> nachziehen.
      </p>
      <p class="muted" style="margin:10px 0 0;font-size:0.88em;">
        Danach optional Flow <a href="pa-termine-sync.html">Termine → Kalender</a> und
        <a href="pa-schularbeiten-mail.html">Status-Mail</a>.
      </p>
      <div class="sa-import-box" style="margin-top:14px;padding-top:12px;border-top:1px solid var(--border);">
        <p style="margin:0 0 8px;font-size:0.92rem;"><strong>Demo-JSON importieren</strong> (z. B. <code>docs/demo-data/schularbeiten-2026-27.json</code>)</p>
        <label class="btn sa-import-btn" for="saImportJsonAdmin">
          <i class="bi bi-upload"></i>JSON-Datei wählen
        </label>
        <input type="file" id="saImportJsonAdmin" accept=".json,application/json" hidden>
      </div>
    </section>
    ${renderAdminListResetPanel(state)}
    ${renderIntranetPanel(state)}`;
}

function calEventsHtml(list, state, labels, canDrag, calMode) {
    const weekLayout = calMode === 'week';
    return sortSchularbeitenByBeginn(list)
        .map((sa) => {
            const color = fachColor(sa.fachCode, -1, state.fachMeta);
            const fg = contrastOn(color);
            const fachLabel = labels.fach[sa.fachCode] || sa.fachCode || '?';
            const title = schularbeitDisplayTitle(sa);
            const zeitKurz = formatSchularbeitZeitKurz(sa);
            const zeitLabel = zeitKurz || '–';
            const drag =
                canDrag && (sa.status === 'beantragt' || sa.status === 'fixiert' || sa.status === 'abgelehnt');
            const tip =
                fachLabel +
                ' · ' +
                (formatSchularbeitZeitspanne(sa) || 'ohne Uhrzeit') +
                ' · ' +
                title +
                ' · ' +
                statusLabel(sa.status) +
                (drag ? ' · Ziehen zum Verschieben' : '');
            const timeHtml = weekLayout
                ? `<span class="sa-cal__ev-time">${esc(zeitLabel)}</span>`
                : `<span class="sa-cal__ev-time sa-cal__ev-time--compact">${esc(zeitKurz ? zeitKurz.split('–')[0] : '–')}</span>`;
            return `<button type="button" class="sa-cal__ev sa-cal__ev--${esc(sa.status)}${
                weekLayout ? ' sa-cal__ev--week' : ''
            }${drag ? ' sa-cal__ev--draggable' : ''}" style="--sa-fach:${esc(color)};--sa-fach-fg:${esc(fg)}"${
                drag ? ` draggable="true" data-sa-cal-drag="${esc(sa.itemId)}"` : ''
            } data-sa-detail="${esc(sa.itemId)}" title="${esc(tip)}">${timeHtml}<span class="sa-cal__ev-title">${esc(
                title
            )}</span></button>`;
        })
        .join('');
}

function calCellHtml(iso, dayLabel, list, state, todayIso, canDrag, labels, calMode) {
    const blocked = (state.windows || []).some((w) => w.typ === 'gesperrt' && dateInWindow(iso, w));
    return `<div class="sa-cal__cell${iso === todayIso ? ' is-today' : ''}${blocked ? ' is-blocked' : ''}" data-sa-cal-drop="${esc(
        iso
    )}"${canDrag ? ' data-sa-cal-droppable="1"' : ''}>
          <div class="sa-cal__day">${esc(String(dayLabel))}</div>
          <div class="sa-cal__events">${calEventsHtml(list, state, labels, canDrag, calMode)}</div>
        </div>`;
}

function calDayLabel(iso, calMode) {
    const dayNum = parseInt(String(iso).slice(8, 10), 10);
    if (calMode !== 'week') return String(dayNum);
    const p = String(iso).split('-').map((x) => parseInt(x, 10));
    if (p.length < 3) return String(dayNum);
    const d = new Date(p[0], p[1] - 1, p[2]);
    const wd = d.toLocaleDateString('de-AT', { weekday: 'short' });
    return `${wd} ${dayNum}.`;
}

function resolveCalWeekMonday(state) {
    const today = toIsoDateOnly(new Date());
    const fromState = String(state.calWeekMonday || '').trim().slice(0, 10);
    if (fromState) return mondayOfWeekContaining(fromState) || fromState;
    const anchor = `${state.calYear}-${String(state.calMonth + 1).padStart(2, '0')}-15`;
    return mondayOfWeekContaining(anchor) || mondayOfWeekContaining(today) || today;
}

function renderKalender(state) {
    const labels = labelMaps(state.stammdaten);
    const isAdmin = state.role === 'admin';
    const canDrag = isAdmin;
    const baseItems =
        isAdmin
            ? filterSchularbeiten(state.items, state.filters, scopeFromState(state, { scopeAll: true }))
            : scopedItems(state, state.role !== 'schueler');
    const items = filterItemsForCalendarView(baseItems, state);
    const calShow = state.calShow || { beantragt: true, fixiert: true, abgelehnt: false };
    const calMode = state.calMode === 'week' ? 'week' : 'month';

    const y = state.calYear;
    const m = state.calMonth;
    const first = new Date(y, m, 1);
    const startPad = (first.getDay() + 6) % 7; // Montag = 0
    const daysInMonth = new Date(y, m + 1, 0).getDate();
    const monthName = first.toLocaleDateString('de-AT', { month: 'long', year: 'numeric' });
    const today = toIsoDateOnly(new Date());
    const weekMonday = resolveCalWeekMonday(state);
    const weekSunday = addDays(weekMonday, 6) || weekMonday;
    const weekTitle =
        formatDeDate(weekMonday) + ' – ' + formatDeDate(weekSunday) + ' · ' + String(weekMonday).slice(0, 4);
    const periodTitle = calMode === 'week' ? weekTitle : monthName;

    const byDay = {};
    items.forEach((sa) => {
        const d = String(sa.datum || '').slice(0, 10);
        if (!d) return;
        if (!byDay[d]) byDay[d] = [];
        byDay[d].push(sa);
    });
    Object.keys(byDay).forEach((d) => {
        byDay[d] = sortSchularbeitenByBeginn(byDay[d]);
    });

    let gridClass = 'sa-cal__grid';
    const cells = [];
    if (calMode === 'week') {
        gridClass += ' sa-cal__grid--week';
        for (let i = 0; i < 7; i++) {
            const iso = addDays(weekMonday, i) || '';
            if (!iso) continue;
            cells.push(calCellHtml(iso, calDayLabel(iso, 'week'), byDay[iso] || [], state, today, canDrag, labels, 'week'));
        }
    } else {
        for (let i = 0; i < startPad; i++) cells.push('<div class="sa-cal__cell sa-cal__cell--empty"></div>');
        for (let day = 1; day <= daysInMonth; day++) {
            const iso = `${y}-${String(m + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
            cells.push(calCellHtml(iso, calDayLabel(iso, 'month'), byDay[iso] || [], state, today, canDrag, labels, 'month'));
        }
    }

    const navPrevLabel = calMode === 'week' ? 'Vorherige Woche' : 'Vorheriger Monat';
    const navNextLabel = calMode === 'week' ? 'Nächste Woche' : 'Nächster Monat';

    return `
    ${filterBar(state, { scopeAll: state.role === 'admin' })}
    <section class="tm-panel">
      <div class="sa-cal__head">
        <h3>${esc(periodTitle)}</h3>
        <div class="sa-cal__nav">
          <div class="sa-cal__mode" role="group" aria-label="Kalenderansicht">
            <button type="button" class="${calMode === 'month' ? 'is-active' : ''}" id="saCalModeMonth" aria-pressed="${calMode === 'month'}">Monat</button>
            <button type="button" class="${calMode === 'week' ? 'is-active' : ''}" id="saCalModeWeek" aria-pressed="${calMode === 'week'}">Woche</button>
          </div>
          <button type="button" class="btn" id="saCalPrev" aria-label="${esc(navPrevLabel)}"><i class="bi bi-chevron-left"></i></button>
          <button type="button" class="btn" id="saCalToday">Heute</button>
          <button type="button" class="btn" id="saCalNext" aria-label="${esc(navNextLabel)}"><i class="bi bi-chevron-right"></i></button>
        </div>
      </div>
      ${
          isAdmin
              ? `<div class="sa-cal__admin-bar" role="group" aria-label="Kalender anzeigen">
          <span class="sa-cal__admin-label">Anzeigen:</span>
          <label class="sa-cal__check"><input type="checkbox" id="saCalShowBeantragt"${calShow.beantragt ? ' checked' : ''}> Geplant (beantragt)</label>
          <label class="sa-cal__check"><input type="checkbox" id="saCalShowFixiert"${calShow.fixiert ? ' checked' : ''}> Fixiert</label>
          <label class="sa-cal__check"><input type="checkbox" id="saCalShowAbgelehnt"${calShow.abgelehnt ? ' checked' : ''}> Abgelehnt</label>
          <span class="muted sa-cal__drag-hint"><i class="bi bi-arrows-move"></i> Termine per Drag &amp; Drop auf einen Tag verschieben</span>
        </div>`
              : ''
      }
      <div class="sa-cal__weekdays"><span>Mo</span><span>Di</span><span>Mi</span><span>Do</span><span>Fr</span><span>Sa</span><span>So</span></div>
      <div class="${gridClass}">${cells.join('')}</div>
      <div class="sa-cal__legend">
        <span><i class="sa-dot sa-dot--warn"></i>Beantragt (gestrichelt)</span>
        <span><i class="sa-dot sa-dot--ok"></i>Fixiert</span>
        <span><i class="sa-dot sa-dot--bad"></i>Abgelehnt (blass)</span>
        <span><i class="sa-blocked-swatch"></i>Gesperrter Tag</span>
      </div>
    </section>`;
}

function renderRegeln(state) {
    const r = state.rules || {};
    const cards = [
        {
            t: `Max. ${r.maxProTag || 1} pro Tag / max. ${r.maxProWoche || 2} pro Woche`,
            d: 'Pro Klasse: Obergrenze Schularbeiten je Tag und Kalenderwoche.',
            c: '§ 7 Abs. 8 LBVO'
        },
        {
            t: 'Ankündigungsfrist',
            d: `Mindestens ${r.ankuendigungsfristTage || 7} Tage vor dem Termin.`,
            c: '§ 7 Abs. 1 LBVO'
        },
        {
            t: 'Kein Termin nach schulfreien Tagen',
            d: 'Warnung, wenn der Vortag in einer Sperrzeit liegt.',
            c: '§ 7 Abs. 8 LBVO'
        },
        {
            t: 'Sperrzeit vor Notenkonferenz',
            d: `Konfigurierbar (${r.sperreVorNotenkonferenzTage || 7} Tage) bzw. über Terminfenster.`,
            c: 'Schulautonom'
        },
        {
            t: 'Anzahl pro Fach und Semester',
            d: 'Typisch 2 – Warnung bei Überschreitung (Kontingent).',
            c: 'HAK Lehrplan'
        },
        {
            t: 'Dauer',
            d: 'Üblich 50–150 Minuten je nach Jahrgang/Fach.',
            c: 'HAK Lehrplan'
        }
    ];
    return `
    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-book"></i>
        <div><h3>Regelwerk</h3><p>Grundlage SchUG § 17 und LBVO § 7 (sinngemäß für HAK)</p></div>
      </div>
      <div class="sa-rule-cards">
        ${cards
            .map(
                (c) => `<article class="sa-rule-card">
          <h4>${esc(c.t)}</h4><p>${esc(c.d)}</p><span>${esc(c.c)}</span>
        </article>`
            )
            .join('')}
      </div>
      <div class="sa-alert sa-alert--info" style="margin-top:16px;">
        <strong>Fehler</strong> blockieren das Einreichen. <strong>Warnungen</strong> sind erlaubt, werden aber hervorgehoben.
        Administrator:innen pflegen Sperrzeiten und Zahlenwerte unter Administration.
      </div>
    </section>`;
}

function personalCalendarSyncItems(state) {
    const scope = scopeFromState(state, { onlyMine: true });
    return filterSchularbeiten(state.items, state.filters, scope).filter((s) => {
        const st = String(s.status || '').toLowerCase();
        return st === 'fixiert' || st === 'beantragt';
    });
}

function renderPersonalCalendarCard(state, syncableItems) {
    if (state.role === 'schueler') return '';
    const email = String(state.accountEmail || '').trim();
    const n = syncableItems.length;
    const linked = countPersonalCalendarLinked(syncableItems, email);
    let hint = '';
    let disabled = !email || n === 0;
    if (!email) {
        hint = 'Bitte oben rechts mit Microsoft anmelden.';
    } else if (n === 0) {
        hint =
            'Keine eigenen Termine (fixiert oder beantragt) im aktuellen Filter. Stammdaten: Lehrer:in mit Ihrer E-Mail oder Antrag mit beantragtVon.';
    } else {
        hint =
            'Angemeldet als ' +
            email +
            ' · ' +
            n +
            ' Termine werden geschrieben (fixiert + beantragt).' +
            (linked ? ' ' + linked + ' sind auf diesem Gerät bereits verknüpft (Update statt Duplikat).' : '') +
            ' Berechtigung: Calendars.ReadWrite.';
    }

    return `
      <section class="tm-panel sa-export-card sa-export-card--wide">
        <h3><i class="bi bi-calendar-check"></i> Mein Outlook-Kalender</h3>
        <p>
          Trägt <strong>Ihre</strong> Schularbeiten aus dem Filter in den <strong>persönlichen</strong> Microsoft-365-Kalender ein
          (Outlook / Teams) – ohne Meeting-Einladungen und ohne Erinnerungen. Verknüpfung pro Gerät im Browser gespeichert.
        </p>
        <p class="muted" style="margin:0 0 10px;font-size:0.9rem;line-height:1.45;">${esc(hint)}</p>
        <button type="button" class="btn" data-sa-sync-my-cal${disabled ? ' disabled' : ''}>
          <i class="bi bi-calendar-plus"></i>In meinen Kalender schreiben (${n})
        </button>
      </section>`;
}

function renderExportGroupCalendarCard(state, fixiertItems, labels) {
    if (state.role !== 'admin') return '';
    const klasse = String((state.filters && state.filters.klasse) || '').trim();
    const n = fixiertItems.length;
    const classCount = new Set(fixiertItems.map((s) => s.klasseCode).filter(Boolean)).size;
    const synced = fixiertItems.filter((s) => s.teamsCalendarEventId).length;
    let hint = '';
    let hintClass = 'muted';
    let disabled = !state.ctx || n === 0;

    if (!state.ctx) {
        hint = 'SharePoint-Site unter Administration laden.';
    } else if (n === 0) {
        hint =
            'Keine fixierten Termine in der Vorschau. Status „Fixiert“ setzen oder Filter anpassen (iCal/Druck zeigen auch beantragte Termine).';
    } else if (klasse) {
        const link = describeClassGroupLink(klasse);
        hint = link.message;
        hintClass = link.ok ? 'sa-export-cal-ok' : 'sa-export-cal-warn';
        if (!link.ok) disabled = true;
        else {
            hint +=
                ' · ' +
                n +
                ' fixierte Termine' +
                (synced ? ' (' + synced + ' bereits im Kalender verknüpft)' : '');
        }
    } else {
        hint =
            n +
            ' fixierte Termine über ' +
            classCount +
            ' Klassen – es wird je Klasse der passende Gruppenkalender verwendet. Optional Klasse oben im Filter wählen, um nur eine Gruppe zu beschreiben.' +
            (synced ? ' (' + synced + ' bereits verknüpft)' : '');
    }

    const btnLabel = klasse
        ? 'In Gruppenkalender schreiben (' + (labels.klasse[klasse] || klasse) + ')'
        : 'In Klassen-Gruppenkalender schreiben (' + n + ')';

    return `
      <section class="tm-panel sa-export-card sa-export-card--wide">
        <h3><i class="bi bi-people"></i> Microsoft 365 Gruppenkalender</h3>
        <p>
          Schreibt <strong>nur fixierte</strong> Termine aus der Vorschau in den Outlook-Kalender der
          verknüpften Klassen-Teams-Gruppe (Graph) – <strong>ohne Einladungen</strong> an Mitglieder und
          <strong>ohne Erinnerungen</strong>. Stammdaten:
          <a href="../tenant.html">Klassen-Teams / Gruppenabgleich</a>.
          Berechtigung: <code>Group.ReadWrite.All</code>.
        </p>
        <p class="${hintClass}" style="margin:0 0 10px;font-size:0.9rem;line-height:1.45;">${esc(hint)}</p>
        <button type="button" class="btn" id="saBtnExportGroupCal"${disabled ? ' disabled' : ''}>
          <i class="bi bi-calendar-plus"></i>${esc(btnLabel)}
        </button>
      </section>`;
}

function renderExport(state) {
    const labels = labelMaps(state.stammdaten);
    const items =
        state.role === 'admin'
            ? filterSchularbeiten(state.items, state.filters, scopeFromState(state, { scopeAll: true }))
            : scopedItems(state, state.role !== 'schueler');
    const fixed =
        state.role === 'schueler'
            ? items
            : items.filter((s) => s.status === 'fixiert' || s.status === 'beantragt');
    const fixiertOnly = items.filter((s) => String(s.status || '').toLowerCase() === 'fixiert');
    const ownSyncable = personalCalendarSyncItems(state);

    return `
    ${filterBar(state, { scopeAll: state.role === 'admin' })}
    <div class="sa-export-cards">
      <section class="tm-panel sa-export-card">
        <h3><i class="bi bi-calendar-plus"></i> iCal (.ics)</h3>
        <p>Import in Outlook, Google Calendar, Apple Kalender.</p>
        <button type="button" class="btn" id="saBtnIcs"><i class="bi bi-download"></i>.ics herunterladen</button>
      </section>
      <section class="tm-panel sa-export-card">
        <h3><i class="bi bi-printer"></i> Drucken / PDF</h3>
        <p>Druckdialog – dort „Als PDF speichern“ wählen.</p>
        <button type="button" class="btn" id="saBtnPrint"><i class="bi bi-printer"></i>Drucken</button>
      </section>
      ${renderPersonalCalendarCard(state, ownSyncable)}
      ${renderExportGroupCalendarCard(state, fixiertOnly, labels)}
    </div>
    <section class="tm-panel sa-print-area" id="saPrintArea">
      <div class="tm-panel__head"><i class="bi bi-eye"></i>
        <div><h3>Vorschau (${fixed.length} Termine, davon ${fixiertOnly.length} fixiert)</h3><p>${esc(printTableCaption(fixed))}</p></div>
      </div>
      ${tableSchularbeiten(fixed, labels, { showActions: false, fachMeta: state.fachMeta })}
    </section>`;
}

function renderDetailModal(state) {
    const sa = state.items.find((x) => x.itemId === state.detailId);
    if (!sa) return '';
    const labels = labelMaps(state.stammdaten);
    const scope = scopeFromState(state);
    const canEdit = canEditSchularbeit(sa, scope);
    const canDel = canDeleteSchularbeit(sa, scope);
    const canDecide = canAdminDecide(sa, scope);
    const check = validateSchularbeit({
        draft: sa,
        existing: state.items,
        rules: state.rules,
        windows: state.windows
    });

    const fachLabel = labels.fach[sa.fachCode] || sa.fachCode || '—';
    const klasseLabel = labels.klasse[sa.klasseCode] || sa.klasseCode || '—';
    const lehrerLabel = labels.lehrer[sa.lehrerCode] || sa.lehrerCode || '';
    const zeitLabel = formatSchularbeitZeitspanne(sa) || 'Uhrzeit fehlt';
    const semLabel = sa.semester === 'SS' ? 'Sommersemester' : sa.semester === 'WS' ? 'Wintersemester' : '';

    return `
    <div class="sa-modal" role="dialog" aria-modal="true" aria-labelledby="saDetailTitle">
      <div class="sa-modal__card">
        <header class="sa-modal__head">
          <div class="sa-modal__head-main">
            <h3 id="saDetailTitle">${esc(schularbeitDisplayTitle(sa))}</h3>
            <p class="sa-modal__sub">${fachChip(sa.fachCode, labels, state.fachMeta)} <span class="sa-modal__sub-sep">·</span> ${esc(klasseLabel)}</p>
          </div>
          <div class="sa-modal__head-side">
            ${statusBadge(sa.status)}
            <button type="button" class="btn btn-sm sa-modal__icon-close" id="saDetailClose" aria-label="Schließen"><i class="bi bi-x-lg"></i></button>
          </div>
        </header>
        <div class="sa-modal__when" role="group" aria-label="Termin">
          <div class="sa-modal__when-item">
            <span class="sa-modal__when-label">Datum</span>
            <strong>${esc(formatDeDate(sa.datum))}</strong>
          </div>
          <div class="sa-modal__when-item">
            <span class="sa-modal__when-label">Zeit</span>
            <strong>${esc(zeitLabel)}</strong>
          </div>
          <div class="sa-modal__when-item">
            <span class="sa-modal__when-label">Dauer</span>
            <strong>${esc(sa.dauerMinuten)} Min.</strong>
          </div>
        </div>
        <dl class="sa-dl sa-dl--detail">
          <div><dt>Fach</dt><dd>${esc(fachLabel)}</dd></div>
          <div><dt>Klasse</dt><dd>${esc(klasseLabel)}</dd></div>
          <div><dt>Lehrer:in</dt><dd>${esc(lehrerLabel || '—')}</dd></div>
          ${semLabel ? `<div><dt>Semester</dt><dd>${esc(semLabel)}</dd></div>` : ''}
          ${sa.thema ? `<div class="sa-dl__full"><dt>Thema</dt><dd>${esc(sa.thema)}</dd></div>` : ''}
        </dl>
        ${sa.notiz ? `<div class="sa-modal__note"><span class="sa-modal__note-label">Notiz</span><p>${esc(sa.notiz)}</p></div>` : ''}
        ${sa.ablehnungsGrund ? `<div class="sa-modal__note sa-modal__note--warn"><span class="sa-modal__note-label">Ablehnung</span><p>${esc(sa.ablehnungsGrund)}</p></div>` : ''}
        ${
            check.errors.length || check.warnings.length
                ? `<div class="sa-detail-rules">
            ${check.errors.map((e) => `<p class="sa-rule-list--err">⚠ ${esc(e)}</p>`).join('')}
            ${check.warnings.map((w) => `<p class="sa-rule-list--warn">ℹ ${esc(w)}</p>`).join('')}
          </div>`
                : ''
        }
        <div class="sa-modal__actions">
          ${
              canDecide
                  ? `<button type="button" class="btn btn-sm btn-success" data-sa-fix="${esc(sa.itemId)}"><i class="bi bi-check-lg"></i>Fixieren</button>
             <button type="button" class="btn btn-sm" data-sa-reject="${esc(sa.itemId)}"><i class="bi bi-x-lg"></i>Ablehnen</button>`
                  : ''
          }
          ${
              canEdit
                  ? `<button type="button" class="btn btn-sm" data-sa-edit="${esc(sa.itemId)}"><i class="bi bi-pencil"></i>Bearbeiten</button>`
                  : ''
          }
          ${
              canDel
                  ? `<button type="button" class="btn btn-sm" data-sa-del="${esc(sa.itemId)}"><i class="bi bi-trash"></i>Löschen</button>`
                  : ''
          }
          <button type="button" class="btn btn-sm sa-modal__btn-close" id="saDetailClose2">Schließen</button>
        </div>
      </div>
    </div>`;
}

/**
 * Formulardaten aus dem DOM lesen (View neu).
 */
export function readFormFromDom() {
    const val = (id) => {
        const el = document.getElementById(id);
        return el ? String(el.value || '').trim() : '';
    };
    const teacherSel = document.getElementById('saLehrer');
    let lehrerEmail = '';
    if (teacherSel && teacherSel.selectedOptions && teacherSel.selectedOptions[0]) {
        /* email comes from state sync in entry */
    }
    return {
        titel: val('saTitel'),
        thema: val('saThema'),
        fachCode: val('saFach'),
        klasseCode: val('saKlasse'),
        lehrerCode: val('saLehrer'),
        lehrerEmail,
        datum: val('saDatum'),
        beginnUhrzeit: normalizeBeginnUhrzeit(val('saBeginn')),
        dauerMinuten: Number(val('saDauer')) || 100,
        semester: val('saSemester') || 'WS',
        notiz: val('saNotiz')
    };
}

export function readFiltersFromDom(currentFilters) {
    const g = (id) => {
        const el = document.getElementById(id);
        return el ? String(el.value || '') : '';
    };
    const cur = currentFilters && typeof currentFilters === 'object' ? currentFilters : {};
    return {
        klasse: g('saFilterKlasse'),
        lehrer: g('saFilterLehrer'),
        fach: String(cur.fach || ''),
        status: String(cur.status || '')
    };
}

export function readFensterForm() {
    const g = (id) => {
        const el = document.getElementById(id);
        return el ? String(el.value || '').trim() : '';
    };
    const sjEl = document.getElementById('saSchuljahr');
    const schuljahr = sjEl ? String(sjEl.value || '').trim() : '';
    return {
        titel: g('saTfTitel'),
        startdatum: g('saTfVon'),
        enddatum: g('saTfBis'),
        beschreibung: g('saTfDesc'),
        typ: 'gesperrt',
        schuljahr
    };
}

export function readRulesForm() {
    const g = (id) => {
        const el = document.getElementById(id);
        return el ? String(el.value || '').trim() : '';
    };
    return {
        name: g('saRwName'),
        maxProTag: Number(g('saRwTag')) || 1,
        maxProWoche: Number(g('saRwWoche')) || 2,
        ankuendigungsfristTage: Number(g('saRwFrist')) || 7,
        sperreVorNotenkonferenzTage: Number(g('saRwSperre')) || 7,
        aktiv: true
    };
}
