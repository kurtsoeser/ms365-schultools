/**
 * UI-Rendering Schulaktivitäten-Planer.
 */
import {
    computeAktivitaetStats,
    validateAktivitaet,
    DEFAULT_RULES,
    resolveKostenLimits,
    schoolYearLabel,
    isoInSchoolYear
} from './schulaktivitaeten-planer-logic.js';
import { ueChoices, formatUeZeitraum, applyUeToDraft } from './schulaktivitaeten-planer-zeitraster.js';
import {
    navSectionsForRole,
    countOpenEntscheidungen,
    countUntisPending,
    filterItems,
    labelMaps,
    formatDeDate,
    statusLabel,
    typLabel,
    roleLabel,
    emptyForm,
    canEditItem,
    canDecide,
    canKvStellungnahme,
    canMarkUntis,
    scopeFromState,
    sortByStartdatum,
    typChoices,
    teilnahmeModusChoices,
    teilnahmeKurz,
    isAktStaffRole,
    statusFilterChoices,
    hasAktKvCapability,
    resolveUserRelations,
    formatRelationSummary,
    kvKlasseCodesForState,
    effectiveAccountEmail,
    effectiveTeacherCode,
    teilnahmeModusLabel,
    studentCountForKlasse,
    scheduleStatusOf,
    scheduleStatusLabel
} from './schulaktivitaeten-planer-state.js';
import { parseCodeList, joinCodeList, klasseCodesOf } from './schulaktivitaeten-planer-logic.js';
import {
    AKT_ROLE_ORDER,
    canUseAktRoleSwitcher,
    isAktDemoRoleUiEnabled
} from './schulaktivitaeten-planer-entra-role.js';
import { renderAntragWizard } from './schulaktivitaeten-planer-wizard-ui.js';
import { renderActivityDrawer, renderSubmitSuccess } from './schulaktivitaeten-planer-detail-ui.js';
import {
    renderMs365CalendarShell,
    startOfWeekMonday
} from '../../shared/ms365-calendar-shell.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function renderAktRoleBlock(state) {
    const demoRoleUi = isAktDemoRoleUiEnabled(state.entraGroupsConfigured, state.demoRoleOverride);
    const showRoleSwitcher = canUseAktRoleSwitcher(state.planerRoles || [], demoRoleUi);
    // Produktive Lehrerwelt: nie Admin/KV/Lehrer-Chips – nur explizit ?demoRole=1.
    if (!showRoleSwitcher) {
        if (state.planerAccessDenied) {
            return `<p class="akt-nav__hint">${esc(state.roleHint || 'Kein Zugang.')}</p>`;
        }
        return '';
    }
    const labels = { lehrer: 'Lehrer', kv: 'KV', admin: 'Admin' };
    const switchable = AKT_ROLE_ORDER;
    return (
        '<span class="akt-nav__role-label">Ansicht (Vorschau)</span>' +
        '<div class="akt-nav__role-btns">' +
        switchable
            .map(
                (r) =>
                    `<button type="button" class="akt-chip${state.role === r ? ' is-active' : ''}" data-akt-role="${esc(
                        r
                    )}">${esc(labels[r] || r)}</button>`
            )
            .join('') +
        '</div>'
    );
}

function statusBadge(status) {
    const s = String(status || 'beantragt').toLowerCase();
    let mod = 'warn';
    if (s === 'genehmigt') mod = 'ok';
    else if (s === 'abgelehnt' || s === 'abgesagt') mod = 'bad';
    else if (s === 'nachforderung') mod = 'info';
    else if (s === 'kv_stellungnahme' || s === 'beantragt') mod = 'info';
    else if (s === 'entwurf') mod = 'muted';
    return `<span class="akt-status akt-status--${mod}" data-status="${esc(s)}"><span class="akt-status__dot" aria-hidden="true"></span>${esc(
        statusLabel(s)
    )}</span>`;
}

/**
 * Feldbezogene Hinweise aus Validierungstexten (Calm UI / Progressive Disclosure).
 * @param {{ errors?: string[], warnings?: string[] }} preview
 */
function fieldHintsFromPreview(preview) {
    /** @type {Record<string, { error?: string, warn?: string }>} */
    const map = {};
    const put = (field, kind, msg) => {
        if (!map[field]) map[field] = {};
        if (!map[field][kind]) map[field][kind] = msg;
    };
    (preview.errors || []).forEach((raw) => {
        const m = String(raw || '');
        const low = m.toLowerCase();
        if (low.includes('titel')) put('titel', 'error', m);
        else if (low.includes('klasse')) put('klasseCodes', 'error', m);
        else if (low.includes('startdatum') || low.includes('enddatum') || low.includes('datum'))
            put('startdatum', 'error', m);
        else if (low.includes('vorlauf')) put('startdatum', 'error', m);
        else if (low.includes('gruppe')) put('gruppenText', 'error', m);
        else if (low.includes('teilnehmend')) put('teilnehmendeText', 'error', m);
        else if (low.includes('unterrichtseinheit') || low.includes('zeit')) put('startUe', 'error', m);
        else put('_form', 'error', m);
    });
    (preview.warnings || []).forEach((raw) => {
        const m = String(raw || '');
        const low = m.toLowerCase();
        if (low.includes('ort')) put('ort', 'warn', m);
        else if (low.includes('überschneid') || low.includes('ueberschneid')) put('startdatum', 'warn', m);
        else if (low.includes('vorlauf')) put('startdatum', 'warn', m);
        else if (low.includes('begleit') || low.includes('zeitangabe')) put('startUe', 'warn', m);
        else put('_form', 'warn', m);
    });
    return map;
}

function fieldMsg(hints, name) {
    const h = hints && hints[name];
    if (!h) return '';
    if (h.error) {
        return `<p class="akt-field-msg akt-field-msg--error" role="alert">${esc(h.error)}</p>`;
    }
    if (h.warn) {
        return `<p class="akt-field-msg akt-field-msg--warn">${esc(h.warn)}</p>`;
    }
    return '';
}

/**
 * Mehrfachauswahl als Checkbox-Liste (anhaken, suchbar, Tab-freundlich).
 * @param {string} name
 * @param {{code?:string,value?:string,name?:string,label?:string}[]} items
 * @param {string} selectedCsv
 * @param {{ placeholder?: string, emptyHint?: string, ariaLabel?: string }} [opts]
 */
function renderMultiPick(name, items, selectedCsv, opts) {
    const o = opts || {};
    const selected = parseCodeList(selectedCsv);
    const selectedSet = new Set(selected.map(String));
    const rows = (items || [])
        .map((it) => {
            const v = String(it.code || it.value || '').trim();
            if (!v) return '';
            const label = it.name || it.label || v;
            const checked = selectedSet.has(v);
            return `<label class="akt-mcheck__item">
        <input type="checkbox" name="${esc(name)}" value="${esc(v)}"${checked ? ' checked' : ''}>
        <span class="akt-mcheck__label">${esc(label)}</span>
      </label>`;
        })
        .filter(Boolean)
        .join('');
    const n = selected.length;
    return `<div class="akt-mcheck" data-akt-mcheck="${esc(name)}">
      <input type="search" class="akt-control akt-mcheck__filter" data-akt-mcheck-filter="${esc(
          name
      )}" placeholder="${esc(o.placeholder || 'Suchen …')}" autocomplete="off" aria-label="Liste filtern">
      <div class="akt-mcheck__list" role="group" aria-label="${esc(o.ariaLabel || 'Mehrfachauswahl')}">
        ${
            rows ||
            `<p class="akt-mcheck__empty">${esc(o.emptyHint || 'Keine Einträge verfügbar')}</p>`
        }
      </div>
      <p class="akt-mcheck__count" data-akt-mcheck-count="${esc(name)}">${
          n ? n + ' ausgewählt' : 'Keine Auswahl'
      }</p>
    </div>`;
}

function formSection(title, lead, bodyHtml) {
    return `<fieldset class="akt-form-section">
      <legend>${esc(title)}</legend>
      ${lead ? `<p class="akt-form-section__lead">${esc(lead)}</p>` : ''}
      <div class="akt-form-grid">${bodyHtml}</div>
    </fieldset>`;
}

/**
 * Kompakte Mehrfachauswahl in der Filterleiste.
 * @param {string} key
 * @param {{code?:string,value?:string,name?:string,label?:string}[]} items
 * @param {string} selectedCsv
 * @param {boolean} [open]
 * @param {{ allLabel?: string, countLabel?: string, ariaLabel?: string, searchPlaceholder?: string, emptyHint?: string, showSearch?: boolean, query?: string }} [opts]
 */
function renderFilterMulti(key, items, selectedCsv, open, opts) {
    const o = opts || {};
    const selected = parseCodeList(selectedCsv);
    const selectedSet = new Set(selected.map(String));
    const n = selected.length;
    const allLabel = o.allLabel || 'Alle';
    const countLabel = o.countLabel || 'ausgewählt';
    const query = String(o.query || '')
        .trim()
        .toLowerCase();
    let summary = allLabel;
    if (n === 1) {
        const hit = (items || []).find((it) => String(it.code || it.value || '').trim() === selected[0]);
        summary = (hit && (hit.name || hit.label)) || selected[0];
    } else if (n > 1) {
        summary = n + ' ' + countLabel;
    }
    const inputName = 'aktFilter' + key.charAt(0).toUpperCase() + key.slice(1);
    const rows = (items || [])
        .map((it) => {
            const v = String(it.code || it.value || '').trim();
            if (!v) return '';
            const label = it.name || it.label || v;
            const checked = selectedSet.has(v);
            const hay = String(label)
                .trim()
                .toLowerCase();
            const hidden = !!(query && hay.indexOf(query) === -1);
            return `<label class="akt-filter-multi__item${hidden ? ' is-filtered-out' : ''}">
        <input type="checkbox" name="${esc(inputName)}" value="${esc(v)}"${checked ? ' checked' : ''}>
        <span>${esc(label)}</span>
      </label>`;
        })
        .filter(Boolean)
        .join('');
    const showSearch = o.showSearch !== false;
    return `<details class="akt-filter-multi${n ? ' is-filtered' : ''}" data-akt-filter-multi="${esc(key)}"${
        open ? ' open' : ''
    }>
      <summary class="akt-filter-multi__summary" aria-label="${esc(o.ariaLabel || allLabel)}">${esc(
          summary
      )}</summary>
      <div class="akt-filter-multi__panel">
        ${
            showSearch
                ? `<input type="search" class="akt-control akt-filter-multi__search" data-akt-filter-multi-q="${esc(
                      key
                  )}" value="${esc(o.query || '')}" placeholder="${esc(
                      o.searchPlaceholder || 'Suchen …'
                  )}" autocomplete="off" aria-label="Liste filtern">`
                : ''
        }
        <div class="akt-filter-multi__list" role="group" aria-label="${esc(o.ariaLabel || allLabel)}">
          ${rows || `<p class="akt-filter-multi__empty">${esc(o.emptyHint || 'Keine Einträge')}</p>`}
        </div>
        ${
            n
                ? `<button type="button" class="akt-text-link akt-filter-multi__clear" data-akt-filter-multi-clear="${esc(
                      key
                  )}">Auswahl leeren</button>`
                : ''
        }
      </div>
    </details>`;
}

/** @param {object} filters @param {string} preset */
function statusPresetActive(filters, preset) {
    const list = parseCodeList(filters && filters.status).map((s) => String(s).toLowerCase());
    const hasExtra = !!(filters && filters.verarbeitung);
    if (!preset) return list.length === 0 && !hasExtra;
    if (hasExtra) return false;
    return list.length === 1 && list[0] === String(preset).toLowerCase();
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

/**
 * Seiten-Header (HTML) an App-Rolle anpassen: Lehrer ohne Setup-Ballast.
 * @param {object} state
 */
export function syncAktPageChrome(state) {
    const page = document.querySelector('.akt-page');
    if (page) {
        page.classList.toggle('akt-page--lehrer', state.role === 'lehrer');
        page.classList.toggle('akt-page--staff', isAktStaffRole(state.role));
    }
    try {
        document.documentElement.setAttribute('data-akt-app-role', String(state.role || 'lehrer'));
    } catch {
        /* ignore */
    }
}

function renderTeacherTop(state, opts) {
    const o = opts || {};
    const showCta = o.showPrimaryCta !== false && state.view !== 'antrag';
    return `
    <header class="akt-top akt-top--lehrer">
      <div class="akt-top__intro">
        <h2 class="akt-top__title">${esc(o.title || 'Meine Schulaktivitäten')}</h2>
        ${o.lead ? `<p class="akt-top__lead">${esc(o.lead)}</p>` : ''}
      </div>
      ${
          showCta
              ? `<button type="button" class="btn akt-btn-primary akt-top__cta" data-akt-view-jump="antrag"><i class="bi bi-plus-lg" aria-hidden="true"></i>Neue Aktivität</button>`
              : ''
      }
    </header>`;
}

function renderStaffTop(state) {
    const openN = countOpenEntscheidungen(state.items);
    const titles = {
        dashboard: 'Übersicht',
        liste: 'Alle Aktivitäten',
        kalender: 'Kalender',
        freigabe: 'Entscheidungen',
        berichte: 'Berichte',
        limits: 'Kostenschwellen',
        regeln: 'Regelwerk',
        system: 'Systemadministration',
        antrag: 'Neue Aktivität'
    };
    const title = titles[state.view] || 'Leitung & Betrieb';
    const lead =
        state.view === 'dashboard'
            ? 'Entscheidungen, Nacharbeit und die nächsten Termine im Blick.'
            : state.view === 'freigabe'
              ? 'Anträge prüfen, Rückfragen stellen und Entscheidungen nachvollziehen.'
              : state.view === 'system'
                ? 'Verbinden, konfigurieren und diagnostizieren – getrennt vom Schulbetrieb.'
                : '';
    const workViews = state.view === 'dashboard' || state.view === 'freigabe' || state.view === 'liste';
    const showOpenBadge = workViews && openN > 0;
    const showLeitung = workViews || state.view === 'kalender';
    return `
    <header class="akt-top akt-top--staff-calm">
      <div class="akt-top__copy">
        <h1>${esc(title)}</h1>
        ${lead ? `<p class="akt-top__lead">${esc(lead)}</p>` : ''}
      </div>
      <div class="akt-top__user">
        ${showLeitung ? '<span class="akt-badge akt-badge--info">Leitung</span>' : ''}
        ${state.localDemoOnly && state.view === 'system' ? '<span class="akt-badge akt-badge--warn">Demo lokal</span>' : ''}
        ${
            showOpenBadge
                ? `<button type="button" class="akt-chip is-active" data-akt-view-jump="freigabe">${openN} offen</button>`
                : ''
        }
        ${
            state.accountEmail
                ? `<small>${esc(state.accountName || state.accountEmail)}</small>`
                : workViews
                  ? '<small class="muted">Nicht angemeldet</small>'
                  : ''
        }
      </div>
    </header>`;
}

export function renderApp(state, root) {
    if (!root) return;
    syncAktPageChrome(state);
    const kvKlassenNav = kvKlasseCodesForState(state);
    const openEntscheidungen = countOpenEntscheidungen(state.items);
    const sections = navSectionsForRole(state.role, state.planerRoles, { kvKlassen: kvKlassenNav }).map(
        (sec) => ({
            ...sec,
            items: (sec.items || []).map((v) =>
                v.id === 'antrag' ? { ...v, label: 'Neue Aktivität' } : v
            )
        })
    );
    const staff = isAktStaffRole(state.role);
    const teacherLike = !staff;
    const roleBlock = renderAktRoleBlock(state);
    const personaBanner = state.demoPersonaEmail
        ? `<div class="akt-alert akt-alert--info" role="status">Demo-Ansicht als <strong>${esc(
              state.demoPersonaName || 'Lehrkraft'
          )}</strong> (${esc(state.demoPersonaEmail)}). Angemeldet bleibt ${esc(
              state.accountName || state.accountEmail || 'IT'
          )}.</div>`
        : '';
    const teacherTopOpts =
        state.view === 'dashboard'
            ? { title: 'Meine Schulaktivitäten', showPrimaryCta: true }
            : state.view === 'liste'
              ? { title: 'Meine Aktivitäten', showPrimaryCta: true }
              : state.view === 'kalender'
                ? { title: 'Kalender', showPrimaryCta: true }
                : state.view === 'meine-klasse'
                  ? { title: 'Meine Klasse', showPrimaryCta: false }
                  : state.view === 'antrag'
                    ? {
                          title: state.editingItemId ? 'Aktivität bearbeiten' : 'Neue Aktivität',
                          showPrimaryCta: false
                      }
                    : { title: 'Schulaktivitäten', showPrimaryCta: true };
    const navHtml = sections
        .map((sec) => {
            const label = sec.label
                ? `<span class="akt-nav__section-label">${esc(sec.label)}</span>`
                : '';
            const btns = (sec.items || [])
                .map((v) => {
                    const badge =
                        staff && v.id === 'freigabe' && openEntscheidungen
                            ? `<span class="akt-nav__badge">${openEntscheidungen}</span>`
                            : '';
                    return `
            <button type="button" class="akt-nav__btn${state.view === v.id ? ' is-active' : ''}${
                v.id === 'antrag' ? ' akt-nav__btn--cta' : ''
            }" data-akt-view="${esc(v.id)}">
              <i class="bi ${v.icon}" aria-hidden="true"></i><span class="akt-nav__btn-label">${esc(
                        v.label
                    )}</span>${badge}
            </button>`;
                })
                .join('');
            return `<div class="akt-nav__section">${label}${btns}</div>`;
        })
        .join('');
    root.innerHTML = `
    <div class="akt-shell${teacherLike ? ' akt-shell--lehrer' : ''}${staff ? ' akt-shell--staff' : ''}">
      <aside class="akt-nav" aria-label="Navigation">
        <div class="akt-nav__brand">
          <i class="bi bi-bus-front-fill" aria-hidden="true"></i>
          <div>
            <strong>Schulaktivitäten</strong>
            <span>${staff ? 'Leitung &amp; Betrieb' : 'Planen &amp; begleiten'}</span>
          </div>
        </div>
        <nav class="akt-nav__list">
          ${navHtml}
        </nav>
        ${roleBlock ? `<div class="akt-nav__role">${roleBlock}</div>` : ''}
      </aside>
      <div class="akt-main">
        ${staff ? renderStaffTop(state) : renderTeacherTop(state, teacherTopOpts)}
        ${personaBanner}
        ${
            state.localDemoOnly && !state.demoPersonaEmail && state.view === 'system'
                ? '<div class="akt-alert akt-alert--info" role="status">Lokaler Demo-Modus – Anzeige ohne SharePoint.</div>'
                : ''
        }
        ${state.error ? `<div class="akt-alert akt-alert--bad" role="alert">${esc(state.error)}</div>` : ''}
        ${state.loading ? '<div class="akt-skel" aria-busy="true"><div class="akt-skel__card"></div><div class="akt-skel__card"></div></div>' : ''}
        <div id="aktView" class="akt-view${state.loading ? ' is-loading' : ''}">
          ${
              state.loading
                  ? ''
                  : state.submitSuccess
                    ? renderSubmitSuccess(state)
                    : renderView(state)
          }
        </div>
      </div>
      ${state.detailId ? renderActivityDrawer(state) : ''}
    </div>
  `;
}

/**
 * Statuszeile für abonnierbaren ICS-Feed (nach Veröffentlichung).
 * @param {object} state
 */
function renderFeedBar(state) {
    const feed = state && state.feedMeta;
    if (!feed || !feed.httpsUrl) {
        return `
        <div class="akt-feed" role="note">
          <p class="akt-feed__hint"><i class="bi bi-rss" aria-hidden="true"></i>
            <strong>Kalender-Abo:</strong> „Feed“ veröffentlicht nur <em>genehmigte</em> Aktivitäten als ICS auf der Site
            (Titel, Klasse, Verantwortlich, Begleitung …). Danach in Outlook/Apple als Internetkalender abonnieren.</p>
        </div>`;
    }
    const when = feed.publishedAt
        ? (() => {
              try {
                  return new Date(feed.publishedAt).toLocaleString('de-AT');
              } catch {
                  return feed.publishedAt;
              }
          })()
        : '';
    return `
    <div class="akt-feed akt-feed--ready" role="status">
      <div class="akt-feed__meta">
        <span class="akt-badge akt-badge--ok">Feed aktiv</span>
        <span>${Number(feed.count) || 0} genehmigte Termine</span>
        ${when ? `<span class="muted">Stand ${esc(when)}</span>` : ''}
        ${feed.path ? `<code class="akt-feed__path">${esc(feed.path)}</code>` : ''}
      </div>
      <div class="akt-feed__actions">
        <button type="button" class="btn" id="aktBtnFeedCopyWebcal" title="webcal-Link kopieren (Abo)"><i class="bi bi-clipboard"></i>webcal kopieren</button>
        <button type="button" class="btn" id="aktBtnFeedCopyHttps" title="HTTPS-Link kopieren"><i class="bi bi-link-45deg"></i>HTTPS kopieren</button>
        <a class="btn" id="aktBtnFeedOpen" href="${esc(feed.webcalUrl || feed.httpsUrl)}" rel="noopener"><i class="bi bi-calendar-plus"></i>Abonnieren</a>
        <button type="button" class="btn" id="aktBtnFeedRefresh" title="Feed erneut aus genehmigten Terminen schreiben"><i class="bi bi-arrow-repeat"></i>Aktualisieren</button>
      </div>
    </div>`;
}

function filtersActive(filters) {
    const f = filters || {};
    return !!(f.klasse || f.typ || f.status || f.lehrer || f.beteiligung || f.verarbeitung);
}

function filterBar(state, opts) {
    const sd = state.stammdaten;
    const scope = scopeFromState(state, opts || {});
    const count = filterItems(state.items, state.filters, scope).length;
    const staff = isAktStaffRole(state.role);
    const active = filtersActive(state.filters);
    const calPresets = opts && opts.calendarPresets;
    const todayBtn = opts && opts.showToday;
    return `
    ${
        calPresets
            ? `<div class="akt-seg" role="group" aria-label="Status-Voreinstellung">
        <button type="button" class="akt-seg__btn${statusPresetActive(state.filters, '') ? ' is-active' : ''}" data-akt-status-preset="">Alle Aktivitäten</button>
        <button type="button" class="akt-seg__btn${statusPresetActive(state.filters, 'genehmigt') ? ' is-active' : ''}" data-akt-status-preset="genehmigt">Genehmigt</button>
        <button type="button" class="akt-seg__btn${statusPresetActive(state.filters, 'beantragt') ? ' is-active' : ''}" data-akt-status-preset="beantragt">Offen</button>
      </div>`
            : ''
    }
    <div class="akt-filters${active ? ' akt-filters--active' : ''}" data-akt-filters>
      ${
          todayBtn
              ? '<button type="button" class="btn akt-btn-ghost akt-btn-compact" id="aktCalToday">Heute</button>'
              : ''
      }
      ${renderFilterMulti('klasse', sd.classes, state.filters.klasse, !!(state.filterMultiOpen && state.filterMultiOpen.klasse), {
          allLabel: 'Alle Klassen',
          countLabel: 'Klassen',
          ariaLabel: 'Klassen filtern',
          searchPlaceholder: 'Klasse suchen …',
          emptyHint: 'Keine Klassen verfügbar',
          query: state.filterMultiQuery && state.filterMultiQuery.klasse
      })}
      ${renderFilterMulti('typ', typChoices(), state.filters.typ, !!(state.filterMultiOpen && state.filterMultiOpen.typ), {
          allLabel: 'Alle Arten',
          countLabel: 'Arten',
          ariaLabel: 'Art filtern',
          emptyHint: 'Keine Arten',
          showSearch: false
      })}
      ${renderFilterMulti(
          'status',
          statusFilterChoices(),
          state.filters.status,
          !!(state.filterMultiOpen && state.filterMultiOpen.status),
          {
              allLabel: 'Alle Status',
              countLabel: 'Status',
              ariaLabel: 'Status filtern',
              emptyHint: 'Keine Status',
              showSearch: false
          }
      )}
      ${
          !staff
              ? `<label class="akt-filter-field${state.filters.beteiligung ? ' is-filtered' : ''}"><span class="akt-sr-only">Beteiligung</span>
            <select id="aktFilterBeteiligung" aria-label="Meine Beteiligung">
              <option value="">Alle Beteiligungen</option>
              <option value="applicant"${state.filters.beteiligung === 'applicant' ? ' selected' : ''}>Von mir eingereicht</option>
              <option value="responsible"${state.filters.beteiligung === 'responsible' ? ' selected' : ''}>Von mir verantwortlich</option>
              <option value="companion"${state.filters.beteiligung === 'companion' ? ' selected' : ''}>Ich begleite</option>
            </select>
          </label>`
              : ''
      }
      ${
          staff
              ? `${renderFilterMulti(
                    'lehrer',
                    sd.teachers,
                    state.filters.lehrer,
                    !!(state.filterMultiOpen && state.filterMultiOpen.lehrer),
                    {
                        allLabel: 'Alle Lehrkräfte',
                        countLabel: 'Lehrkräfte',
                        ariaLabel: 'Lehrkraft filtern',
                        searchPlaceholder: 'Lehrkraft suchen …',
                        emptyHint: 'Keine Lehrkräfte verfügbar',
                        query: state.filterMultiQuery && state.filterMultiQuery.lehrer
                    }
                )}
        <label class="akt-filter-field${state.filters.verarbeitung ? ' is-filtered' : ''}"><span class="akt-sr-only">Nacharbeit</span>
          <select id="aktFilterVerarbeitung" aria-label="Nacharbeit">
            <option value="">Nacharbeit: alle</option>
            <option value="offen"${state.filters.verarbeitung === 'offen' ? ' selected' : ''}>Eintragung ausständig</option>
            <option value="erledigt"${state.filters.verarbeitung === 'erledigt' ? ' selected' : ''}>In Untis verarbeitet</option>
          </select>
        </label>`
              : ''
      }
      ${
          active
              ? '<button type="button" class="akt-text-link akt-filter-reset" id="aktFilterReset">Filter zurücksetzen</button>'
              : ''
      }
      <span class="akt-filters__meta">${count} ${count === 1 ? 'Eintrag' : 'Einträge'}</span>
    </div>`;
}

function renderView(state) {
    switch (state.view) {
        case 'liste':
            return renderListe(state);
        case 'kalender':
            return renderKalender(state);
        case 'meine-klasse':
            return renderMeineKlasse(state);
        case 'limits':
            return renderKostenLimits(state);
        case 'berichte':
            return renderBerichte(state);
        case 'antrag':
            return renderAntrag(state);
        case 'freigabe':
            return renderFreigabe(state);
        case 'regeln':
            return renderRegeln(state);
        case 'system':
            return renderSystemAdmin(state);
        default:
            return renderDashboard(state);
    }
}

function todayIsoLocal() {
    const t = new Date();
    return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`;
}

function involvementWho(state) {
    return {
        email: effectiveAccountEmail(state),
        teacherCode: effectiveTeacherCode(state)
    };
}

function involvementChip(it, state) {
    const rels = resolveUserRelations(it, involvementWho(state));
    const summary = formatRelationSummary(rels);
    if (!summary) return '';
    const mod = rels.every((r) => r.type === 'companion')
        ? 'begleitung'
        : rels.some((r) => r.type === 'companion')
          ? 'mixed'
          : 'organizer';
    return `<span class="akt-involve akt-involve--${mod}">${esc(summary)}</span>`;
}

function renderNextActivityCard(it, state) {
    const labels = labelMaps(state.stammdaten);
    const id = it.aktivitaetId || it.itemId;
    const zeit = formatUeZeitraum(it) || [it.startZeit, it.endZeit].filter(Boolean).join('–');
    const meta = [teilnahmeKurz(it, labels.klasse), it.ort || '', zeit].filter(Boolean).join(' · ');
    return `
    <button type="button" class="akt-next-card" data-akt-detail="${esc(id)}">
      <div class="akt-next-card__date">${esc(formatDeDate(it.startdatum))}${
          it.enddatum && it.enddatum !== it.startdatum ? ' – ' + esc(formatDeDate(it.enddatum)) : ''
      }</div>
      <div class="akt-next-card__body">
        <strong>${esc(it.titel || 'Ohne Titel')}</strong>
        <span>${esc(meta || 'Details folgen')} ${involvementChip(it, state)}</span>
      </div>
      ${statusBadge(it.status)}
    </button>`;
}

function renderLehrerDashboard(state) {
    const scope = scopeFromState(state, { onlyMine: true });
    const mine = filterItems(state.items, {}, scope);
    const needsAction = mine.filter((i) => String(i.status).toLowerCase() === 'nachforderung');
    const drafts = mine.filter((i) => String(i.status).toLowerCase() === 'entwurf');
    const kvPending = hasAktKvCapability(state)
        ? state.items.filter((i) => canKvStellungnahme(i, state))
        : [];
    const today = todayIsoLocal();
    const upcoming = mine
        .filter((i) => {
            const st = String(i.status).toLowerCase();
            if (st === 'abgelehnt' || st === 'abgesagt') return false;
            const d = String(i.startdatum || '');
            return d && d >= today;
        })
        .slice()
        .sort((a, b) => String(a.startdatum || '').localeCompare(String(b.startdatum || '')))
        .slice(0, 5);

    let actionBlock = `
    <p class="akt-quiet-status" role="status">
      <i class="bi bi-check-circle" aria-hidden="true"></i>
      <span><strong>Alles erledigt</strong> · Aktuell ist nichts von dir erforderlich.</span>
    </p>`;

    if (kvPending.length) {
        const first = kvPending[0];
        actionBlock = `
    <section class="akt-calm-banner akt-calm-banner--action" role="status">
      <div class="akt-calm-banner__text">
        <p><strong>${kvPending.length === 1 ? '1 Stellungnahme ist erforderlich' : kvPending.length + ' Stellungnahmen sind erforderlich'}</strong></p>
        <p class="muted">${esc(first.titel || 'Ohne Titel')} · ${esc(teilnahmeKurz(first, labelMaps(state.stammdaten).klasse))}</p>
      </div>
      <button type="button" class="btn akt-btn-primary" data-akt-view-jump="meine-klasse">Zu Meine Klasse →</button>
    </section>`;
    } else if (needsAction.length) {
        const first = needsAction[0];
        actionBlock = `
    <section class="akt-calm-banner akt-calm-banner--action" role="status">
      <div class="akt-calm-banner__text">
        <p><strong>${needsAction.length === 1 ? '1 Aktivität braucht deine Ergänzung' : needsAction.length + ' Aktivitäten brauchen deine Ergänzung'}</strong></p>
        <p class="muted">${esc(first.titel || 'Ohne Titel')}${first.rueckfrageText ? ' – ' + esc(first.rueckfrageText) : ''}</p>
      </div>
      <button type="button" class="btn akt-btn-primary" data-akt-edit="${esc(first.itemId)}">Öffnen →</button>
    </section>`;
    } else if (drafts.length) {
        const first = drafts[0];
        actionBlock = `
    <section class="akt-calm-banner akt-calm-banner--action" role="status">
      <div class="akt-calm-banner__text">
        <p><strong>${drafts.length === 1 ? '1 Entwurf wartet auf dich' : drafts.length + ' Entwürfe warten auf dich'}</strong></p>
        <p class="muted">${esc(first.titel || 'Ohne Titel')}</p>
      </div>
      <button type="button" class="btn akt-btn-primary" data-akt-edit="${esc(first.itemId)}">Weiter bearbeiten</button>
    </section>`;
    }

    return `
    ${actionBlock}
    <section class="akt-panel akt-panel--calm">
      <h2>Als Nächstes</h2>
      ${
          upcoming.length
              ? `<div class="akt-next-list">${upcoming.map((it) => renderNextActivityCard(it, state)).join('')}</div>`
              : '<p class="muted">Keine kommenden Aktivitäten geplant.</p>'
      }
      <button type="button" class="akt-text-link" data-akt-view-jump="liste">Alle Aktivitäten anzeigen <i class="bi bi-arrow-right" aria-hidden="true"></i></button>
    </section>`;
}

function renderLeitungDashboard(state) {
    const labels = labelMaps(state.stammdaten);
    const offen = sortByStartdatum(
        (state.items || []).filter((i) => String(i.status || '').toLowerCase() === 'beantragt')
    );
    const untisN = countUntisPending(state.items);
    const today = todayIsoLocal();
    const in14 = new Date();
    in14.setDate(in14.getDate() + 14);
    const until = `${in14.getFullYear()}-${String(in14.getMonth() + 1).padStart(2, '0')}-${String(
        in14.getDate()
    ).padStart(2, '0')}`;
    const upcoming = sortByStartdatum(
        (state.items || []).filter((i) => {
            const st = String(i.status || '').toLowerCase();
            if (st === 'abgelehnt' || st === 'abgesagt' || st === 'entwurf') return false;
            const d = String(i.startdatum || '');
            return d && d >= today && d <= until;
        })
    ).slice(0, 5);

    const first = offen[0];
    const attention = offen.length
        ? `<section class="akt-calm-banner akt-calm-banner--action" role="status">
      <div class="akt-calm-banner__text">
        <p class="akt-calm-banner__eyebrow">Was braucht deine Aufmerksamkeit?</p>
        <p><strong>${offen.length === 1 ? '1 Antrag wartet auf Entscheidung' : offen.length + ' Anträge warten auf Entscheidung'}</strong></p>
        ${
            first
                ? `<p class="muted">${esc(first.titel || 'Ohne Titel')}<br>${esc(
                      formatDeDate(first.startdatum)
                  )}${
                      formatUeZeitraum(first) ? ' · ' + esc(formatUeZeitraum(first)) : ''
                  } · ${esc(teilnahmeKurz(first, labels.klasse))}${
                      first.ort ? ' · ' + esc(first.ort) : ''
                  }</p>`
                : ''
        }
      </div>
      <button type="button" class="btn akt-btn-primary" data-akt-view-jump="freigabe">${
          offen.length === 1 ? 'Antrag prüfen' : 'Entscheidungen öffnen'
      } →</button>
    </section>`
        : `<p class="akt-quiet-status" role="status">
      <i class="bi bi-check-circle" aria-hidden="true"></i>
      <span><strong>Aktuell wartet keine Entscheidung auf dich.</strong></span>
    </p>`;

    const nacharbeit = untisN
        ? `<section class="akt-calm-banner akt-calm-banner--ops" role="status">
      <div class="akt-calm-banner__icon" aria-hidden="true"><i class="bi bi-calendar2-check"></i></div>
      <div class="akt-calm-banner__text">
        <p class="akt-calm-banner__eyebrow">Organisatorische Nacharbeit</p>
        <p><strong>${
            untisN === 1
                ? '1 genehmigte Aktivität wartet noch auf Untis'
                : untisN + ' genehmigte Aktivitäten warten noch auf Untis'
        }</strong></p>
        <p class="muted">Als verarbeitet markieren, sobald die Eintragung erledigt ist.</p>
      </div>
      <button type="button" class="btn akt-btn-secondary" data-akt-view-jump="liste" data-akt-filter-status="genehmigt" data-akt-filter-verarbeitung="offen">Offene Eintragungen →</button>
    </section>`
        : '';

    return `
    <section class="akt-panel akt-panel--calm">
      <h2 class="akt-sr-only">Aufmerksamkeit</h2>
      ${attention}
      ${nacharbeit}
    </section>
    <section class="akt-panel akt-panel--calm">
      <h2>Als Nächstes</h2>
      <p class="muted" style="margin-top:0">Schulweit in den nächsten 14 Tagen.</p>
      ${
          upcoming.length
              ? `<div class="akt-next-list">${upcoming.map((it) => renderNextActivityCard(it, state)).join('')}</div>`
              : '<p class="muted">Keine kommenden Aktivitäten in diesem Zeitraum.</p>'
      }
      <button type="button" class="akt-text-link" data-akt-view-jump="liste">Alle Aktivitäten anzeigen <i class="bi bi-arrow-right" aria-hidden="true"></i></button>
    </section>`;
}

function renderDashboard(state) {
    if (!isAktStaffRole(state.role)) return renderLehrerDashboard(state);
    return renderLeitungDashboard(state);
}

/** Liste/Kalender (Lehrerwelt): persönliche Beteiligung – nicht Klassenkontext. */
function listScopeOpts(state) {
    if (isAktStaffRole(state.role)) return { scopeAll: true };
    return { onlyMine: true };
}

function calendarScopeOpts(state) {
    if (isAktStaffRole(state.role)) return { scopeAll: true };
    return { onlyMine: true };
}

function renderListe(state) {
    const staff = isAktStaffRole(state.role);
    const scope = scopeFromState(state, listScopeOpts(state));
    const items = sortByStartdatum(filterItems(state.items, state.filters, scope));
    const title = staff ? 'Alle Aktivitäten' : 'Meine Aktivitäten';
    const emptyHint = staff
        ? 'Noch keine Aktivitäten in der Schule.'
        : 'Hier erscheinen Aktivitäten, die du selbst einreichst, verantwortest oder begleitest.';
    return `
    <section class="akt-panel akt-panel--list">
      <h2 class="akt-sr-only">${esc(title)}</h2>
      ${filterBar(state, listScopeOpts(state))}
      ${
          items.length
              ? renderTable(items, state)
              : `<div class="akt-empty">
            <p><strong>Noch keine Aktivitäten</strong></p>
            <p class="muted">${esc(emptyHint)}</p>
            ${
                !staff && (state.items || []).length
                    ? `<p class="muted">${(state.items || []).length} Einträge geladen, aber keiner ist dir zugeordnet (Antragsteller/Verantwortlich/Begleitung). Konto: ${esc(String(state.accountEmail || '–'))}. Bitte Seite neu laden oder erneut einreichen.</p>`
                    : !staff
                      ? '<p class="muted">Wenn der Eintrag in SharePoint sichtbar ist, aber hier fehlt: Seite neu laden – die App liest die Liste jetzt auch per SharePoint-REST nach.</p>'
                      : ''
            }
            ${
                staff
                    ? ''
                    : '<button type="button" class="btn akt-btn-primary" data-akt-view-jump="antrag">Neue Aktivität</button>'
            }
          </div>`
      }
    </section>`;
}

function renderMeineKlasse(state) {
    if (!hasAktKvCapability(state)) {
        return '<p class="muted">Keine Klassenvorstand-Zuordnung gefunden.</p>';
    }
    const labels = labelMaps(state.stammdaten);
    const kvCodes = kvKlasseCodesForState(state);
    const primary =
        String(state.meineKlasseCode || '').trim() ||
        kvCodes[0] ||
        '';
    const classList = (state.stammdaten.classes || []).filter((c) =>
        kvCodes.includes(String(c.code || '').trim())
    );
    const className =
        (labels.klasse && labels.klasse[primary]) ||
        primary ||
        (classList[0] && (classList[0].name || classList[0].code)) ||
        'Meine Klasse';
    const scope = scopeFromState(state, {
        scopeAll: true,
        kvKlassen: primary ? [primary] : kvCodes
    });
    const classItems = sortByStartdatum(filterItems(state.items, {}, scope));
    const pending = classItems.filter((i) => canKvStellungnahme(i, state));
    const today = todayIsoLocal();
    const upcoming = classItems
        .filter((i) => {
            const st = String(i.status || '').toLowerCase();
            if (st === 'abgelehnt' || st === 'abgesagt' || st === 'entwurf') return false;
            return String(i.startdatum || '') >= today;
        })
        .slice(0, 5);

    const weekLoad = (() => {
        const start = new Date();
        const day = start.getDay() || 7;
        start.setDate(start.getDate() - day + 1);
        const end = new Date(start);
        end.setDate(start.getDate() + 6);
        const nextStart = new Date(start);
        nextStart.setDate(start.getDate() + 7);
        const nextEnd = new Date(nextStart);
        nextEnd.setDate(nextStart.getDate() + 6);
        const iso = (d) =>
            `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        const inRange = (it, a, b) => {
            const s = String(it.startdatum || '');
            const e = String(it.enddatum || it.startdatum || '');
            return s && e >= a && s <= b;
        };
        const countUe = (it) => {
            const a = Number(it.startUe);
            const b = Number(it.endUe);
            if (Number.isFinite(a) && Number.isFinite(b) && b >= a) return b - a + 1;
            return 0;
        };
        const thisWeek = classItems.filter((it) => inRange(it, iso(start), iso(end)));
        const nextWeek = classItems.filter((it) => inRange(it, iso(nextStart), iso(nextEnd)));
        const ueThis = thisWeek.reduce((n, it) => n + countUe(it), 0);
        const ueNext = nextWeek.reduce((n, it) => n + countUe(it), 0);
        return { thisWeek, nextWeek, ueThis, ueNext };
    })();

    const classPicker =
        classList.length > 1
            ? `<label class="akt-field akt-field--inline">Klasse
          <select id="aktMeineKlassePick">${classList
              .map((c) => {
                  const code = c.code || '';
                  return `<option value="${esc(code)}"${code === primary ? ' selected' : ''}>${esc(
                      c.name || code
                  )}</option>`;
              })
              .join('')}</select>
        </label>`
            : '';

    let action = `
    <p class="akt-quiet-status" role="status">
      <i class="bi bi-check-circle" aria-hidden="true"></i>
      <span><strong>Aktuell ist nichts von dir erforderlich.</strong></span>
    </p>`;
    if (pending.length) {
        const first = pending[0];
        const zeit =
            formatUeZeitraum(first) ||
            [first.startZeit, first.endZeit].filter(Boolean).join('–') ||
            '';
        action = `
    <section class="akt-calm-banner akt-calm-banner--action" role="status">
      <div class="akt-calm-banner__text">
        <p class="akt-calm-banner__eyebrow">Stellungnahme erforderlich</p>
        <p><strong>${esc(first.titel || 'Ohne Titel')}</strong></p>
        <p class="muted">${esc(formatDeDate(first.startdatum))}${zeit ? ' · ' + esc(zeit) : ''} · ${esc(
            teilnahmeKurz(first, labels.klasse)
        )}${first.ort ? ' · ' + esc(first.ort) : ''}</p>
        <p>Die Direktion benötigt deine Stellungnahme.</p>
      </div>
      <button type="button" class="btn akt-btn-primary" data-akt-kv="${esc(first.itemId)}">Stellungnahme abgeben</button>
    </section>`;
    }

    return `
    <section class="akt-panel akt-panel--calm">
      <header class="akt-wizard-head" style="margin-bottom:12px">
        <h2>Meine Klasse</h2>
        <p><strong>${esc(className)}</strong> · Schulaktivitäten, Termine und offene Stellungnahmen im Blick.</p>
        ${classPicker}
      </header>
      ${action}
      <div class="akt-class-load" aria-label="Orientierung">
        <div><strong>Diese Woche</strong><span>${weekLoad.thisWeek.length} ${
            weekLoad.thisWeek.length === 1 ? 'Aktivität' : 'Aktivitäten'
        }${weekLoad.ueThis ? ` · ${weekLoad.ueThis} UE` : ''}</span></div>
        <div><strong>Nächste Woche</strong><span>${weekLoad.nextWeek.length} ${
            weekLoad.nextWeek.length === 1 ? 'Aktivität' : 'Aktivitäten'
        }${weekLoad.ueNext ? ` · ${weekLoad.ueNext} UE` : ''}</span></div>
      </div>
      <h3 class="akt-subhead">Als Nächstes mit der ${esc(className)}</h3>
      ${
          upcoming.length
              ? `<div class="akt-next-list">${upcoming
                    .map((it) => {
                        const zeit =
                            formatUeZeitraum(it) ||
                            [it.startZeit, it.endZeit].filter(Boolean).join('–') ||
                            '';
                        const teil = teilnahmeKurz(it, labels.klasse);
                        return `<button type="button" class="akt-next-card" data-akt-detail="${esc(
                            it.aktivitaetId || it.itemId
                        )}">
            <div class="akt-next-card__date">${esc(formatDeDate(it.startdatum))}</div>
            <div class="akt-next-card__body">
              <strong>${esc(it.titel || 'Ohne Titel')}</strong>
              <span>${esc([zeit, it.ort, teilnahmeModusLabel(it.teilnahmeModus), teil].filter(Boolean).join(' · '))}</span>
            </div>
            ${statusBadge(it.status)}
          </button>`;
                    })
                    .join('')}</div>`
              : `<div class="akt-empty">
            <p><strong>Aktuell keine Schulaktivitäten für deine Klasse.</strong></p>
            <p class="muted">Neue Aktivitäten erscheinen hier automatisch, sobald deine Klasse oder eine Gruppe daraus beteiligt ist.</p>
          </div>`
      }
      ${
          classItems.length
              ? `<button type="button" class="akt-text-link" data-akt-view-jump="liste">Persönliche Aktivitäten anzeigen <i class="bi bi-arrow-right" aria-hidden="true"></i></button>`
              : ''
      }
    </section>`;
}

function statBars(map, labelFn) {
    const entries = Object.entries(map || {}).sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
    if (!entries.length) return '<p class="muted">Keine Daten.</p>';
    const max = Math.max(...entries.map((e) => e[1]), 1);
    return `<ul class="akt-stat-bars">${entries
        .map(([k, n]) => {
            const pct = Math.round((n / max) * 100);
            return `<li class="akt-stat-bars__row">
        <span class="akt-stat-bars__label">${esc(labelFn ? labelFn(k) : k)}</span>
        <span class="akt-stat-bars__track"><span class="akt-stat-bars__fill" style="width:${pct}%"></span></span>
        <span class="akt-stat-bars__n">${esc(n)}</span>
      </li>`;
        })
        .join('')}</ul>`;
}

function countUeForItem(it) {
    const a = Number(it && it.startUe);
    const b = Number(it && it.endUe);
    if (Number.isFinite(a) && Number.isFinite(b) && b >= a) return b - a + 1;
    return 0;
}

function countDaysForItem(it) {
    const s = String((it && it.startdatum) || '').trim();
    const e = String((it && it.enddatum) || s).trim();
    if (!s) return 0;
    const a = new Date(s + 'T12:00:00');
    const b = new Date(e + 'T12:00:00');
    if (Number.isNaN(a.getTime()) || Number.isNaN(b.getTime())) return 1;
    const diff = Math.round((b - a) / 86400000);
    return Math.max(1, diff + 1);
}

function renderBerichte(state) {
    const sj = Number(state.reportYear) || new Date().getFullYear();
    const sjLabel = schoolYearLabel(sj);
    const scope = scopeFromState(state, { scopeAll: state.role === 'admin' });
    const scoped = filterItems(state.items, {}, scope);
    const yearItems = scoped
        .filter((it) => isoInSchoolYear(it.startdatum, sj))
        .slice()
        .sort((a, b) => String(a.startdatum).localeCompare(String(b.startdatum)));
    const stats = computeAktivitaetStats(scoped, { schoolYearStart: sj });
    const labels = labelMaps(state.stammdaten);
    const loadDays = yearItems.reduce((n, it) => n + countDaysForItem(it), 0);
    const loadUe = yearItems.reduce((n, it) => n + countUeForItem(it), 0);
    return `
    <section class="akt-panel akt-panel--calm">
      <div class="akt-cal__head">
        <p class="akt-report-sj">Schuljahr <strong>${esc(sjLabel)}</strong></p>
        <div class="akt-top__row">
          <button type="button" class="btn" id="aktReportPrev" aria-label="Vorheriges Schuljahr"><i class="bi bi-chevron-left"></i></button>
          <button type="button" class="btn" id="aktReportNext" aria-label="Nächstes Schuljahr"><i class="bi bi-chevron-right"></i></button>
        </div>
      </div>
      <p class="akt-report-hero"><strong>${stats.gesamt}</strong> ${
          stats.gesamt === 1 ? 'Aktivität' : 'Aktivitäten'
      }</p>
      <ul class="akt-report-status">
        <li><strong>${stats.byStatus.genehmigt || 0}</strong> genehmigt</li>
        <li><strong>${stats.byStatus.abgelehnt || 0}</strong> abgelehnt</li>
        <li><strong>${
            (stats.byStatus.beantragt || 0) + (stats.byStatus.kv_stellungnahme || 0)
        }</strong> eingereicht</li>
      </ul>
      <p class="muted akt-report-load">${loadDays} betroffene Unterrichtstage · ${loadUe} UE (Summe über Termine)</p>
    </section>
    <section class="akt-panel akt-panel--calm">
      <h3>Nach Art</h3>
      ${statBars(stats.byTyp, (t) => typLabel(t))}
      <h3 style="margin-top:18px">Nach Klasse</h3>
      ${statBars(stats.byKlasse, (c) => labels.klasse[c] || c)}
    </section>
    <section class="akt-panel akt-panel--calm">
      <h3>Jahresliste</h3>
      ${
          yearItems.length
              ? renderTable(yearItems, state)
              : '<p class="muted">Keine Einträge in diesem Schuljahr.</p>'
      }
    </section>`;
}

function renderTable(items, state) {
    const labels = labelMaps(state.stammdaten);
    const staff = isAktStaffRole(state.role);
    return `
    <div class="akt-table-wrap">
      <table class="akt-table akt-table--rows">
        <thead>
          <tr>
            <th>Zeitraum</th>
            ${staff ? '<th>UE / Zeit</th>' : ''}
            <th>Titel</th>
            <th>Art</th>
            <th>Klasse</th>
            <th>Ort</th>
            <th>Status</th>
            ${staff ? '<th>Nacharbeit</th>' : ''}
            <th class="akt-table__chevron"><span class="akt-sr-only">Öffnen</span></th>
          </tr>
        </thead>
        <tbody>
          ${items
              .map((it) => {
                  const id = it.aktivitaetId || it.itemId;
                  const zeit =
                      formatDeDate(it.startdatum) +
                      (it.enddatum && it.enddatum !== it.startdatum ? ' – ' + formatDeDate(it.enddatum) : '');
                  const sched = scheduleStatusLabel(scheduleStatusOf(it));
                  const ort = String(it.ort || '').trim();
                  return `<tr class="akt-table__row" tabindex="0" role="link" data-akt-detail="${esc(id)}" aria-label="${esc(
                      (it.titel || 'Aktivität') + ' öffnen'
                  )}">
              <td>${esc(zeit)}</td>
              ${staff ? `<td>${esc(formatUeZeitraum(it) || 'ohne Zeitangabe')}</td>` : ''}
              <td>
                <strong>${esc(it.titel || 'Ohne Titel')}</strong>
                <div class="akt-table__rel">${involvementChip(it, state)}</div>
              </td>
              <td>${esc(typLabel(it.typ))}</td>
              <td>${esc(teilnahmeKurz(it, labels.klasse))}</td>
              <td>${ort ? esc(ort) : '<span class="akt-missing">Ort nicht erfasst</span>'}</td>
              <td>${statusBadge(it.status)}</td>
              ${
                  staff
                      ? `<td>${
                            sched
                                ? `<span class="akt-status-note">${esc(sched)}</span>`
                                : '<span class="muted">—</span>'
                        }</td>`
                      : ''
              }
              <td class="akt-table__chevron"><i class="bi bi-chevron-right" aria-hidden="true"></i></td>
            </tr>`;
              })
              .join('')}
        </tbody>
      </table>
    </div>`;
}

/**
 * Aktivität → gemeinsames Kalender-Event.
 * @param {object} it
 * @param {Record<string, string>} klasseLabels
 */
function aktItemToCalEvent(it, klasseLabels) {
    const normalized = applyUeToDraft(it || {}).draft;
    const startDay = String(normalized.startdatum || '').trim();
    const endDay = String(normalized.enddatum || startDay).trim() || startDay;
    if (!startDay) return null;
    const startZeit = String(normalized.startZeit || '').trim();
    const endZeit = String(normalized.endZeit || '').trim();
    const timed =
        startDay === endDay &&
        /^\d{1,2}:\d{2}$/.test(startZeit) &&
        /^\d{1,2}:\d{2}$/.test(endZeit);
    const pad = (t) => {
        const m = /^(\d{1,2}):(\d{2})$/.exec(t);
        return m ? String(m[1]).padStart(2, '0') + ':' + m[2] + ':00' : '12:00:00';
    };
    const id = String(it.aktivitaetId || it.itemId || '');
    const kurz = teilnahmeKurz(normalized, klasseLabels);
    const st = String(normalized.status || 'beantragt').toLowerCase();
    return {
        id,
        title: String(normalized.titel || 'Aktivität'),
        start: timed ? startDay + 'T' + pad(startZeit) : startDay + 'T00:00:00',
        end: timed ? endDay + 'T' + pad(endZeit) : endDay + 'T23:59:59',
        allDay: !timed,
        tone: st,
        subtitle: [statusLabel(st), kurz].filter(Boolean).join(' · '),
        openAttr: 'data-akt-detail="' + esc(id) + '"'
    };
}

function renderKalender(state) {
    const mode = String(state.calMode || 'month');
    const y = Number(state.calYear) || new Date().getFullYear();
    const m = Number(state.calMonth) || new Date().getMonth() + 1;
    let weekAnchor = new Date(state.calWeekAnchor || Date.now());
    if (Number.isNaN(weekAnchor.getTime())) weekAnchor = startOfWeekMonday(new Date());
    else weekAnchor = startOfWeekMonday(weekAnchor);

    const scope = scopeFromState(state, calendarScopeOpts(state));
    const items = filterItems(state.items, state.filters, scope).filter((i) => {
        const st = String(i.status).toLowerCase();
        return st !== 'abgelehnt' && st !== 'abgesagt' && st !== 'entwurf';
    });
    const labels = labelMaps(state.stammdaten);
    const events = items
        .map((it) => aktItemToCalEvent(it, labels.klasse))
        .filter(Boolean);

    const shell = renderMs365CalendarShell({
        mode,
        events,
        weekAnchor,
        year: y,
        month0: m - 1,
        prefix: 'akt-cal',
        hint: 'Aktivität antippen für Details · Agenda, Woche und Monat wie in Kalender & Termine.',
        actionsHtml: ''
    });

    return `
    <section class="akt-panel akt-cal-layout">
      <h2 class="akt-sr-only">Kalender</h2>
      ${filterBar(state, { ...calendarScopeOpts(state), calendarPresets: true, showToday: true })}
      ${shell}
    </section>`;
}

function renderAntrag(state) {
    return renderAntragWizard(state, { renderMultiPick });
}

function teacherDisplayName(state, code, email) {
    const teachers = (state.stammdaten && state.stammdaten.teachers) || [];
    const c = String(code || '').trim().toUpperCase();
    const em = String(email || '')
        .trim()
        .toLowerCase();
    const t =
        (c && teachers.find((x) => String(x.code || '').toUpperCase() === c)) ||
        (em &&
            teachers.find(
                (x) =>
                    String(x.email || '')
                        .trim()
                        .toLowerCase() === em
            )) ||
        null;
    if (t && t.name) return String(t.name);
    if (c || em) return c || em;
    return '';
}

function renderEntscheidungRow(it, state, labels) {
    const zeit = formatUeZeitraum(it) || [it.startZeit, it.endZeit].filter(Boolean).join('–');
    const lehrer = teacherDisplayName(state, it.lehrerCode, it.lehrerEmail);
    const verantwortlich = lehrer
        ? `Verantwortlich: ${lehrer}`
        : 'Verantwortliche Lehrkraft nicht erfasst';
    return `
    <button type="button" class="akt-decision-row akt-decision-row--task" data-akt-detail="${esc(it.aktivitaetId || it.itemId)}">
      <div class="akt-decision-row__main">
        <strong>${esc(it.titel || 'Ohne Titel')}</strong>
        <span>${esc(formatDeDate(it.startdatum))}${zeit ? ' · ' + esc(zeit) : ''} · ${esc(
            teilnahmeKurz(it, labels.klasse)
        )}${it.ort ? ' · ' + esc(it.ort) : ''}</span>
        <span class="${lehrer ? 'muted' : 'akt-missing'}">${esc(verantwortlich)}</span>
        ${it.kvStellungnahme ? `<span class="akt-decision-row__kv">KV: ${esc(it.kvStellungnahme)}</span>` : ''}
      </div>
      <div class="akt-decision-row__side">
        ${statusBadge(it.status)}
        <span class="btn akt-btn-primary akt-btn-compact">Prüfen</span>
      </div>
    </button>`;
}

function renderFreigabe(state) {
    if (!canDecide(state) && !hasAktKvCapability(state)) {
        return '<p class="muted">Nur für Klassenvorstand oder Direktion.</p>';
    }
    const labels = labelMaps(state.stammdaten);
    const kvPending = sortByStartdatum(
        state.items.filter(
            (i) =>
                canKvStellungnahme(i, state) ||
                (canDecide(state) && String(i.status).toLowerCase() === 'kv_stellungnahme')
        )
    );
    const open = canDecide(state)
        ? sortByStartdatum(
              filterItems(state.items, { status: 'beantragt' }, { role: 'admin', onlyOpen: true })
          )
        : [];
    return `
    ${
        hasAktKvCapability(state) || kvPending.length
            ? `<section class="akt-panel akt-panel--calm">
      <div class="akt-section-head">
        <h3>KV-Stellungnahme${kvPending.length ? ` · ${kvPending.length}` : ''}</h3>
        <p class="muted">Zuerst Klassenvorstand, danach Direktion.</p>
      </div>
      ${
          kvPending.length
              ? `<div class="akt-decision-list">${kvPending
                    .map((it) => renderEntscheidungRow(it, state, labels))
                    .join('')}</div>`
              : '<p class="muted">Keine Anträge warten auf KV-Stellungnahme.</p>'
      }
    </section>`
            : ''
    }
    ${
        canDecide(state)
            ? `<section class="akt-panel akt-panel--calm" id="aktEntscheidungenOpen">
      <div class="akt-section-head">
        <h3>Offene Anträge${open.length ? ` · ${open.length}` : ''}</h3>
        <p class="muted">Entscheidung im Activity Drawer: genehmigen, Ergänzung fordern oder ablehnen.</p>
      </div>
      ${
          open.length
              ? `<div class="akt-decision-list">${open
                    .map((it) => renderEntscheidungRow(it, state, labels))
                    .join('')}</div>`
              : '<p class="akt-quiet-status"><i class="bi bi-check-circle" aria-hidden="true"></i><span>Keine offenen Anträge.</span></p>'
      }
    </section>`
            : ''
    }`;
}

function renderKostenLimits(state) {
    const limits = resolveKostenLimits(state.rules);
    const canEdit = isAktStaffRole(state.role);
    const dirty = !!state.limitsDirty;
    const flash = state.settingsSavedFlash === 'limits';
    const cards = limits
        .map((l, idx) => {
            if (!canEdit) {
                return `<article class="akt-threshold-card">
          <h3>${esc(l.typ)}</h3>
          <dl>
            <div><dt>Ohne Rücksprache bis</dt><dd>${esc(l.okBis)} €</dd></div>
            <div><dt>Rücksprache erforderlich bis</dt><dd>${esc(l.warnBis)} €</dd></div>
            <div><dt>Darüber</dt><dd>besondere Prüfung</dd></div>
          </dl>
        </article>`;
            }
            return `<article class="akt-threshold-card">
          <label class="akt-threshold-card__title"><span class="akt-sr-only">Art</span>
            <input type="text" name="limitTyp" data-idx="${idx}" value="${esc(l.typ)}">
          </label>
          <div class="akt-threshold-card__grid">
            <label><span>Ohne Rücksprache bis</span>
              <span class="akt-money-input"><input type="number" min="0" step="1" name="limitOk" data-idx="${idx}" value="${esc(l.okBis)}"><span class="akt-money-input__suffix">€</span></span>
            </label>
            <label><span>Rücksprache erforderlich bis</span>
              <span class="akt-money-input"><input type="number" min="0" step="1" name="limitWarn" data-idx="${idx}" value="${esc(l.warnBis)}"><span class="akt-money-input__suffix">€</span></span>
            </label>
          </div>
          <p class="muted akt-threshold-hint">Darüber: besondere Prüfung</p>
        </article>`;
        })
        .join('');
    return `
    <section class="akt-panel akt-panel--calm">
      <header class="akt-wizard-head" style="margin-bottom:12px">
        <h2 class="akt-sr-only">Kostenschwellen</h2>
        <p>Orientierung bei der Planung – Entscheidungsschwellen, keine Zahlung.</p>
      </header>
      <form id="aktKostenLimitsForm" data-akt-dirty-form="limits">
        <div class="akt-threshold-grid">${cards}</div>
        ${
            canEdit
                ? `<div class="akt-actions" style="margin-top:14px">
            <button type="submit" class="btn akt-btn-primary" ${dirty ? '' : 'disabled'}><i class="bi bi-save"></i>Änderungen speichern</button>
            <button type="button" class="btn" id="aktKostenLimitsReset"><i class="bi bi-arrow-counterclockwise"></i>Standard wiederherstellen</button>
            ${flash ? '<p class="akt-save-flash" role="status">✓ Änderungen gespeichert</p>' : ''}
          </div>`
                : ''
        }
      </form>
    </section>`;
}

function renderRegeln(state) {
    const r = state.rules || { ...DEFAULT_RULES, itemId: '', title: '' };
    const pfad = String(r.genehmigungPfad || 'direktion');
    const vorlauf = String(r.vorlaufModus || 'sperre');
    const canSave = !!(r.itemId || state.localDemoOnly);
    const dirty = !!state.rulesDirty;
    const flash = state.settingsSavedFlash === 'regeln';
    return `
    <section class="akt-panel akt-panel--calm">
      <form id="aktRulesForm" class="akt-form akt-form--regeln" data-akt-dirty-form="regeln">
        <div class="akt-regeln-block">
          <h3>Antragstellung</h3>
          <label class="akt-regeln-field">
            <span>Mindestvorlauf</span>
            <div class="akt-regeln-inline">
              <input type="number" min="0" name="minVorlaufTage" value="${esc(r.minVorlaufTage)}">
              <span>Tage</span>
            </div>
            <small>Wie viele Tage vor einer Aktivität muss der Antrag spätestens eingereicht werden?</small>
          </label>
          <label class="akt-regeln-field">
            <span>Bei Unterschreitung</span>
            <select name="vorlaufModus">
              <option value="sperre" ${vorlauf === 'sperre' ? 'selected' : ''}>Antrag blockieren</option>
              <option value="warnung" ${vorlauf === 'warnung' ? 'selected' : ''}>Nur warnen</option>
              <option value="ausnahme" ${vorlauf === 'ausnahme' ? 'selected' : ''}>Ausnahme mit Begründung</option>
            </select>
          </label>
        </div>
        <div class="akt-regeln-block">
          <h3>Überschneidungen</h3>
          <label class="akt-regeln-field">
            <span>Maximal gleichzeitige Aktivitäten pro Klasse</span>
            <input type="number" min="1" name="maxGleichzeitigProKlasse" value="${esc(r.maxGleichzeitigProKlasse)}">
          </label>
        </div>
        <div class="akt-regeln-block">
          <h3>Genehmigungsweg</h3>
          <label class="akt-regeln-field">
            <span>Pfad</span>
            <select name="genehmigungPfad">
              <option value="direktion" ${pfad === 'direktion' ? 'selected' : ''}>Direktion entscheidet</option>
              <option value="kv_dann_direktion" ${pfad === 'kv_dann_direktion' ? 'selected' : ''}>KV-Stellungnahme → Direktion</option>
            </select>
          </label>
        </div>
        <div class="akt-actions" style="margin-top:14px">
          <button type="submit" class="btn akt-btn-primary" ${canSave && dirty ? '' : 'disabled'}><i class="bi bi-save"></i>Regelwerk speichern</button>
          ${flash ? '<p class="akt-save-flash" role="status">✓ Änderungen gespeichert</p>' : ''}
        </div>
        ${
            canSave
                ? ''
                : '<p class="muted" style="margin-top:10px">Kein Regelwerk geladen – unter Systemadministration verbinden und Daten laden.</p>'
        }
      </form>
      <p class="muted" style="margin-top:14px;font-size:0.88em;">Kostenschwellen pflegst du unter <button type="button" class="akt-text-link" data-akt-view-jump="limits">Kostenschwellen</button>.</p>
    </section>`;
}

function renderSystemAdmin(state) {
    const connected = !!(state.ctx && state.ctx.siteId) || (!!state.siteUrl && !state.error);
    const statusLabel = state.loading
        ? 'Verbindung wird geprüft…'
        : state.error
          ? 'Fehler'
          : connected
            ? 'Verbunden'
            : state.siteUrl
              ? 'URL gesetzt – noch nicht geladen'
              : 'Nicht verbunden';
    const statusMod = state.error ? 'bad' : connected ? 'ok' : 'warn';
    const live = state.stammdatenLive || {};
    const aktLive = state.aktivitaetenLive || {};
    const localClasses = ((state.stammdaten && state.stammdaten.classes) || []).length;
    const localTeachers = ((state.stammdaten && state.stammdaten.teachers) || []).length;
    const liveOk =
        (live.source === 'sharepoint' || live.source === 'sharepoint-spo') && live.classCount > 0;
    const liveWarn =
        !liveOk &&
        (live.error || live.classCount === 0 || live.source === 'local-fallback' || live.source === 'error');
    const liveLine = liveOk
        ? `SharePoint${live.source === 'sharepoint-spo' ? ' (REST)' : ''}: ${live.classCount} Klassen, ${live.teacherCount} Lehrkräfte`
        : live.source === 'local-fallback'
          ? `SharePoint-Klassen leer/fehlend – Fallback lokal (${localClasses} Klassen)`
          : live.error
            ? `Live-Laden: ${live.error}`
            : `Im Planer: ${localClasses} Klassen, ${localTeachers} Lehrkräfte`;
    const aktCount =
        typeof aktLive.count === 'number' ? aktLive.count : ((state.items || []).length || 0);
    const aktSource = String(aktLive.source || '');
    const aktOk = aktSource === 'graph' || aktSource === 'spo-rest';
    const aktWarn =
        !aktOk &&
        (aktSource === 'graph-error' ||
            aktSource === 'error' ||
            aktSource === 'demo' ||
            !!aktLive.error ||
            state.localDemoOnly);
    const aktLine = aktOk
        ? `Aktivitäten: ${aktCount} Einträge via ${
              aktSource === 'spo-rest' ? 'SharePoint REST' : 'Microsoft Graph'
          }`
        : aktSource === 'demo' || state.localDemoOnly
          ? `Aktivitäten: ${aktCount} Einträge (Demo lokal)`
          : aktSource === 'graph-error'
            ? `Aktivitäten: Graph eingeschränkt${aktCount ? ` – ${aktCount} lokal sichtbar` : ''}`
            : aktLive.error
              ? `Aktivitäten: ${aktLive.error}`
              : aktCount
                ? `Aktivitäten: ${aktCount} Einträge geladen`
                : 'Aktivitäten: noch nicht geladen';
    const noSite = !String(state.siteUrl || '').trim();
    return `
    <section class="akt-panel akt-panel--calm akt-sys-card">
      <div class="akt-section-head">
        <h3>Datenquelle</h3>
        <p class="muted">SharePoint-Site der Schule – verbinden und Listen verwalten.</p>
      </div>
      ${
          noSite
              ? `<div class="akt-inline-warn" role="status">
        <strong>Keine SharePoint-URL in diesem Browser.</strong>
        <p class="muted" style="margin:6px 0 0">Dass die Listen in SharePoint sichtbar sind, reicht nicht – der Planer braucht die Site-Adresse.
        IT: unter Admin → Schule die <em>Tool-Site-URL</em> auf die MS365-Schultools-Site setzen (dann hat jeder Lehrer-Browser denselben Zeiger),
        oder hier die URL eintragen und „Verbindung prüfen“. Danach IT-Sicherung ablegen.</p>
      </div>`
              : ''
      }
      <p class="akt-sys-status"><span class="akt-status akt-status--${statusMod}"><span class="akt-status__dot" aria-hidden="true"></span>${esc(
          statusLabel
      )}</span></p>
      <p class="akt-sys-status" style="margin-top:8px"><span class="akt-status akt-status--${
          liveOk ? 'ok' : liveWarn ? 'warn' : 'warn'
      }"><span class="akt-status__dot" aria-hidden="true"></span>${esc(liveLine)}</span></p>
      <p class="akt-sys-status" style="margin-top:8px"><span class="akt-status akt-status--${
          aktOk ? 'ok' : aktWarn ? 'warn' : 'warn'
      }"><span class="akt-status__dot" aria-hidden="true"></span>${esc(aktLine)}</span></p>
      ${
          liveWarn
              ? `<div class="akt-inline-warn" style="margin-top:10px" role="status">
        <strong>Klassenauswahl braucht befüllte SharePoint-Listen.</strong>
        <p class="muted" style="margin:6px 0 0">„Nur Berechtigungen“ reicht nicht – unter Listen verwalten „Nur Stammdaten befüllen“ oder „Paket anlegen“ ausführen.</p>
      </div>`
              : ''
      }
      <label for="aktSiteUrl">SharePoint-URL</label>
      <div class="akt-top__row" style="margin-top:6px;flex-wrap:wrap;gap:8px">
        <input type="url" id="aktSiteUrl" value="${esc(state.siteUrl)}" placeholder="https://…sharepoint.com/sites/…" spellcheck="false" style="flex:1;min-width:220px">
        <button type="button" class="btn akt-btn-primary" id="aktBtnLoad"><i class="bi bi-plug"></i>Verbindung prüfen</button>
      </div>
      <div class="akt-actions" style="margin-top:12px">
        <a class="btn" href="sharepoint-liste-schulaktivitaeten.html" style="text-decoration:none"><i class="bi bi-list-ul"></i>Listen verwalten</a>
        <a class="btn" href="schulaktivitaeten-setup.html" style="text-decoration:none"><i class="bi bi-lightning-charge"></i>Flow-Setup</a>
        <a class="btn" href="../tenant.html" style="text-decoration:none"><i class="bi bi-gear"></i>Stammdaten</a>
      </div>
    </section>
    <section class="akt-panel akt-panel--calm akt-sys-card">
      <div class="akt-section-head">
        <h3>Kalenderintegration</h3>
        <p class="muted">Genehmigte Termine als ICS oder abonnierbarer Feed.</p>
      </div>
      ${renderFeedBar(state)}
      <div class="akt-actions" style="margin-top:10px">
        <button type="button" class="btn" id="aktBtnFeed"><i class="bi bi-rss"></i>Feed verwalten</button>
        <button type="button" class="btn" id="aktBtnIcs"><i class="bi bi-download"></i>ICS herunterladen</button>
      </div>
    </section>
    <section class="akt-panel akt-panel--calm akt-sys-card">
      <div class="akt-section-head">
        <h3>Datenimport</h3>
        <p class="muted">JSON-Backup importieren (Absicherung, kein Ersatz für Live-Listen).</p>
      </div>
      <div class="akt-actions">
        <label class="btn" for="aktImportJson"><i class="bi bi-upload"></i>JSON importieren</label>
        <input type="file" id="aktImportJson" accept=".json,application/json" hidden>
      </div>
    </section>
    <section class="akt-panel akt-panel--calm akt-sys-card">
      <div class="akt-section-head">
        <h3>Demo &amp; Diagnose</h3>
        <p class="muted">Vorführung und Fehlersuche – nicht für den Schulalltag.</p>
      </div>
      <div class="akt-actions">
        <button type="button" class="btn" id="aktBtnDemo"><i class="bi bi-database"></i>Demo-Daten laden</button>
        <a class="btn" href="../hilfe.html#tool-schulaktivitaeten-planer" style="text-decoration:none"><i class="bi bi-life-preserver"></i>Hilfe / Diagnose</a>
      </div>
    </section>
    <section class="akt-panel akt-panel--calm akt-sys-card akt-sys-card--danger">
      <div class="akt-section-head">
        <h3>Danger Zone</h3>
        <p class="muted">Unwiderrufliche bzw. disruptive Aktionen.</p>
      </div>
      <button type="button" class="btn akt-btn-danger" id="aktBtnDemoReset"><i class="bi bi-trash"></i>Demo-Daten zurücksetzen</button>
    </section>`;
}

export function readFormFromDom() {
    const form = document.getElementById('aktForm');
    if (!form) return emptyForm();
    const fd = new FormData(form);
    const get = (k) => String(fd.get(k) || '').trim();
    const getAll = (k) =>
        fd
            .getAll(k)
            .map((v) => String(v || '').trim())
            .filter(Boolean);
    const klasseCodes = joinCodeList(getAll('klasseCodes'));
    const begleitungCodes = joinCodeList(getAll('begleitungCodes'));
    const teilnahmeModus = get('teilnahmeModus') || 'klasse';
    const raw = {
        titel: get('titel'),
        typ: get('typ') || 'Exkursion',
        teilnahmeModus,
        klasseCodes,
        klasseCode: klasseCodesOf({ klasseCodes })[0] || '',
        gruppenText: get('gruppenText'),
        teilnehmendeText: get('teilnehmendeText'),
        nichtTeilnehmende: get('nichtTeilnehmende'),
        lehrerCode: get('lehrerCode'),
        lehrerEmail: get('lehrerEmail'),
        beantragtVon: get('beantragtVon'),
        ort: get('ort'),
        unterrichtsbezug: get('unterrichtsbezug'),
        externeVortragende: get('externeVortragende'),
        startdatum: get('startdatum'),
        enddatum: get('enddatum') || get('startdatum'),
        startUe: get('startUe'),
        endUe: get('endUe'),
        startZeit: get('startZeit'),
        endZeit: get('endZeit'),
        begleitungCodes,
        begleitung: begleitungCodes,
        ablauf: get('ablauf'),
        verkehrsmittel: get('verkehrsmittel'),
        kostenHinweis: get('kostenHinweis'),
        notiz: get('notiz')
    };
    return applyUeToDraft(raw).draft;
}

function readMultiFilterFromDom(key, inputName) {
    const multi = document.querySelector(`[data-akt-filter-multi="${key}"]`);
    if (!multi) return null;
    const checked = Array.from(
        multi.querySelectorAll(`input[type="checkbox"][name="${inputName}"]:checked`)
    )
        .map((el) => String(el.value || '').trim())
        .filter(Boolean);
    return joinCodeList(checked);
}

export function readFiltersFromDom() {
    const val = (id) => {
        const el = document.getElementById(id);
        return el ? String(el.value || '').trim() : '';
    };
    return {
        klasse: readMultiFilterFromDom('klasse', 'aktFilterKlasse') || '',
        typ: readMultiFilterFromDom('typ', 'aktFilterTyp') || '',
        status: readMultiFilterFromDom('status', 'aktFilterStatus') || '',
        lehrer: readMultiFilterFromDom('lehrer', 'aktFilterLehrer') || '',
        beteiligung: val('aktFilterBeteiligung'),
        verarbeitung: val('aktFilterVerarbeitung')
    };
}

export function readRulesForm() {
    const form = document.getElementById('aktRulesForm');
    if (!form) return null;
    const fd = new FormData(form);
    return {
        minVorlaufTage: Number(fd.get('minVorlaufTage')) || 0,
        maxGleichzeitigProKlasse: Number(fd.get('maxGleichzeitigProKlasse')) || 1,
        vorlaufModus: String(fd.get('vorlaufModus') || 'sperre'),
        genehmigungPfad: String(fd.get('genehmigungPfad') || 'direktion')
    };
}

/** @returns {{ typ: string, okBis: number, warnBis: number }[]|null} */
export function readKostenLimitsForm() {
    const form = document.getElementById('aktKostenLimitsForm');
    if (!form) return null;
    const typs = Array.from(form.querySelectorAll('input[name="limitTyp"]'));
    const oks = Array.from(form.querySelectorAll('input[name="limitOk"]'));
    const warns = Array.from(form.querySelectorAll('input[name="limitWarn"]'));
    const out = [];
    for (let i = 0; i < typs.length; i++) {
        const typ = String((typs[i] && typs[i].value) || '').trim();
        if (!typ) continue;
        out.push({
            typ,
            okBis: Number(oks[i] && oks[i].value) || 0,
            warnBis: Number(warns[i] && warns[i].value) || 0
        });
    }
    return out;
}
