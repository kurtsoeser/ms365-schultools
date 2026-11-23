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

function statusPillMod(status) {
    const s = String(status || 'beantragt').toLowerCase();
    if (s === 'genehmigt') return 'ok';
    if (s === 'abgelehnt' || s === 'abgesagt') return 'bad';
    return 'warn';
}

function statusBadge(status) {
    const s = String(status || 'beantragt').toLowerCase();
    const mod = statusPillMod(s);
    return `<span class="fb-pill fb-pill--${mod}" data-status="${esc(s)}">${esc(statusLabel(s))}</span>`;
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
    const selfTitled =
        state.view === 'dashboard' ||
        state.view === 'liste' ||
        state.view === 'meine-klasse' ||
        state.view === 'freigabe' ||
        state.view === 'antrag';
    if (selfTitled) return '';
    const showCta = o.showPrimaryCta !== false && state.view !== 'antrag';
    return `
    <header class="akt-top akt-top--lehrer">
      <div class="akt-top__intro">
        <h2 class="akt-top__title">${esc(o.title || 'Meine Schulaktivitäten')}</h2>
        ${o.lead ? `<p class="akt-top__lead">${esc(o.lead)}</p>` : ''}
      </div>
      ${
          showCta
              ? `<button type="button" class="fb-btn akt-top__cta" data-akt-view-jump="antrag">Neue Aktivität</button>`
              : ''
      }
    </header>`;
}

function renderStaffTop(state) {
    const openN = countOpenEntscheidungen(state.items);
    const selfTitled =
        state.view === 'dashboard' ||
        state.view === 'liste' ||
        state.view === 'freigabe' ||
        state.view === 'berichte' ||
        state.view === 'antrag';
    if (selfTitled) {
        if (!(openN > 0 && (state.view === 'dashboard' || state.view === 'liste'))) return '';
        return `<div class="akt-top-refresh">
          <button type="button" class="fb-btn fb-btn--ghost" data-akt-view-jump="freigabe">${openN} offen</button>
        </div>`;
    }
    const titles = {
        kalender: 'Kalender',
        limits: 'Kostenschwellen',
        regeln: 'Regelwerk',
        system: 'Systemadministration'
    };
    const title = titles[state.view] || 'Leitung & Betrieb';
    const lead =
        state.view === 'system'
            ? 'Site, Berechtigungen, Kalender und Wartung in Tabs – Danger Zone bewusst getrennt.'
            : state.view === 'kalender'
              ? 'Genehmigte und offene Aktivitäten im Monats- oder Wochenraster.'
              : '';
    return `
    <header class="akt-top akt-top--staff-calm">
      <div class="akt-top__copy">
        <h1>${esc(title)}</h1>
        ${lead ? `<p class="akt-top__lead">${esc(lead)}</p>` : ''}
      </div>
      <div class="akt-top__user">
        ${state.localDemoOnly && state.view === 'system' ? '<span class="fb-pill fb-pill--warn">Demo lokal</span>' : ''}
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
            <strong>Kalender-Abo:</strong> „Feed“ schreibt genehmigte Termine als ICS auf <em>diese</em> Schul-Site (funktioniert pro Schule mit eigener SharePoint-URL).
            Outlook: HTTPS-Link → Aus dem Internet (webcal:// auf SharePoint oft Login-Fehler).
            Überall zuverlässig: Button <em>ICS</em> in der Kalenderansicht und Import in Outlook.</p>
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
    const outlookUrl = String(feed.outlookHttpsUrl || feed.anonymousUrl || feed.httpsUrl || '').trim();
    const badge = feed.anonymousShare
        ? '<span class="akt-badge akt-badge--ok">Öffentliches Abo möglich</span>'
        : '<span class="akt-badge akt-badge--warn">Nur mit Anmeldung</span>';
    const hint = String(feed.subscribeHint || '').trim();
    return `
    <div class="akt-feed akt-feed--ready" role="status">
      <div class="akt-feed__meta">
        ${badge}
        <span>${Number(feed.count) || 0} genehmigte Termine</span>
        ${when ? `<span class="muted">Stand ${esc(when)}</span>` : ''}
        ${feed.path ? `<code class="akt-feed__path">${esc(feed.path)}</code>` : ''}
      </div>
      ${hint ? `<p class="akt-feed__hint muted" style="margin:8px 0 0">${esc(hint)}</p>` : ''}
      <div class="akt-feed__actions">
        <button type="button" class="btn akt-btn-primary" id="aktBtnFeedCopyHttps" title="HTTPS-Link für Outlook kopieren"><i class="bi bi-clipboard"></i>Outlook-Link kopieren</button>
        <button type="button" class="btn" id="aktBtnFeedCopyWebcal" title="webcal für Apple Kalender"><i class="bi bi-apple"></i>webcal (Apple)</button>
        <a class="btn" id="aktBtnFeedOpen" href="${esc(outlookUrl)}" target="_blank" rel="noopener"><i class="bi bi-box-arrow-up-right"></i>Link öffnen</a>
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
      ${
          opts && opts.showIcsDownload
              ? `<button type="button" class="btn akt-btn-ghost akt-btn-compact" id="aktBtnCalIcs" title="Sichtbare Termine als ICS herunterladen (Outlook-Import)"><i class="bi bi-download" aria-hidden="true"></i>ICS</button>`
              : ''
      }
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
    const range =
        it.enddatum && it.enddatum !== it.startdatum
            ? `${formatDeDate(it.startdatum)} – ${formatDeDate(it.enddatum)}`
            : formatDeDate(it.startdatum);
    return `<button type="button" class="fb-rowbtn" data-akt-detail="${esc(id)}">
      <div class="fb-rowbtn__main">
        <div class="fb-rowbtn__title">${esc(it.titel || 'Ohne Titel')}</div>
        <div class="fb-rowbtn__meta">${esc(meta || 'Details folgen')}</div>
      </div>
      <div class="fb-rowbtn__date">${esc(range)}</div>
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

    const attentionN = kvPending.length || needsAction.length || drafts.length;
    let headline =
        attentionN === 0
            ? '<span class="fb-num">0</span> offene Aufgaben.'
            : attentionN === 1
              ? '<span class="fb-num">1</span> Aufgabe wartet auf dich.'
              : `<span class="fb-num">${attentionN}</span> Aufgaben warten auf dich.`;
    let sub = 'Aktuell ist nichts von dir erforderlich.';
    let attentionRows = '';

    if (kvPending.length) {
        headline =
            kvPending.length === 1
                ? '<span class="fb-num">1</span> Stellungnahme erforderlich.'
                : `<span class="fb-num">${kvPending.length}</span> Stellungnahmen erforderlich.`;
        sub = 'Klassenvorstand zuerst – danach Direktion.';
        attentionRows = `<div class="fb-sec-label">KV-STELLUNGNAHME</div>
      <div class="fb-rowlist">${kvPending
          .slice(0, 6)
          .map((it) => renderEntscheidungRow(it, state, labelMaps(state.stammdaten)))
          .join('')}</div>`;
    } else if (needsAction.length) {
        headline =
            needsAction.length === 1
                ? '<span class="fb-num">1</span> Aktivität braucht Ergänzung.'
                : `<span class="fb-num">${needsAction.length}</span> Aktivitäten brauchen Ergänzung.`;
        sub = 'Zeile öffnen und Rückfrage beantworten.';
        attentionRows = `<div class="fb-sec-label">NACHFORDERUNG</div>
      <div class="fb-rowlist">${needsAction.slice(0, 6).map((it) => renderNextActivityCard(it, state)).join('')}</div>`;
    } else if (drafts.length) {
        headline =
            drafts.length === 1
                ? '<span class="fb-num">1</span> Entwurf wartet.'
                : `<span class="fb-num">${drafts.length}</span> Entwürfe warten.`;
        sub = 'Weiterbearbeiten und einreichen.';
        attentionRows = `<div class="fb-sec-label">ENTWÜRFE</div>
      <div class="fb-rowlist">${drafts.slice(0, 6).map((it) => renderNextActivityCard(it, state)).join('')}</div>`;
    }

    return `
    <h1 class="fb-greet">${headline}</h1>
    <p class="fb-sub">${esc(sub)}</p>
    ${attentionRows}
    <div class="fb-sec-label">ALS NÄCHSTES</div>
    ${
        upcoming.length
            ? `<div class="fb-rowlist">${upcoming.map((it) => renderNextActivityCard(it, state)).join('')}</div>`
            : '<p class="muted" style="margin:6px 0 0;">Keine kommenden Aktivitäten geplant.</p>'
    }
    <div class="fb-actions-row">
      <button type="button" class="fb-btn" data-akt-view-jump="antrag">Neue Aktivität</button>
      <button type="button" class="fb-btn fb-btn--ghost" data-akt-view-jump="liste">Alle Aktivitäten</button>
      ${
          kvPending.length
              ? '<button type="button" class="fb-btn fb-btn--ghost" data-akt-view-jump="meine-klasse">Meine Klasse</button>'
              : ''
      }
    </div>`;
}

function renderLeitungDashboard(state) {
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
    const labels = labelMaps(state.stammdaten);

    const headline =
        offen.length === 1
            ? '<span class="fb-num">1</span> Antrag ohne Entscheidung.'
            : `<span class="fb-num">${offen.length}</span> ohne Entscheidung.`;
    const sub = offen.length
        ? 'Zeile anklicken, Details prüfen und im Drawer entscheiden.'
        : untisN
          ? `${untisN} genehmigte Aktivität${untisN === 1 ? '' : 'en'} warten noch auf Untis.`
          : 'Aktuell wartet keine Entscheidung auf dich.';

    return `
    <h1 class="fb-greet">${headline}</h1>
    <p class="fb-sub">${esc(sub)}</p>
    ${
        offen.length
            ? `<div class="fb-sec-label">OFFEN</div>
               <div class="fb-rowlist">${offen
                   .slice(0, 8)
                   .map((it) => renderEntscheidungRow(it, state, labels))
                   .join('')}</div>`
            : ''
    }
    ${
        untisN
            ? `<div class="fb-sec-label">UNTIS-NACHARBEIT · ${untisN}</div>
               <p class="muted" style="margin:6px 0 0;">Als verarbeitet markieren, sobald die Eintragung erledigt ist.</p>`
            : ''
    }
    <div class="fb-sec-label">NÄCHSTE 14 TAGE</div>
    ${
        upcoming.length
            ? `<div class="fb-rowlist">${upcoming.map((it) => renderNextActivityCard(it, state)).join('')}</div>`
            : '<p class="muted" style="margin:6px 0 0;">Keine kommenden Aktivitäten in diesem Zeitraum.</p>'
    }
    <div class="fb-actions-row">
      ${
          offen.length
              ? '<button type="button" class="fb-btn" data-akt-view-jump="freigabe">Entscheidungen öffnen</button>'
              : '<button type="button" class="fb-btn" data-akt-view-jump="antrag">Neue Aktivität</button>'
      }
      <button type="button" class="fb-btn fb-btn--ghost" data-akt-view-jump="liste">Alle Aktivitäten</button>
      ${
          untisN
              ? '<button type="button" class="fb-btn fb-btn--ghost" data-akt-view-jump="liste" data-akt-filter-status="genehmigt" data-akt-filter-verarbeitung="offen">Untis offen</button>'
              : '<button type="button" class="fb-btn fb-btn--ghost" data-akt-view-jump="kalender">Kalender</button>'
      }
    </div>`;
}

function renderDashboard(state) {
    if (!isAktStaffRole(state.role)) return renderLehrerDashboard(state);
    return renderLeitungDashboard(state);
}

/** Liste: persönliche Beteiligung. */
function listScopeOpts(state) {
    if (isAktStaffRole(state.role)) return { scopeAll: true };
    return { onlyMine: true };
}

/**
 * Kalender: alle Lehrkräfte sehen den Schulkalender (auch fremde Klassen).
 * Bearbeiten bleibt über canEditItem nur bei eigenen Einträgen.
 */
function calendarScopeOpts(_state) {
    return { scopeAll: true };
}

function renderListe(state) {
    const staff = isAktStaffRole(state.role);
    const scope = scopeFromState(state, listScopeOpts(state));
    const items = sortByStartdatum(filterItems(state.items, state.filters, scope));
    const headline =
        items.length === 1
            ? '<span class="fb-num">1</span> Aktivität im Überblick.'
            : `<span class="fb-num">${items.length}</span> Aktivitäten im Überblick.`;
    const sub = staff
        ? 'Gefiltert oder durchsucht. Zeile öffnet das Detail.'
        : 'Aktivitäten, die du einreichst, verantwortest oder begleitest.';
    const emptyHint = staff
        ? 'Noch keine Aktivitäten in der Schule.'
        : 'Hier erscheinen Aktivitäten, die du selbst einreichst, verantwortest oder begleitest.';
    return `
    <section class="akt-panel akt-panel--calm akt-panel--list">
      <h1 class="fb-greet">${headline}</h1>
      <p class="fb-sub">${esc(sub)}</p>
      ${filterBar(state, listScopeOpts(state))}
      ${
          items.length
              ? renderTable(items, state)
              : `<p class="akt-quiet-status" role="status">
        <i class="bi bi-check-circle" aria-hidden="true"></i>
        <span>${esc(emptyHint)}</span>
      </p>
      ${
          staff
              ? ''
              : '<div class="fb-actions-row"><button type="button" class="fb-btn" data-akt-view-jump="antrag">Neue Aktivität</button></div>'
      }`
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

/**
 * @param {Record<string, number>} map
 * @param {(k: string) => string} [labelFn]
 * @param {{ filterKey?: string, activeKey?: string, toneFn?: (k: string) => string }} [opts]
 */
function statBars(map, labelFn, opts) {
    const o = opts || {};
    const entries = Object.entries(map || {}).sort(
        (a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0]), 'de')
    );
    if (!entries.length) return '<p class="muted">Keine Daten.</p>';
    const max = Math.max(...entries.map((e) => e[1]), 1);
    const total = entries.reduce((s, e) => s + e[1], 0) || 1;
    const filterKey = o.filterKey || '';
    return `<ul class="akt-stat-bars">${entries
        .map(([k, n]) => {
            const pct = Math.round((n / max) * 100);
            const share = Math.round((n / total) * 100);
            const label = labelFn ? labelFn(k) : k;
            const active = filterKey && String(o.activeKey || '') === String(k);
            const tone = o.toneFn ? o.toneFn(k) : '';
            const fillClass = tone ? ` akt-stat-bars__fill--${tone}` : '';
            const inner = `
        <span class="akt-stat-bars__label" title="${esc(label)}">${esc(label)}</span>
        <span class="akt-stat-bars__track"><span class="akt-stat-bars__fill${fillClass}" style="width:${pct}%"></span></span>
        <span class="akt-stat-bars__n" title="${share} %">${esc(n)}</span>`;
            if (!filterKey) {
                return `<li class="akt-stat-bars__row">${inner}</li>`;
            }
            return `<li class="akt-stat-bars__row">
        <button type="button" class="akt-stat-bars__btn${active ? ' is-active' : ''}" data-akt-report-filter="${esc(
                filterKey
            )}" data-akt-report-value="${esc(k)}" aria-pressed="${active ? 'true' : 'false'}">${inner}</button>
      </li>`;
        })
        .join('')}</ul>`;
}

function reportMonthLabel(ym) {
    const m = /^(\d{4})-(\d{2})$/.exec(String(ym || ''));
    if (!m) return String(ym || '–');
    const months = [
        'Jan',
        'Feb',
        'Mär',
        'Apr',
        'Mai',
        'Jun',
        'Jul',
        'Aug',
        'Sep',
        'Okt',
        'Nov',
        'Dez'
    ];
    const idx = Number(m[2]) - 1;
    return `${months[idx] || m[2]} ${m[1]}`;
}

function statusBarTone(st) {
    const s = String(st || '').toLowerCase();
    if (s === 'genehmigt') return 'ok';
    if (s === 'abgelehnt' || s === 'abgesagt') return 'bad';
    if (s === 'nachforderung') return 'warn';
    if (s === 'beantragt' || s === 'kv_stellungnahme') return 'info';
    return 'muted';
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
    const allYear = scoped
        .filter((it) => isoInSchoolYear(it.startdatum, sj))
        .slice()
        .sort((a, b) => String(a.startdatum).localeCompare(String(b.startdatum)));
    const stats = computeAktivitaetStats(scoped, { schoolYearStart: sj });
    const labels = labelMaps(state.stammdaten);
    const filterTyp = String(state.reportFilterTyp || '').trim();
    const filterStatus = String(state.reportFilterStatus || '').trim().toLowerCase();
    const filterKlasse = String(state.reportFilterKlasse || '').trim().toUpperCase();
    const yearItems = allYear.filter((it) => {
        if (filterTyp && String(it.typ || '') !== filterTyp) return false;
        if (filterStatus && String(it.status || '').toLowerCase() !== filterStatus) return false;
        if (filterKlasse && !klasseCodesOf(it).includes(filterKlasse)) return false;
        return true;
    });
    const loadDays = allYear.reduce((n, it) => n + countDaysForItem(it), 0);
    const loadUe = allYear.reduce((n, it) => n + countUeForItem(it), 0);
    const offen =
        (stats.byStatus.beantragt || 0) + (stats.byStatus.kv_stellungnahme || 0);
    const breakdownRaw = String(state.reportBreakdown || 'typ').toLowerCase();
    const breakdown =
        breakdownRaw === 'klasse' ||
        breakdownRaw === 'status' ||
        breakdownRaw === 'month'
            ? breakdownRaw
            : 'typ';
    const hasListFilter = !!(filterTyp || filterStatus || filterKlasse);
    const filterChips = [];
    if (filterTyp) {
        filterChips.push(
            `<button type="button" class="akt-report-chip is-active" data-akt-report-filter="typ" data-akt-report-value="">${esc(
                typLabel(filterTyp)
            )} <i class="bi bi-x" aria-hidden="true"></i></button>`
        );
    }
    if (filterStatus) {
        filterChips.push(
            `<button type="button" class="akt-report-chip is-active" data-akt-report-filter="status" data-akt-report-value="">${esc(
                statusLabel(filterStatus)
            )} <i class="bi bi-x" aria-hidden="true"></i></button>`
        );
    }
    if (filterKlasse) {
        filterChips.push(
            `<button type="button" class="akt-report-chip is-active" data-akt-report-filter="klasse" data-akt-report-value="">${esc(
                labels.klasse[filterKlasse] || filterKlasse
            )} <i class="bi bi-x" aria-hidden="true"></i></button>`
        );
    }
    let breakdownHtml = '';
    if (breakdown === 'klasse') {
        breakdownHtml = statBars(stats.byKlasse, (c) => labels.klasse[c] || c, {
            filterKey: 'klasse',
            activeKey: filterKlasse
        });
    } else if (breakdown === 'status') {
        breakdownHtml = statBars(stats.byStatus, (s) => statusLabel(s), {
            filterKey: 'status',
            activeKey: filterStatus,
            toneFn: statusBarTone
        });
    } else if (breakdown === 'month') {
        const byMonthSorted = Object.fromEntries(
            Object.entries(stats.byMonth || {}).sort((a, b) => String(a[0]).localeCompare(String(b[0])))
        );
        breakdownHtml = statBars(byMonthSorted, reportMonthLabel);
    } else {
        breakdownHtml = statBars(stats.byTyp, (t) => typLabel(t), {
            filterKey: 'typ',
            activeKey: filterTyp
        });
    }
    const seg = (id, label) =>
        `<button type="button" class="akt-seg__btn${
            breakdown === id ? ' is-active' : ''
        }" data-akt-report-breakdown="${id}">${label}</button>`;

    return `
    <section class="akt-panel akt-panel--calm akt-report">
      <div class="akt-report__head">
        <div>
          <p class="akt-report-sj">Schuljahr <strong>${esc(sjLabel)}</strong></p>
          <p class="akt-report__lead muted">Überblick und Jahresliste – Balken filtern die Liste.</p>
        </div>
        <div class="akt-report__year-nav" role="group" aria-label="Schuljahr wechseln">
          <button type="button" class="btn akt-btn-ghost akt-btn-compact" id="aktReportPrev" aria-label="Vorheriges Schuljahr"><i class="bi bi-chevron-left"></i></button>
          <button type="button" class="btn akt-btn-ghost akt-btn-compact" id="aktReportNext" aria-label="Nächstes Schuljahr"><i class="bi bi-chevron-right"></i></button>
        </div>
      </div>

      <div class="akt-report-kpis" role="list">
        <button type="button" class="akt-report-kpi${
            !hasListFilter ? ' is-active' : ''
        }" data-akt-report-clear role="listitem">
          <span class="akt-report-kpi__n">${stats.gesamt}</span>
          <span class="akt-report-kpi__l">${stats.gesamt === 1 ? 'Aktivität' : 'Aktivitäten'}</span>
        </button>
        <button type="button" class="akt-report-kpi akt-report-kpi--ok${
            filterStatus === 'genehmigt' ? ' is-active' : ''
        }" data-akt-report-filter="status" data-akt-report-value="genehmigt" role="listitem">
          <span class="akt-report-kpi__n">${stats.byStatus.genehmigt || 0}</span>
          <span class="akt-report-kpi__l">Genehmigt</span>
        </button>
        <button type="button" class="akt-report-kpi akt-report-kpi--info${
            filterStatus === 'beantragt' ? ' is-active' : ''
        }" data-akt-report-filter="status" data-akt-report-value="beantragt" role="listitem">
          <span class="akt-report-kpi__n">${offen}</span>
          <span class="akt-report-kpi__l">Eingereicht</span>
        </button>
        <button type="button" class="akt-report-kpi akt-report-kpi--bad${
            filterStatus === 'abgelehnt' ? ' is-active' : ''
        }" data-akt-report-filter="status" data-akt-report-value="abgelehnt" role="listitem">
          <span class="akt-report-kpi__n">${stats.byStatus.abgelehnt || 0}</span>
          <span class="akt-report-kpi__l">Abgelehnt</span>
        </button>
        <div class="akt-report-kpi akt-report-kpi--static" role="listitem">
          <span class="akt-report-kpi__n">${loadDays}</span>
          <span class="akt-report-kpi__l">Unterrichtstage</span>
        </div>
        <div class="akt-report-kpi akt-report-kpi--static" role="listitem">
          <span class="akt-report-kpi__n">${loadUe}</span>
          <span class="akt-report-kpi__l">UE gesamt</span>
        </div>
      </div>

      <div class="akt-report-breakdown">
        <div class="akt-report-breakdown__toolbar">
          <div class="akt-seg akt-seg--report" role="tablist" aria-label="Aufschlüsselung">
            ${seg('typ', 'Art')}
            ${seg('klasse', 'Klasse')}
            ${seg('status', 'Status')}
            ${seg('month', 'Monat')}
          </div>
          ${
              hasListFilter
                  ? `<button type="button" class="akt-text-link" data-akt-report-clear>Filter zurücksetzen</button>`
                  : '<span class="muted akt-report-breakdown__hint">Tippe auf einen Balken zum Filtern</span>'
          }
        </div>
        ${
            filterChips.length
                ? `<div class="akt-report-chips">${filterChips.join('')}</div>`
                : ''
        }
        <div class="akt-report-breakdown__chart">
          ${breakdownHtml}
        </div>
      </div>
    </section>

    <section class="akt-panel akt-panel--calm akt-report-list">
      <div class="akt-report-list__head">
        <h3>Jahresliste</h3>
        <p class="muted">${yearItems.length}${
            hasListFilter ? ` von ${allYear.length}` : ''
        } ${yearItems.length === 1 ? 'Eintrag' : 'Einträge'}</p>
      </div>
      ${
          yearItems.length
              ? renderTable(yearItems, state)
              : `<p class="muted">${
                    hasListFilter
                        ? 'Keine Einträge für diesen Filter.'
                        : 'Keine Einträge in diesem Schuljahr.'
                }</p>`
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
        hint: 'Schulkalender: alle eingereichten/genehmigten Termine · Antippen für Details (fremde nur lesen) · ICS = Filteransicht für Outlook.',
        actionsHtml:
            '<button type="button" class="btn akt-btn-ghost akt-btn-compact" id="aktBtnCalIcsHeader" title="Sichtbare Termine als ICS herunterladen">' +
            '<i class="bi bi-download" aria-hidden="true"></i>ICS herunterladen</button>'
    });

    return `
    <section class="akt-panel akt-cal-layout">
      <h2 class="akt-sr-only">Kalender</h2>
      ${filterBar(state, {
          ...calendarScopeOpts(state),
          calendarPresets: true,
          showToday: true,
          showIcsDownload: true
      })}
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
    const metaParts = [
        formatDeDate(it.startdatum),
        zeit,
        teilnahmeKurz(it, labels.klasse),
        it.ort || '',
        lehrer ? `Verantwortlich: ${lehrer}` : '',
        it.kvStellungnahme ? `KV: ${it.kvStellungnahme}` : ''
    ].filter(Boolean);
    return `<button type="button" class="fb-rowbtn" data-akt-detail="${esc(it.aktivitaetId || it.itemId)}">
      <div class="fb-rowbtn__main">
        <div class="fb-rowbtn__title">${esc(it.titel || 'Ohne Titel')}</div>
        <div class="fb-rowbtn__meta">${esc(metaParts.join(' · ') || 'Aktivität')}</div>
      </div>
      <div class="fb-rowbtn__date">${esc(formatDeDate(it.startdatum))}</div>
      ${statusBadge(it.status)}
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
    const total = open.length + kvPending.length;
    const headline =
        total === 1
            ? '<span class="fb-num">1</span> Antrag ohne Entscheidung.'
            : `<span class="fb-num">${total}</span> ohne Entscheidung.`;
    return `
    <section class="akt-panel akt-panel--calm akt-panel--list">
      <h1 class="fb-greet">${headline}</h1>
      <p class="fb-sub">Zeile anklicken, Details prüfen, rechts im Drawer entscheiden.</p>
      ${
          hasAktKvCapability(state) || kvPending.length
              ? `<div class="fb-sec-label">KV-STELLUNGNAHME${kvPending.length ? ` · ${kvPending.length}` : ''}</div>
      ${
          kvPending.length
              ? `<div class="fb-rowlist">${kvPending
                    .map((it) => renderEntscheidungRow(it, state, labels))
                    .join('')}</div>`
              : '<p class="muted">Keine Anträge warten auf KV-Stellungnahme.</p>'
      }`
              : ''
      }
      ${
          canDecide(state)
              ? `<div class="fb-sec-label" id="aktEntscheidungenOpen">OFFENE ANTRÄGE${
                    open.length ? ` · ${open.length}` : ''
                }</div>
      ${
          open.length
              ? `<div class="fb-rowlist">${open
                    .map((it) => renderEntscheidungRow(it, state, labels))
                    .join('')}</div>`
              : '<p class="akt-quiet-status"><i class="bi bi-check-circle" aria-hidden="true"></i><span>Keine offenen Anträge.</span></p>'
      }`
              : ''
      }
    </section>`;
}

function renderKostenLimits(state) {
    const limits = resolveKostenLimits(state.rules);
    const canEdit = isAktStaffRole(state.role);
    const dirty = !!state.limitsDirty;
    const flash = state.settingsSavedFlash === 'limits';
    const cards = limits
        .map((l, idx) => {
            const ok = Math.max(0, Number(l.okBis) || 0);
            const warn = Math.max(ok, Number(l.warnBis) || 0);
            const title = typLabel(l.typ);
            const scale = `<div class="akt-threshold-scale" aria-hidden="true">
          <span class="akt-threshold-scale__ok" style="flex:${Math.max(ok, 1)}"></span>
          <span class="akt-threshold-scale__warn" style="flex:${Math.max(warn - ok, 1)}"></span>
          <span class="akt-threshold-scale__teuer" style="flex:${Math.max(Math.round(warn * 0.35) || 1, 1)}"></span>
        </div>
        <div class="akt-threshold-scale__legend">
          <span><i class="akt-dot akt-dot--ok" aria-hidden="true"></i>bis ${esc(ok)} €</span>
          <span><i class="akt-dot akt-dot--warn" aria-hidden="true"></i>bis ${esc(warn)} €</span>
          <span><i class="akt-dot akt-dot--teuer" aria-hidden="true"></i>darüber Prüfung</span>
        </div>`;
            if (!canEdit) {
                return `<article class="akt-threshold-card">
          <h3>${esc(title)}</h3>
          ${scale}
        </article>`;
            }
            return `<article class="akt-threshold-card">
          <h3 class="akt-threshold-card__title">${esc(title)}</h3>
          <input type="hidden" name="limitTyp" data-idx="${idx}" value="${esc(l.typ)}">
          ${scale}
          <div class="akt-threshold-card__grid">
            <label class="akt-threshold-field">
              <span class="akt-threshold-field__label">Ohne Rücksprache bis</span>
              <span class="akt-money-input"><input type="number" min="0" step="1" name="limitOk" data-idx="${idx}" value="${esc(
                l.okBis
            )}"><span class="akt-money-input__suffix">€</span></span>
            </label>
            <label class="akt-threshold-field">
              <span class="akt-threshold-field__label">Rücksprache bis</span>
              <span class="akt-money-input"><input type="number" min="0" step="1" name="limitWarn" data-idx="${idx}" value="${esc(
                l.warnBis
            )}"><span class="akt-money-input__suffix">€</span></span>
            </label>
          </div>
        </article>`;
        })
        .join('');
    return `
    <section class="akt-panel akt-panel--calm akt-limits">
      <header class="akt-limits__head">
        <h2 class="akt-sr-only">Kostenschwellen</h2>
        <p class="akt-limits__lead">Eine Schwelle pro Antrags-Art (wie im Formular) – Ampel im Kostenfeld, <strong>keine Zahlung</strong>, kein Blockieren.</p>
        <p class="muted akt-limits__note">Exkursion, Schulaktivität, Veranstaltung, Klassenabwesenheit, Sonstiges – Beträge hier pflegen.</p>
      </header>
      <form id="aktKostenLimitsForm" data-akt-dirty-form="limits">
        <div class="akt-threshold-grid">${cards}</div>
        ${
            canEdit
                ? `<div class="akt-actions" style="margin-top:16px">
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
    const tabRaw = String(state.systemTab || 'site').toLowerCase();
    const tab =
        tabRaw === 'perms' ||
        tabRaw === 'kalender' ||
        tabRaw === 'wartung' ||
        tabRaw === 'danger'
            ? tabRaw
            : 'site';
    const tabBtn = (id, label, opts) => {
        const danger = !!(opts && opts.danger);
        const active = tab === id;
        return `<button type="button" class="akt-sys-tabs__btn${active ? ' is-active' : ''}${
            danger ? ' akt-sys-tabs__btn--danger' : ''
        }" role="tab" id="aktSysTab-${id}" data-akt-sys-tab="${id}" aria-selected="${
            active ? 'true' : 'false'
        }" aria-controls="aktSysPanel-${id}">${esc(label)}</button>`;
    };
    const panelOpen = (id) => (tab === id ? '' : ' hidden');
    return `
    <div class="akt-sys">
      <div class="akt-sys-tabs" role="tablist" aria-label="Systemadministration">
        ${tabBtn('site', '1 · Datenquelle')}
        ${tabBtn('perms', '2 · Berechtigungen')}
        ${tabBtn('kalender', '3 · Kalender')}
        ${tabBtn('wartung', '4 · Wartung')}
        ${tabBtn('danger', 'Danger Zone', { danger: true })}
      </div>

      <section class="akt-sys-card" id="aktSysPanel-site" role="tabpanel" aria-labelledby="aktSysTab-site"${panelOpen(
          'site'
      )}>
        <header class="akt-sys-card__head">
          <span class="akt-sys-card__step">1</span>
          <div>
            <h3>Datenquelle</h3>
            <p>SharePoint-Site verbinden – Grundlage für Listen, Genehmiger und Kalender.</p>
          </div>
        </header>
        ${
            noSite
                ? `<div class="akt-inline-warn" role="status">
          <strong>Keine SharePoint-URL in diesem Browser.</strong>
          <p class="muted" style="margin:6px 0 0">IT: unter Admin → Schule die <em>Tool-Site-URL</em> setzen, oder hier eintragen und „Verbindung prüfen“.</p>
        </div>`
                : ''
        }
        <div class="akt-sys-health" role="status">
          <span class="akt-status akt-status--${statusMod}"><span class="akt-status__dot" aria-hidden="true"></span>${esc(statusLabel)}</span>
          <span class="akt-status akt-status--${liveOk ? 'ok' : 'warn'}"><span class="akt-status__dot" aria-hidden="true"></span>${esc(liveLine)}</span>
          <span class="akt-status akt-status--${aktOk ? 'ok' : aktWarn ? 'warn' : 'warn'}"><span class="akt-status__dot" aria-hidden="true"></span>${esc(aktLine)}</span>
        </div>
        ${
            liveWarn
                ? `<div class="akt-inline-warn" role="status">
          <strong>Klassenlisten fehlen oder sind leer.</strong>
          <p class="muted" style="margin:6px 0 0">Unter „Listen verwalten“ Stammdaten befüllen oder Paket anlegen.</p>
        </div>`
                : ''
        }
        <div class="akt-sys-field">
          <label for="aktSiteUrl">SharePoint-URL</label>
          <div class="akt-sys-field__row">
            <input type="url" id="aktSiteUrl" class="akt-sys-input" value="${esc(state.siteUrl)}" placeholder="https://…sharepoint.com/sites/…" spellcheck="false" autocomplete="off">
            <button type="button" class="btn akt-btn-primary" id="aktBtnLoad"><i class="bi bi-plug"></i>Verbindung prüfen</button>
          </div>
        </div>
        <div class="akt-sys-links">
          <a class="btn akt-btn-secondary" href="sharepoint-liste-schulaktivitaeten.html"><i class="bi bi-list-ul"></i>Listen verwalten</a>
          <a class="btn akt-btn-ghost" href="schulaktivitaeten-setup.html"><i class="bi bi-lightning-charge"></i>Flow-Setup</a>
          <a class="btn akt-btn-ghost" href="../tenant.html"><i class="bi bi-gear"></i>Stammdaten</a>
        </div>
      </section>

      <section class="akt-sys-card" id="aktSysPanel-perms" role="tabpanel" aria-labelledby="aktSysTab-perms"${panelOpen(
          'perms'
      )}>
        <header class="akt-sys-card__head">
          <span class="akt-sys-card__step">2</span>
          <div>
            <h3>Berechtigungen</h3>
            <p>Genehmiger und Einreicher (Lehrkräfte) – sonst sehen Lehrer nach „Listenrechte“ die Liste nicht.</p>
          </div>
        </header>
        <div class="akt-sys-callout" role="note">
          <strong>Speichern</strong> schreibt die App-Rollen.
          <strong>Listenrechte anwenden</strong> setzt SharePoint-Rechte (Vererbung bricht) – Schulleitung <em>und</em> Einreicher-Gruppe müssen gesetzt sein.
        </div>
        <div id="aktSysPermAdminMount" class="akt-sys-perm-mount"></div>
        <div id="aktSysPermLehrerMount" class="akt-sys-perm-mount" style="margin-top:12px"></div>
        <div class="akt-sys-actions">
          <button type="button" class="btn akt-btn-primary" id="aktSysPermSave"><i class="bi bi-check2"></i>Berechtigungen speichern</button>
          <button type="button" class="btn akt-btn-secondary" id="aktSysPermApply"><i class="bi bi-shield-check"></i>Listenrechte anwenden</button>
          <a class="btn akt-btn-ghost" href="sharepoint-liste-schulaktivitaeten.html"><i class="bi bi-box-arrow-up-right"></i>Listen-Assistent</a>
        </div>
        <p class="akt-sys-flash muted" id="aktSysPermStatus" hidden></p>
      </section>

      <section class="akt-sys-card" id="aktSysPanel-kalender" role="tabpanel" aria-labelledby="aktSysTab-kalender"${panelOpen(
          'kalender'
      )}>
        <header class="akt-sys-card__head">
          <span class="akt-sys-card__step">3</span>
          <div>
            <h3>Kalenderintegration</h3>
            <p>Genehmigte Termine als ICS oder abonnierbarer Feed.</p>
          </div>
        </header>
        ${renderFeedBar(state)}
        <div class="akt-sys-actions">
          <button type="button" class="btn akt-btn-secondary" id="aktBtnFeed"><i class="bi bi-rss"></i>Feed verwalten</button>
          <button type="button" class="btn akt-btn-ghost" id="aktBtnIcs"><i class="bi bi-download"></i>ICS herunterladen</button>
        </div>
      </section>

      <section class="akt-sys-card" id="aktSysPanel-wartung" role="tabpanel" aria-labelledby="aktSysTab-wartung"${panelOpen(
          'wartung'
      )}>
        <header class="akt-sys-card__head">
          <span class="akt-sys-card__step">4</span>
          <div>
            <h3>Wartung &amp; Diagnose</h3>
            <p>Backup, Demo und Hilfe – getrennt vom Schulalltag.</p>
          </div>
        </header>
        <div class="akt-sys-split">
          <div class="akt-sys-split__block">
            <h4>Datenimport</h4>
            <p class="muted">JSON-Backup (Absicherung, kein Ersatz für Live-Listen).</p>
            <label class="btn akt-btn-secondary" for="aktImportJson"><i class="bi bi-upload"></i>JSON importieren</label>
            <input type="file" id="aktImportJson" accept=".json,application/json" hidden>
          </div>
          <div class="akt-sys-split__block">
            <h4>Demo &amp; Hilfe</h4>
            <p class="muted">Vorführung und Fehlersuche.</p>
            <div class="akt-sys-actions">
              <button type="button" class="btn akt-btn-secondary" id="aktBtnDemo"><i class="bi bi-database"></i>Demo-Daten laden</button>
              <a class="btn akt-btn-ghost" href="../hilfe.html#tool-schulaktivitaeten-planer"><i class="bi bi-life-preserver"></i>Hilfe</a>
            </div>
          </div>
        </div>
      </section>

      <section class="akt-sys-card akt-sys-card--danger" id="aktSysPanel-danger" role="tabpanel" aria-labelledby="aktSysTab-danger"${panelOpen(
          'danger'
      )}>
        <header class="akt-sys-card__head">
          <span class="akt-sys-card__step akt-sys-card__step--danger" aria-hidden="true"><i class="bi bi-exclamation-triangle"></i></span>
          <div>
            <h3>Danger Zone</h3>
            <p>Disruptive Aktionen – nur bewusst ausführen.</p>
          </div>
        </header>
        <button type="button" class="btn akt-btn-danger" id="aktBtnDemoReset"><i class="bi bi-trash"></i>Demo-Daten zurücksetzen</button>
      </section>
    </div>`;
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
