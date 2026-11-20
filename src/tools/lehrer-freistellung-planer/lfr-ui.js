/**
 * UI: Freistellungen Lehrkräfte (approval-planer Calm, Vorbild Schulaktivitäten / Schüler-FR).
 */
import {
    navSectionsForRole,
    visibleItemsForState,
    canDecide,
    countOpenAusstehend
} from './lfr-state.js';
import { KATEGORIE_CHOICES as KAT } from './lfr-schema.js';
import {
    filterItems,
    formatDeDateTimeRange,
    formatDisplayTextWithDeDates,
    normalizeStatus
} from './lfr-logic.js';
import {
    LFR_ROLE_ORDER,
    canUseLfrRoleSwitcher,
    isLfrDemoRoleUiEnabled
} from './lfr-entra-role.js';
import { renderLfrDrawer } from './lfr-detail-ui.js';
import { buildLfrCalendarEvents, renderLfrCalendarShell } from './lfr-calendar.js';
import {
    normalizeSystemTab,
    renderSystemTabBar,
    systemPanelHiddenAttr
} from '../../shared/system-admin-shell.js';
import {
    loadLfrPermissions,
    resolveLfrApproverEmail,
    lfrPermsStepConfigured
} from './lfr-permissions.js';

function esc(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function statusBadge(status) {
    const s = normalizeStatus(status);
    const mod = s === 'Genehmigt' ? 'ok' : s === 'Abgelehnt' ? 'bad' : 'warn';
    return `<span class="lfr-status lfr-status--${mod}"><span class="lfr-status__dot" aria-hidden="true"></span>${esc(
        s
    )}</span>`;
}

function roleSubtitle(role) {
    return role === 'direktion' ? 'Verwaltung' : 'Lehrkraft';
}

function topTitle(state) {
    if (state.view === 'einrichtung') return 'Systemadministration';
    if (state.view === 'freigabe') return 'Entscheidungen';
    if (state.view === 'export') return 'Kalender-Export';
    if (state.view === 'antrag') return 'Antrag stellen';
    if (state.view === 'meine') return 'Meine Anträge';
    if (state.view === 'liste') return 'Alle Anträge';
    if (state.view === 'kalender') return 'Kalender';
    return 'Übersicht';
}

function renderLfrRoleBlock(state) {
    const demoRoleUi = isLfrDemoRoleUiEnabled(state.entraGroupsConfigured, state.demoRoleOverride);
    const showRoleSwitcher = canUseLfrRoleSwitcher(state.planerRoles || [], demoRoleUi);
    if (!showRoleSwitcher) {
        if (state.planerAccessDenied) {
            return `<p class="fr-nav__hint">${esc(state.roleHint || 'Kein Zugang.')}</p>`;
        }
        return '';
    }
    return (
        '<span class="fr-nav__role-label">Ansicht (Vorschau)</span>' +
        '<div class="fr-nav__role-btns" role="group" aria-label="Rolle wählen">' +
        LFR_ROLE_ORDER.map(
            (r) =>
                `<button type="button" class="fr-chip${state.role === r ? ' is-active' : ''}" data-lfr-role="${esc(
                    r
                )}">${r === 'direktion' ? 'Direktion' : 'Lehrkraft'}</button>`
        ).join('') +
        '</div>' +
        '<p class="fr-nav__hint">IT-Vorschau (?demoRole=1).</p>'
    );
}

function renderMainTop(state) {
    const onSystem = state.view === 'einrichtung';
    const lead = onSystem
        ? 'Site, Berechtigungen, Kalender und Wartung in Tabs – Danger Zone bewusst getrennt.'
        : state.role === 'direktion'
          ? 'Entscheidungen, Übersicht und Kalender.'
          : 'Anträge stellen und Status verfolgen.';
    return `
        <header class="fr-top fr-top--student">
          <div class="fr-top__copy">
            <h2 class="fr-top__student-title">${esc(topTitle(state))}</h2>
            <p class="fr-top__lead">${esc(lead)}</p>
          </div>
          <div class="fr-top__user">
            ${
                state.accountEmail && !onSystem
                    ? `<button type="button" class="btn" id="lfrBtnRefresh" title="Aktualisieren"><i class="bi bi-arrow-repeat"></i>Aktualisieren</button>`
                    : ''
            }
          </div>
        </header>`;
}

function dayCount(beginn, ende) {
    const b = beginn ? new Date(beginn) : null;
    const e = ende ? new Date(ende) : null;
    if (!b || !e || isNaN(b) || isNaN(e)) return null;
    return Math.max(1, Math.round((e.getTime() - b.getTime()) / 86400000) + 1);
}

/**
 * @param {object} state
 * @param {HTMLElement} root
 */
export function renderApp(state, root) {
    const sections = navSectionsForRole(state.role);
    const openCount = countOpenAusstehend(state.items);
    const roleBlock = renderLfrRoleBlock(state);
    const navHtml = sections
        .map((sec) => {
            const label = sec.label
                ? `<span class="fr-nav__section-label">${esc(sec.label)}</span>`
                : '';
            const btns = (sec.items || [])
                .map((v) => {
                    const badge =
                        v.id === 'freigabe' && openCount > 0
                            ? `<span class="fr-nav__badge">${openCount}</span>`
                            : '';
                    return `
            <button type="button" class="fr-nav__btn${state.view === v.id ? ' is-active' : ''}${
                v.id === 'antrag' ? ' fr-nav__btn--cta' : ''
            }" data-lfr-view="${esc(v.id)}">
              <i class="bi ${esc(v.icon)}" aria-hidden="true"></i><span class="fr-nav__btn-label">${esc(
                        v.label
                    )}</span>${badge}
            </button>`;
                })
                .join('');
            return `<div class="fr-nav__section">${label}${btns}</div>`;
        })
        .join('');

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
          ${navHtml}
        </nav>
        ${roleBlock ? `<div class="fr-nav__role">${roleBlock}</div>` : ''}
      </aside>
      <div class="fr-main">
        ${renderMainTop(state)}
        ${
            state.localDemoOnly && state.view === 'einrichtung'
                ? '<div class="fr-alert fr-alert--info" role="status">Lokaler Demo-Modus – SharePoint unter Systemadministration verbinden oder IT-Setup öffnen.</div>'
                : ''
        }
        ${state.error ? `<div class="fr-alert fr-alert--bad" role="alert">${esc(state.error)}</div>` : ''}
        ${state.loading ? '<div class="fr-skel" aria-busy="true"><div class="fr-skel__card"></div><div class="fr-skel__card"></div></div>' : ''}
        <div class="fr-view${state.loading ? ' is-loading' : ''}">
          ${state.loading || state.planerAccessDenied ? '' : renderView(state)}
        </div>
      </div>
    </div>
    ${state.detailId ? renderLfrDrawer(state) : ''}`;
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
    const items = visibleItemsForState(state, {});
    const offen = items.filter((i) => normalizeStatus(i.status) === 'Ausstehend');
    const demnaechst = items
        .filter((i) => normalizeStatus(i.status) === 'Genehmigt')
        .filter((i) => {
            const b = i.beginn ? new Date(i.beginn) : null;
            if (!b || isNaN(b)) return false;
            const now = new Date();
            const in14 = new Date(now.getTime());
            in14.setDate(in14.getDate() + 14);
            return b >= now && b <= in14;
        })
        .slice(0, 5);

    if (state.role === 'direktion') {
        const banner = offen.length
            ? `<section class="fr-calm-banner fr-calm-banner--action" role="status">
      <div class="fr-calm-banner__text">
        <p class="fr-calm-banner__eyebrow">Was braucht deine Aufmerksamkeit?</p>
        <p><strong>${offen.length}</strong> Antrag${offen.length === 1 ? '' : 'e'} wartet auf Entscheidung.</p>
      </div>
      <button type="button" class="btn fr-btn-primary" data-lfr-view-jump="freigabe">Zu Entscheidungen</button>
    </section>`
            : `<p class="fr-quiet-status" role="status">
      <i class="bi bi-check-circle" aria-hidden="true"></i>
      <span>Keine offenen Entscheidungen – alles erledigt.</span>
    </p>`;
        return `
    ${banner}
    <section class="fr-panel fr-panel--calm">
      <div class="fr-section-head">
        <h3>Offene Anträge</h3>
        <button type="button" class="lfr-text-link" data-lfr-view-jump="liste">Alle Anträge</button>
      </div>
      ${
          offen.length
              ? renderTable(offen.slice(0, 8))
              : '<p class="fr-quiet-status" role="status"><i class="bi bi-inbox" aria-hidden="true"></i><span>Keine offenen Anträge.</span></p>'
      }
    </section>
    <section class="fr-panel fr-panel--calm">
      <div class="fr-section-head">
        <h3>Demnächst genehmigt</h3>
        <button type="button" class="lfr-text-link" data-lfr-view-jump="kalender">Kalender</button>
      </div>
      ${
          demnaechst.length
              ? renderTable(demnaechst)
              : '<p class="muted">In den nächsten 14 Tagen keine genehmigten Freistellungen.</p>'
      }
    </section>`;
    }

    const meineOffen = filterItems(
        state.items,
        { status: 'Ausstehend' },
        { scopeAll: false, accountEmail: state.accountEmail }
    );
    const banner = meineOffen.length
        ? `<section class="fr-calm-banner fr-calm-banner--action" role="status">
      <div class="fr-calm-banner__text">
        <p class="fr-calm-banner__eyebrow">Deine offenen Anträge</p>
        <p><strong>${meineOffen.length}</strong> wartet auf die Direktion (Microsoft Approvals).</p>
      </div>
      <button type="button" class="btn" data-lfr-view-jump="meine">Meine Anträge</button>
    </section>`
        : `<section class="fr-calm-banner fr-calm-banner--action" role="status">
      <div class="fr-calm-banner__text">
        <p class="fr-calm-banner__eyebrow">Schnell starten</p>
        <p>Neuen Freistellungsantrag stellen – Genehmigung über Microsoft Approvals.</p>
      </div>
      <button type="button" class="btn fr-btn-primary" data-lfr-view-jump="antrag"><i class="bi bi-plus-lg"></i>Antrag stellen</button>
    </section>`;

    return `
    ${banner}
    <section class="fr-panel fr-panel--calm">
      <div class="fr-section-head">
        <h3>Als Nächstes</h3>
      </div>
      <div class="fr-actions">
        <button type="button" class="btn btn-success" data-lfr-view-jump="antrag"><i class="bi bi-plus-lg"></i>Neuer Antrag</button>
        <button type="button" class="btn" data-lfr-view-jump="meine"><i class="bi bi-person"></i>Meine Anträge</button>
        <button type="button" class="btn" data-lfr-view-jump="kalender"><i class="bi bi-calendar3"></i>Mein Kalender</button>
      </div>
    </section>`;
}

function filterBar(state) {
    const scopeOpts =
        state.view === 'meine' || state.role !== 'direktion'
            ? { scopeAll: false }
            : { scopeAll: true };
    const count = visibleItemsForState(state, state.filters, scopeOpts).length;
    const katOpts = KAT.map(
        (k) =>
            `<option value="${esc(k)}"${state.filters.kategorie === k ? ' selected' : ''}>${esc(
                k
            )}</option>`
    ).join('');
    return `
    <div class="lfr-listbar">
      <select class="lfr-listbar__select" id="lfrFilterStatus" aria-label="Status">
        <option value="">Alle Status</option>
        <option value="Ausstehend"${state.filters.status === 'Ausstehend' ? ' selected' : ''}>Ausstehend</option>
        <option value="Genehmigt"${state.filters.status === 'Genehmigt' ? ' selected' : ''}>Genehmigt</option>
        <option value="Abgelehnt"${state.filters.status === 'Abgelehnt' ? ' selected' : ''}>Abgelehnt</option>
      </select>
      <select class="lfr-listbar__select" id="lfrFilterKat" aria-label="Kategorie">
        <option value="">Alle Kategorien</option>${katOpts}
      </select>
      <input type="search" class="lfr-listbar__search" id="lfrFilterQ" placeholder="Suchen …" aria-label="Suchen" value="${esc(
          state.filters.q || ''
      )}">
      <span class="lfr-listbar__meta">${count} ${count === 1 ? 'Eintrag' : 'Einträge'}</span>
    </div>`;
}

function renderListe(state, opts) {
    const mine = opts && opts.mine;
    const items = visibleItemsForState(
        state,
        state.filters,
        mine || state.role !== 'direktion' ? { scopeAll: false } : { scopeAll: true }
    );
    let emptyHtml =
        '<p class="fr-quiet-status" role="status"><i class="bi bi-inbox" aria-hidden="true"></i><span>Keine Einträge.</span></p>';
    if (
        !mine &&
        !items.length &&
        state.role === 'direktion' &&
        state.siteUrl &&
        !state.localDemoOnly &&
        !(state.filters && (state.filters.status || state.filters.kategorie || state.filters.q))
    ) {
        emptyHtml = `<div class="fr-quiet-status" role="status" style="display:block;text-align:left;max-width:40rem;">
      <p><i class="bi bi-inbox" aria-hidden="true"></i> <strong>Keine Anträge sichtbar.</strong></p>
      <p class="muted" style="margin:8px 0 12px;line-height:1.45;">
        Wenn Lehrkräfte bereits Einträge haben, fehlen der Direktion oft die SharePoint-Listenrechte
        (nach Vererbungsbruch). In der Systemadministration „Berechtigungen anwenden“ bzw. im IT-Setup
        Schritt&nbsp;4 speichern – Schulleitung/Genehmiger werden dabei mitberechtigt.
      </p>
      <button type="button" class="btn" data-lfr-view-jump="einrichtung"><i class="bi bi-shield-lock"></i>Zur Systemadministration</button>
    </div>`;
    }
    return `
    <section class="fr-panel fr-panel--list">
      <h2 class="lfr-sr-only">${mine ? 'Meine Anträge' : 'Alle Anträge'}</h2>
      ${filterBar(state)}
      ${items.length ? renderTable(items) : emptyHtml}
    </section>`;
}

function renderTable(items) {
    const rows = items
        .map((it) => {
            const days = dayCount(it.beginn, it.ende);
            const id = it.antragId || it.itemId;
            return `<tr class="lfr-table__row" tabindex="0" role="link" data-lfr-detail="${esc(
                id
            )}" aria-label="${esc((it.lehrerName || formatDisplayTextWithDeDates(it.titel) || 'Antrag') + ' öffnen')}">
          <td><strong>${esc(it.lehrerName || formatDisplayTextWithDeDates(it.titel) || it.lehrerEmail)}</strong></td>
          <td>${esc(formatDeDateTimeRange(it.beginn, it.ende))}</td>
          <td>${days != null ? days : '–'}${
                days != null && days > 1 ? ' <span class="fr-badge fr-badge--info">mehr</span>' : ''
            }</td>
          <td>${statusBadge(it.status)}</td>
          <td>${esc(it.kategorie)}</td>
          <td>${esc(formatDisplayTextWithDeDates(it.titel))}</td>
          <td class="lfr-table__chevron"><i class="bi bi-chevron-right" aria-hidden="true"></i></td>
        </tr>`;
        })
        .join('');
    return `
    <div class="fr-table-wrap">
      <table class="fr-table lfr-table lfr-table--rows">
        <thead>
          <tr>
            <th>Lehrkraft</th><th>Zeitraum</th><th>Tage</th><th>Status</th><th>Kategorie</th><th>Titel</th>
            <th class="lfr-table__chevron"><span class="lfr-sr-only">Öffnen</span></th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

function renderKalender(state) {
    // Lehrkraft: immer nur eigene Anträge – nie Kollegium im Kalender
    const items = visibleItemsForState(
        state,
        state.filters,
        state.role === 'direktion' ? { scopeAll: true } : { scopeAll: false }
    );
    const events = buildLfrCalendarEvents(items);
    const icsAction = canDecide(state)
        ? `<button type="button" class="btn btn-sm" id="lfrBtnIcs" title="Genehmigte als iCal"><i class="bi bi-download"></i>ICS</button>`
        : '';
    return `
    <section class="fr-panel fr-panel--calm fr-cal-layout">
      ${filterBar(state)}
      ${renderLfrCalendarShell({
          mode: state.calMode || 'month',
          year: state.calYear,
          month: state.calMonth,
          weekAnchor: state.calWeekAnchor,
          events,
          actionsHtml: icsAction
      })}
    </section>`;
}

function renderAntrag(state) {
    const f = state.form || {};
    const lehrerName = String(f.lehrerName || state.accountName || '').trim();
    const lehrerEmail = String(f.lehrerEmail || state.accountEmail || '')
        .trim()
        .toLowerCase();
    const lockedStyle = 'background:var(--surface-2,#f4f6f8);';
    const katOpts = KAT.map(
        (k) => `<option value="${esc(k)}"${f.kategorie === k ? ' selected' : ''}>${esc(k)}</option>`
    ).join('');
    return `
    <section class="fr-panel fr-panel--calm lfr-antrag-panel">
      <div class="lfr-antrag">
        <header class="lfr-antrag__head">
          <h2>Freistellung beantragen</h2>
          <p class="muted">Nach dem Speichern startet (mit Flow) die Genehmigung durch die Direktion – Status erscheint hier und in der Liste.</p>
        </header>

        <form class="lfr-antrag__form" id="lfrAntragForm" autocomplete="off">
          <fieldset class="fr-form-section">
            <legend>Zeitraum</legend>
            <p class="fr-form-section__lead">Beginn und Ende der Abwesenheit.</p>
            <div class="fr-form-grid lfr-antrag__grid lfr-antrag__grid--dates">
              <label>Beginn *
                <input type="datetime-local" id="lfrFormBeginn" value="${esc(f.beginn)}" required step="300">
              </label>
              <label>Ende *
                <input type="datetime-local" id="lfrFormEnde" value="${esc(f.ende)}" required step="300">
              </label>
            </div>
          </fieldset>

          <fieldset class="fr-form-section">
            <legend>Inhalt</legend>
            <p class="fr-form-section__lead">Kurz und klar – die Direktion sieht das im Approvals-Schritt.</p>
            <div class="fr-form-grid lfr-antrag__grid">
              <label class="fr-form-span2">Titel *
                <input type="text" id="lfrFormTitel" value="${esc(f.titel)}" maxlength="200" required placeholder="z. B. Fortbildung HAK OÖ">
              </label>
              <label class="lfr-antrag__kat">Kategorie *
                <select id="lfrFormKat" required>${katOpts}</select>
              </label>
              <label class="fr-form-span2">Beschreibung
                <textarea id="lfrFormBesch" rows="3" placeholder="Kurz den Grund beschreiben…">${esc(
                    f.beschreibung
                )}</textarea>
              </label>
            </div>
          </fieldset>

          <fieldset class="fr-form-section">
            <legend>Antragsteller</legend>
            <div class="fr-form-grid lfr-antrag__grid">
              <label>Lehrkraft
                <input type="text" id="lfrFormName" value="${esc(lehrerName)}" readonly aria-readonly="true" style="${lockedStyle}">
              </label>
              <label>E-Mail
                <input type="text" id="lfrFormMail" value="${esc(lehrerEmail)}" readonly aria-readonly="true" style="${lockedStyle}" autocomplete="off">
              </label>
              <p class="muted fr-form-span2 lfr-antrag__hint">Übernommen aus Ihrer Anmeldung – gilt für diesen Antrag.</p>
            </div>
          </fieldset>

          <div class="fr-actions lfr-antrag__actions">
            <button type="button" class="btn btn-success" id="lfrFormSubmit"><i class="bi bi-send"></i>Antrag absenden</button>
            <button type="button" class="btn" id="lfrFormReset">Zurücksetzen</button>
          </div>
        </form>
      </div>
    </section>`;
}

function renderFreigabe(state) {
    const items = filterItems(state.items, { status: 'Ausstehend' }, { scopeAll: true });
    return `
    <section class="fr-panel fr-panel--calm">
      <h2>Entscheidungen</h2>
      <p class="muted">Manuelle Freigabe für Tests ohne Flow. In Produktion übernimmt Power Automate Approvals und schreibt Status zurück.</p>
      ${
          items.length
              ? `<div class="fr-table-wrap">
      <table class="fr-table lfr-table lfr-table--rows">
        <thead><tr><th>Zeitraum</th><th>Titel</th><th>Lehrkraft</th><th></th><th class="lfr-table__chevron"><span class="lfr-sr-only">Öffnen</span></th></tr></thead>
        <tbody>
          ${items
              .map((it) => {
                  const id = it.antragId || it.itemId;
                  return `<tr class="lfr-table__row" tabindex="0" role="link" data-lfr-detail="${esc(
                      id
                  )}" aria-label="${esc((formatDisplayTextWithDeDates(it.titel) || 'Antrag') + ' öffnen')}">
              <td>${esc(formatDeDateTimeRange(it.beginn, it.ende))}</td>
              <td><strong>${esc(formatDisplayTextWithDeDates(it.titel))}</strong></td>
              <td>${esc(it.lehrerName || it.lehrerEmail)}</td>
              <td class="fr-td-actions">
                <button type="button" class="btn btn-sm btn-success" data-lfr-approve="${esc(
                    id
                )}">Genehmigen</button>
                <button type="button" class="btn btn-sm alt" data-lfr-reject="${esc(
                    id
                )}">Ablehnen</button>
              </td>
              <td class="lfr-table__chevron"><i class="bi bi-chevron-right" aria-hidden="true"></i></td>
            </tr>`;
              })
              .join('')}
        </tbody>
      </table>
    </div>`
              : '<p class="fr-quiet-status" role="status"><i class="bi bi-check-circle" aria-hidden="true"></i><span>Keine offenen Anträge.</span></p>'
      }
    </section>`;
}

function renderExport(state) {
    const approved = filterItems(state.items, { status: 'Genehmigt' }, { scopeAll: true });
    const calUser = String(state.outlookCalendarUser || '')
        .trim()
        .toLowerCase();
    const me = String(state.accountEmail || '')
        .trim()
        .toLowerCase();
    const otherTarget = !!(calUser && me && calUser !== me);
    const calHint = calUser
        ? `Setup-Ziel: <code>${esc(calUser)}</code>${
              state.outlookCalendarId ? ' (eigene Kalender-ID)' : ''
          }${
              otherTarget
                  ? ' – Schreibzugriff nur, wenn dieser Kalender dir freigegeben ist (Global Admin reicht nicht).'
                  : ''
          }`
        : 'Kalender-Ziel noch nicht konfiguriert – im <a href="lehrer-freistellung-setup.html">IT-Setup</a>, oder unten „In meinen Kalender“.';
    return `
    <section class="fr-panel fr-panel--calm">
      <h2>Kalender-Export</h2>
      <p class="muted">Genehmigte Freistellungen als Datei oder per Microsoft Graph in einen Outlook-Kalender.</p>
      <p><strong>${approved.length}</strong> genehmigte Einträge.</p>
      <div class="fr-actions">
        <button type="button" class="btn btn-success" id="lfrBtnIcs"><i class="bi bi-download"></i>iCal herunterladen</button>
        <button type="button" class="btn" id="lfrBtnOutlookSyncMe"><i class="bi bi-person-check"></i>In meinen Kalender</button>
        <button type="button" class="btn" id="lfrBtnOutlookSync"${calUser ? '' : ' disabled'} title="${
            calUser ? 'Setup-Zielkalender' : 'Zuerst Kalender-Besitzer im IT-Setup setzen'
        }"><i class="bi bi-calendar-plus"></i>Setup-Ziel synchronisieren</button>
      </div>
      <p class="muted" style="margin-top:14px;line-height:1.45;">${calHint}</p>
    </section>`;
}

function renderEinrichtung(state) {
    const connected = !!(state.ctx && state.ctx.siteId) || (!!state.siteUrl && !state.error && !state.localDemoOnly);
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
    const teacherCount = ((state.stammdaten && state.stammdaten.teachers) || []).length;
    const itemCount = (state.items || []).length;
    const itemLine = state.localDemoOnly
        ? `Anträge: ${itemCount} Einträge (Demo lokal)`
        : itemCount
          ? `Anträge: ${itemCount} Einträge geladen`
          : 'Anträge: noch nicht geladen';
    const teacherLine = teacherCount
        ? `Stammdaten: ${teacherCount} Lehrkräfte`
        : 'Stammdaten: noch keine Lehrkräfte';
    const perms = loadLfrPermissions();
    const approver = resolveLfrApproverEmail(perms);
    const permsOk = lfrPermsStepConfigured(perms) || !!approver;
    const calUser = String(state.outlookCalendarUser || '').trim();
    const tab = normalizeSystemTab(state.systemTab);
    const noSite = !String(state.siteUrl || '').trim();
    const panelOpen = (id) => systemPanelHiddenAttr(tab, id);
    return `
    <div class="ms-sys">
      ${renderSystemTabBar({
          prefix: 'lfr',
          activeTab: tab,
          dataAttr: 'data-lfr-sys-tab',
          ariaLabel: 'Systemadministration Freistellungen'
      })}

      <section class="ms-sys-card" id="lfrSysPanel-site" role="tabpanel" aria-labelledby="lfrSysTab-site"${panelOpen(
          'site'
      )}>
        <header class="ms-sys-card__head">
          <span class="ms-sys-card__step">1</span>
          <div>
            <h3>Datenquelle</h3>
            <p>SharePoint-Liste für Lehrer-Freistellungen (getrennt von der Schüler-Liste).</p>
          </div>
        </header>
        ${
            noSite
                ? `<div class="ms-sys-inline-warn" role="status">
          <strong>Keine SharePoint-URL in diesem Browser.</strong>
          <p class="muted" style="margin:6px 0 0">Site hier eintragen oder im IT-Setup hinterlegen, dann „Verbindung prüfen“.</p>
        </div>`
                : ''
        }
        <div class="ms-sys-health" role="status">
          <span class="ms-sys-status ms-sys-status--${statusMod}"><span class="ms-sys-status__dot" aria-hidden="true"></span>${esc(
              statusLabel
          )}</span>
          <span class="ms-sys-status ms-sys-status--${teacherCount ? 'ok' : 'warn'}"><span class="ms-sys-status__dot" aria-hidden="true"></span>${esc(
              teacherLine
          )}</span>
          <span class="ms-sys-status ms-sys-status--${
              itemCount && !state.localDemoOnly ? 'ok' : 'warn'
          }"><span class="ms-sys-status__dot" aria-hidden="true"></span>${esc(itemLine)}</span>
        </div>
        <div class="ms-sys-field">
          <label for="lfrSetupSite">SharePoint-URL</label>
          <div class="ms-sys-field__row">
            <input type="url" id="lfrSetupSite" class="ms-sys-input" value="${esc(
                state.siteUrl
            )}" placeholder="https://…sharepoint.com/sites/…" spellcheck="false" autocomplete="off">
            <button type="button" class="btn ms-sys-btn-primary" id="lfrBtnReload"><i class="bi bi-plug"></i>Verbindung prüfen</button>
          </div>
        </div>
        <div class="ms-sys-field__grid">
          <div class="ms-sys-field">
            <label for="lfrSetupListName">Listentitel</label>
            <input type="text" id="lfrSetupListName" class="ms-sys-input" value="${esc(state.listName)}">
          </div>
          <div class="ms-sys-field">
            <label for="lfrSetupListId">Listen-ID (optional)</label>
            <input type="text" id="lfrSetupListId" class="ms-sys-input" value="${esc(state.listId)}">
          </div>
        </div>
        <div class="ms-sys-actions">
          <button type="button" class="btn ms-sys-btn-primary" id="lfrBtnSaveSetup"><i class="bi bi-save"></i>Speichern</button>
          <button type="button" class="btn ms-sys-btn-secondary" id="lfrBtnEnsureList"><i class="bi bi-list-plus"></i>Liste anlegen / prüfen</button>
        </div>
        <div class="ms-sys-links" style="margin-top:12px">
          <a class="btn ms-sys-btn-secondary" href="lehrer-freistellung-setup.html"><i class="bi bi-lightning-charge"></i>Flow-Setup</a>
          <a class="btn ms-sys-btn-ghost" href="../tenant.html"><i class="bi bi-gear"></i>Stammdaten</a>
        </div>
      </section>

      <section class="ms-sys-card" id="lfrSysPanel-perms" role="tabpanel" aria-labelledby="lfrSysTab-perms"${panelOpen(
          'perms'
      )}>
        <header class="ms-sys-card__head">
          <span class="ms-sys-card__step">2</span>
          <div>
            <h3>Berechtigungen</h3>
            <p>Genehmiger und Entra-Gruppen – detaillierte Listenrechte im IT-Setup.</p>
          </div>
        </header>
        <div class="ms-sys-callout" role="note">
          <strong>Direktion</strong> setzt den Genehmiger und kann Listenrechte nachziehen.
          Nach Vererbungsbruch braucht Schulleitung/Genehmiger explizit Listen-Vollzugriff – sonst bleiben „Alle Anträge“ leer.
          Beim Anwenden wird auch „nur eigene Elemente“ gesetzt (SharePoint Erweiterte Einstellungen), damit Lehrkräfte in der Liste nur ihre Anträge sehen.
        </div>
        <div class="ms-sys-health" role="status">
          <span class="ms-sys-status ms-sys-status--${permsOk ? 'ok' : 'warn'}"><span class="ms-sys-status__dot" aria-hidden="true"></span>${
              permsOk ? 'Zugang konfiguriert' : 'Gruppen/Genehmiger unvollständig'
          }</span>
          <span class="ms-sys-status ms-sys-status--${approver ? 'ok' : 'warn'}"><span class="ms-sys-status__dot" aria-hidden="true"></span>${
              approver ? 'Genehmiger: ' + esc(approver) : 'Kein Genehmiger gesetzt'
          }</span>
        </div>
        <dl class="ms-sys-dl">
          <dt>Lehrkräfte-Gruppe</dt>
          <dd>${esc(perms.groupLehrerName || perms.groupLehrerId || '–')}</dd>
          <dt>Direktion-Gruppe</dt>
          <dd>${esc(perms.groupDirektionName || perms.groupDirektionId || '– keine Gruppe, Zugriff über Person')}</dd>
          <dt>Direktion-Personen</dt>
          <dd>${esc(
              (perms.direktionUsers || [])
                  .map((u) => u.displayName || u.mail)
                  .filter(Boolean)
                  .join(', ') || '–'
          )}</dd>
        </dl>
        <div class="ms-sys-field">
          <label for="lfrSysApproverEmail">Genehmiger-E-Mail (Direktion / Flow)</label>
          <div class="ms-sys-field__row">
            <input type="email" id="lfrSysApproverEmail" class="ms-sys-input" value="${esc(
                approver
            )}" placeholder="direktion@schule.at" autocomplete="off">
            <button type="button" class="btn ms-sys-btn-primary" id="lfrBtnSaveApprover"><i class="bi bi-check2"></i>Speichern</button>
          </div>
        </div>
        <div class="ms-sys-actions">
          <button type="button" class="btn ms-sys-btn-primary" id="lfrBtnApplyListPerms" ${
              state.siteUrl ? '' : 'disabled'
          } title="Schulleitung, Genehmiger und Gruppen auf die SharePoint-Liste schreiben">
            <i class="bi bi-shield-lock"></i>Listenrechte anwenden
          </button>
          <a class="btn ms-sys-btn-secondary" href="lehrer-freistellung-setup.html"><i class="bi bi-shield-check"></i>IT-Setup · Berechtigungen</a>
        </div>
        <p id="lfrPermsApplyLog" class="muted" style="margin-top:10px;font-size:0.88em;line-height:1.4;" hidden></p>
      </section>

      <section class="ms-sys-card" id="lfrSysPanel-kalender" role="tabpanel" aria-labelledby="lfrSysTab-kalender"${panelOpen(
          'kalender'
      )}>
        <header class="ms-sys-card__head">
          <span class="ms-sys-card__step">3</span>
          <div>
            <h3>Kalenderintegration</h3>
            <p>Zielkalender für genehmigte Freistellungen (Outlook / Graph).</p>
          </div>
        </header>
        <div class="ms-sys-field">
          <label for="lfrSysCalUser">Kalender-Besitzer (UPN / E-Mail)</label>
          <input type="email" id="lfrSysCalUser" class="ms-sys-input" value="${esc(
              calUser
          )}" placeholder="kalender@schule.at" autocomplete="off">
        </div>
        <div class="ms-sys-field">
          <label for="lfrSysCalId">Kalender-ID (optional)</label>
          <input type="text" id="lfrSysCalId" class="ms-sys-input" value="${esc(
              state.outlookCalendarId || ''
          )}" placeholder="leer = Standardkalender" autocomplete="off">
        </div>
        <div class="ms-sys-actions">
          <button type="button" class="btn ms-sys-btn-primary" id="lfrBtnSaveCalendar"><i class="bi bi-check2"></i>Kalender speichern</button>
          <button type="button" class="btn ms-sys-btn-secondary" data-lfr-view-jump="export"><i class="bi bi-download"></i>Kalender-Export öffnen</button>
        </div>
      </section>

      <section class="ms-sys-card" id="lfrSysPanel-wartung" role="tabpanel" aria-labelledby="lfrSysTab-wartung"${panelOpen(
          'wartung'
      )}>
        <header class="ms-sys-card__head">
          <span class="ms-sys-card__step">4</span>
          <div>
            <h3>Wartung &amp; Diagnose</h3>
            <p>IT-Setup, Stammdaten und Hilfe – getrennt vom Schulalltag.</p>
          </div>
        </header>
        <div class="ms-sys-split">
          <div class="ms-sys-split__block">
            <h4>IT-Einrichtung</h4>
            <p class="muted">Vollständiger Assistent: Vorbereitung, Konten, Liste, Rechte, Flow.</p>
            <div class="ms-sys-actions">
              <a class="btn ms-sys-btn-secondary" href="lehrer-freistellung-setup.html"><i class="bi bi-wrench"></i>Lehrer-Freistellungen Setup</a>
            </div>
          </div>
          <div class="ms-sys-split__block">
            <h4>Stammdaten &amp; Hilfe</h4>
            <p class="muted">Klassen/Lehrkräfte und Dokumentation.</p>
            <div class="ms-sys-actions">
              <a class="btn ms-sys-btn-secondary" href="../tenant.html"><i class="bi bi-gear"></i>Stammdaten</a>
              <a class="btn ms-sys-btn-ghost" href="../hilfe.html#tool-lehrer-freistellung-planer"><i class="bi bi-life-preserver"></i>Hilfe</a>
            </div>
          </div>
        </div>
      </section>

      <section class="ms-sys-card ms-sys-card--danger" id="lfrSysPanel-danger" role="tabpanel" aria-labelledby="lfrSysTab-danger"${panelOpen(
          'danger'
      )}>
        <header class="ms-sys-card__head">
          <span class="ms-sys-card__step ms-sys-card__step--danger" aria-hidden="true"><i class="bi bi-exclamation-triangle"></i></span>
          <div>
            <h3>Danger Zone</h3>
            <p>Lokale Demo-Daten verwerfen – SharePoint bleibt unberührt.</p>
          </div>
        </header>
        <button type="button" class="btn ms-sys-btn-danger" id="lfrBtnDemoReset"><i class="bi bi-trash"></i>Lokale Demo-Daten zurücksetzen</button>
      </section>
    </div>`;
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
        siteUrl: String(($('lfrSetupSite') && $('lfrSetupSite').value) || state.siteUrl || '').trim(),
        listName: String(
            ($('lfrSetupListName') && $('lfrSetupListName').value) || state.listName || ''
        ).trim(),
        listId: String(($('lfrSetupListId') && $('lfrSetupListId').value) || state.listId || '').trim()
    };
}

export function readCalendarSetupFromDom(state) {
    const $ = (id) => document.getElementById(id);
    return {
        outlookCalendarUser: String(
            ($('lfrSysCalUser') && $('lfrSysCalUser').value) || state.outlookCalendarUser || ''
        )
            .trim()
            .toLowerCase(),
        outlookCalendarId: String(
            ($('lfrSysCalId') && $('lfrSysCalId').value) || state.outlookCalendarId || ''
        ).trim()
    };
}
