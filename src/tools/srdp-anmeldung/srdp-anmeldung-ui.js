/**
 * UI-Rendering sRDP-Anmeldung (Calm UX, sektioniertes Nav, Multi-Filter, Kalender).
 */
import {
    navSectionsForRole,
    emptyForm,
    itemsForCurrentRole,
    filterItemsForView,
    kvKlasseCodesForState,
    isSrdpStaffRole,
    findItemById,
    parseCodeList,
    countNacharbeitPending,
    klassenOhneAnmeldung,
    effectiveAccountEmail
} from './srdp-anmeldung-state.js';
import {
    isSeminarWahlfach,
    computeAnmeldungStats,
    groupItemsBy,
    buildAnmeldungSummary,
    teacherRolesOnItem,
    filterItemsForTeacher,
    buildPruefplanKurz,
    buildPruefplanDetail,
    isSrdpAnmeldeFensterOffen,
    findAnmeldungBySchuelerEmail
} from './srdp-anmeldung-logic.js';
import { periodenForTerminJahr, MATURA_TERMIN_URL, formatGermanDate } from './srdp-pruefungstermine.js';
import {
    SRDP_ROLE_ORDER,
    canUseSrdpRoleSwitcher,
    isSrdpDemoRoleUiEnabled,
    sortSrdpRoles
} from './srdp-anmeldung-entra-role.js';
import { renderAntragWizard } from './srdp-anmeldung-wizard-ui.js';
import { renderAnmeldungDrawer, renderSubmitSuccess } from './srdp-anmeldung-detail-ui.js';
import {
    renderMs365CalendarShell,
    toIsoDateOnly
} from '../../shared/ms365-calendar-shell.js';
import {
    normalizeSystemTab,
    renderSystemTabBar,
    systemPanelHiddenAttr
} from '../../shared/system-admin-shell.js';
import {
    loadEffectiveSrdpPermissionsConfig,
    formatSrdpAudienceSummary
} from '../sharepoint/sharepoint-liste-srdp-permissions-ui.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function optionsHtml(list, selected) {
    const sel = String(selected || '');
    return (Array.isArray(list) ? list : [])
        .map((v) => {
            const val = typeof v === 'object' ? String(v.label || v.code || '') : String(v);
            return `<option value="${esc(val)}"${val === sel ? ' selected' : ''}>${esc(val)}</option>`;
        })
        .join('');
}

function renderRoleBlock(state) {
    const demoRoleUi = isSrdpDemoRoleUiEnabled(state.entraGroupsConfigured, state.demoRoleOverride);
    const switchable = demoRoleUi ? SRDP_ROLE_ORDER : sortSrdpRoles(state.planerRoles || []);
    const show = canUseSrdpRoleSwitcher(state.planerRoles || [], demoRoleUi);
    if (!show) return '';
    const labels = { schueler: 'Schüler', lehrer: 'Lehrer', admin: 'Admin' };
    return (
        '<span class="srdp-nav__role-label">Ansicht (Vorschau)</span>' +
        '<div class="srdp-nav__role-btns">' +
        switchable
            .map(
                (r) =>
                    `<button type="button" class="srdp-chip${state.role === r ? ' is-active' : ''}" data-srdp-role="${esc(
                        r
                    )}">${esc(labels[r] || r)}</button>`
            )
            .join('') +
        '</div>'
    );
}

function topTitle(state) {
    const map = {
        dashboard: 'Übersicht',
        anmelden: 'sRDP-Anmeldung',
        meine: 'Meine Anmeldung',
        'meine-pruefungen': 'Meine Prüfungen',
        'meine-klasse': 'Meine Klasse',
        liste: 'Alle Anmeldungen',
        termine: 'sRDP-Termine',
        kalender: 'Kalender',
        auswertung: 'Auswertung',
        system: 'Systemadministration',
        einrichtung: 'Systemadministration'
    };
    return map[state.view] || 'sRDP-Anmeldung';
}

function topHint(state) {
    const jahr = (state.options && state.options.jahr) || '';
    const profile = (state.options && state.options.profileLabel) || 'HAK';
    if (state.view === 'system' || state.view === 'einrichtung') {
        return 'Site, Berechtigungen, Kalender und Wartung in Tabs – Danger Zone bewusst getrennt.';
    }
    if (state.role === 'schueler') {
        if (state.view === 'termine') {
            return 'Wann die zentralen Klausuren des Haupttermins ' + jahr + ' und der Nebentermine stattfinden.';
        }
        if (state.view === 'kalender') {
            return 'Deine schriftlichen Klausuren am Haupttermin ' + jahr + '.';
        }
        return 'Melde dich für den Haupttermin ' + jahr + ' an (' + profile + ').';
    }
    if (state.role === 'lehrer') {
        if (state.view === 'meine-klasse') {
            return 'Nur Anmeldungen deiner Klasse · ' + profile + ' ' + jahr;
        }
        if (state.view === 'termine') {
            return 'Zentrale Klausuren laut matura.gv.at · Haupttermin, Herbsttermin, Wintertermin.';
        }
        if (state.view === 'liste' || state.view === 'auswertung' || state.view === 'kalender') {
            return 'Alle Anmeldungen der Schule · ' + profile + ' ' + jahr;
        }
        return 'Prüfungen, in denen du als Lehrkraft eingetragen bist · ' + profile + ' ' + jahr;
    }
    return 'Entscheidungen, Nacharbeit und Auswertung · ' + profile + ' ' + jahr;
}

function phaseBadge(state) {
    const o = state.options || {};
    const offen = isSrdpAnmeldeFensterOffen(o);
    const frist = o.anmeldeFrist ? String(o.anmeldeFrist).slice(0, 10) : '';
    if (offen) {
        return `<span class="srdp-badge srdp-badge--ok" title="${
            frist ? 'offen bis ' + esc(frist) : 'keine Frist gesetzt'
        }">Anmeldung offen${frist ? ' bis ' + esc(frist) : ''}</span>`;
    }
    return `<span class="srdp-badge srdp-badge--warn">Anmeldung geschlossen${
        frist ? ' (Frist ' + esc(frist) + ')' : ''
    }</span>`;
}

function renderTop(state) {
    const selfTitled =
        state.view === 'dashboard' ||
        state.view === 'liste' ||
        state.view === 'meine' ||
        state.view === 'anmelden' ||
        state.view === 'auswertung';
    if (selfTitled) {
        return `<div class="srdp-top-refresh">${phaseBadge(state)}${
            state.localDemoOnly ? '<span class="fb-pill fb-pill--warn">Demo lokal</span>' : ''
        }</div>`;
    }
    return `<header class="srdp-top">
      <div>
        <h2>${esc(topTitle(state))}</h2>
        <p>${esc(topHint(state))}</p>
      </div>
      <div class="srdp-top__user">
        ${phaseBadge(state)}
        ${
            state.localDemoOnly
                ? '<span class="fb-pill fb-pill--warn">Demo lokal</span>'
                : ''
        }
        ${
            state.demoPersonaEmail
                ? `<small>Persona ${esc(state.demoPersonaName || state.demoPersonaEmail)}</small>`
                : ''
        }
      </div>
    </header>`;
}

function quietEmpty(text, actionHtml) {
    return `<p class="srdp-quiet-status" role="status">
      <i class="bi bi-check-circle" aria-hidden="true"></i>
      <span>${text}</span>
    </p>${actionHtml || ''}`;
}

function planDateHtml(row) {
    if (row && row.datum) return `<time datetime="${esc(row.datumIso)}">${esc(row.datum)}</time>`;
    if (row && row.datumHinweis) return `<time class="is-school">${esc(row.datumHinweis)}</time>`;
    return '';
}

function renderPlanSplit(item, profileId, terminJahr, bfkDatum) {
    const plan = buildPruefplanDetail(item, profileId, { terminJahr, bfkDatum });
    const col = (title, rows, mod) => {
        const body = (rows || [])
            .map((r) => {
                if (!r.part) return '';
                return `<li><strong>${esc(r.part)}</strong>${planDateHtml(r)}${
                    r.detail ? `<span>${esc(r.detail)}</span>` : ''
                }</li>`;
            })
            .join('');
        return `<section class="srdp-plan-col srdp-plan-col--${mod}">
        <h4>${esc(title)}</h4>
        <ul>${body || '<li><span>–</span></li>'}</ul>
      </section>`;
    };
    if (!plan.schriftlich.length && !plan.muendlich.length) return '';
    return `<div class="srdp-plan-split">
      ${col('Schriftlich', plan.schriftlich, 'schriftlich')}
      ${col('Mündlich', plan.muendlich, 'muendlich')}
    </div>`;
}

function renderSummaryCard(item, profileId, opts) {
    const o = opts || {};
    const summary = buildAnmeldungSummary(item, profileId);
    return `<button type="button" class="srdp-summary srdp-summary--clickable" data-srdp-open-detail="${esc(
        item.itemId
    )}">
      <div class="srdp-summary__head">
        <h3>${esc(o.heading || summary.title || 'Anmeldung')}</h3>
        <p>${esc(String(item.klasse || '–'))} · ${esc(String(item.variante || '–'))}</p>
      </div>
      ${renderPlanSplit(item, profileId, o.jahr || '', o.bfkDatum || '')}
    </button>`;
}

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
            return `<label class="srdp-filter-multi__item${hidden ? ' is-filtered-out' : ''}">
        <input type="checkbox" data-srdp-filter-multi-item="${esc(key)}" value="${esc(v)}"${
                checked ? ' checked' : ''
            }>
        <span>${esc(label)}</span>
      </label>`;
        })
        .filter(Boolean)
        .join('');
    return `<details class="srdp-filter-multi${n ? ' is-filtered' : ''}" data-srdp-filter-multi="${esc(
        key
    )}"${open ? ' open' : ''}>
      <summary class="srdp-filter-multi__summary">${esc(summary)}</summary>
      <div class="srdp-filter-multi__panel">
        <input type="search" class="srdp-filter-multi__search" data-srdp-filter-multi-q="${esc(
            key
        )}" value="${esc(o.query || '')}" placeholder="Suchen …" autocomplete="off">
        <div class="srdp-filter-multi__list">${rows || '<p class="muted">Keine Einträge</p>'}</div>
        ${
            n
                ? `<button type="button" class="srdp-text-link" data-srdp-filter-multi-clear="${esc(
                      key
                  )}">Auswahl leeren</button>`
                : ''
        }
      </div>
    </details>`;
}

function renderFilters(state) {
    const o = state.options || {};
    const klassen = (o.klassen || []).map((k) => ({ code: k, name: k }));
    const variants = (o.variants || []).map((v) => ({ code: v.label, name: v.label }));
    const lehrer = (o.lehrer || []).map((l) => ({ code: l, name: l }));
    const open = state.filterMultiOpen || {};
    const q = state.filterMultiQuery || {};
    return `<div class="srdp-filters">
      ${renderFilterMulti('klasse', klassen, state.filterKlasse, !!open.klasse, {
          allLabel: 'Alle Klassen',
          countLabel: 'Klassen',
          query: q.klasse || ''
      })}
      ${renderFilterMulti('variante', variants, state.filterVariante, !!open.variante, {
          allLabel: 'Alle Varianten',
          countLabel: 'Varianten',
          query: q.variante || '',
          showSearch: false
      })}
      ${renderFilterMulti('lehrer', lehrer, state.filterLehrer, !!open.lehrer, {
          allLabel: 'Alle Lehrkräfte',
          countLabel: 'Lehrkräfte',
          query: q.lehrer || ''
      })}
      <input type="search" placeholder="Suche Name…" value="${esc(
          state.filterQuery || ''
      )}" data-srdp-filter="query">
      ${
          state.filterKlasse || state.filterVariante || state.filterLehrer || state.filterQuery
              ? '<button type="button" class="srdp-text-link" data-srdp-filter-reset>Filter zurücksetzen</button>'
              : ''
      }
    </div>`;
}

function renderNacharbeitStatus(item) {
    const done = !!(item && item.nacharbeitErledigt);
    const mod = done ? 'ok' : 'warn';
    const label = done ? 'In der Prüfungsliste' : 'Noch offen';
    return `<span class="srdp-status srdp-status--${mod}"><span class="srdp-status__dot" aria-hidden="true"></span>${label}</span>`;
}

function renderGroupBar(state) {
    const cur = String(state.groupBy || '');
    const opts = [
        ['', 'Als Liste'],
        ['klasse', 'Nach Klasse'],
        ['variante', 'Nach Variante'],
        ['wahlfach', 'Nach Wahlfach']
    ];
    const count = filterItemsForView(state.items || [], state).length;
    return `<div class="srdp-listbar">
      <div class="srdp-seg" role="group" aria-label="Gruppieren">
        ${opts
            .map(
                ([id, label]) =>
                    `<button type="button" class="srdp-seg__btn${
                        cur === id ? ' is-active' : ''
                    }" data-srdp-group="${esc(id)}" aria-pressed="${cur === id ? 'true' : 'false'}">${esc(
                        label
                    )}</button>`
            )
            .join('')}
      </div>
      <input type="search" class="srdp-listbar__search" placeholder="Suche Name…" value="${esc(
          state.filterQuery || ''
      )}" data-srdp-filter="query" aria-label="Name suchen">
      <span class="srdp-listbar__meta">${count} ${count === 1 ? 'Eintrag' : 'Einträge'}</span>
    </div>`;
}

function renderFlatTable(items, state, opts) {
    const o = opts || {};
    const list = Array.isArray(items) ? items : [];
    const hide = String(o.hideCol || '');
    const showNach = o.showNacharbeit != null ? !!o.showNacharbeit : isSrdpStaffRole(state.role);
    if (!list.length) {
        return quietEmpty('<strong>Keine Anmeldungen</strong> · In dieser Ansicht ist nichts vorhanden.');
    }
    return `<div class="srdp-table-wrap"><table class="srdp-table srdp-table--rows">
      <thead><tr>
        <th>Name</th>${hide === 'klasse' ? '' : '<th>Klasse</th>'}${
            hide === 'variante' ? '' : '<th>Variante</th>'
        }<th>LFS</th>${hide === 'wahlfach' ? '' : '<th>Wahlfach</th>'}
        ${o.showRoles ? '<th>Deine Rolle</th>' : ''}
        <th>Betreuung DA</th>
        ${showNach ? '<th>Nacharbeit</th>' : ''}
        <th class="srdp-table__chevron"><span class="srdp-sr-only">Öffnen</span></th>
      </tr></thead>
      <tbody>
        ${list
            .map((it) => {
                const name = it.title || [it.nachname, it.vorname].filter(Boolean).join(', ') || 'Anmeldung';
                const roles = o.showRoles
                    ? teacherRolesOnItem(it, state.teacherLabel)
                          .map((r) => `<span class="srdp-pill srdp-pill--role">${esc(r.label)}</span>`)
                          .join('')
                    : '';
                return `<tr class="srdp-table__row" tabindex="0" role="link" data-srdp-open-detail="${esc(
                    it.itemId
                )}" aria-label="${esc(name + ' öffnen')}">
            <td><strong>${esc(name)}</strong></td>
            ${hide === 'klasse' ? '' : `<td>${esc(it.klasse)}</td>`}
            ${hide === 'variante' ? '' : `<td>${esc(it.variante)}</td>`}
            <td>${esc(it.lfs)}${
                it.lehrerLFS ? '<div class="srdp-table__sub">' + esc(it.lehrerLFS) + '</div>' : ''
            }</td>
            ${
                hide === 'wahlfach'
                    ? ''
                    : `<td>${esc(it.wahlfachMuendlich)}${
                          it.seminar ? '<div class="srdp-table__sub">Seminar: ' + esc(it.seminar) + '</div>' : ''
                      }</td>`
            }
            ${o.showRoles ? `<td>${roles || '–'}</td>` : ''}
            <td>${esc(it.betreuungslehrerDA || it.betreuungslehrerABA || '–')}</td>
            ${showNach ? `<td>${renderNacharbeitStatus(it)}</td>` : ''}
            <td class="srdp-table__chevron"><i class="bi bi-chevron-right" aria-hidden="true"></i></td>
          </tr>`;
            })
            .join('')}
      </tbody>
    </table></div>`;
}

function renderTable(items, state, opts) {
    const o = opts || {};
    const list = Array.isArray(items) ? items : [];
    const by = String(o.groupBy != null ? o.groupBy : state.groupBy || '');
    if (by !== 'klasse' && by !== 'variante' && by !== 'wahlfach') {
        return renderFlatTable(list, state, o);
    }
    if (!list.length) {
        return quietEmpty('<strong>Keine Anmeldungen</strong> · In dieser Ansicht ist nichts vorhanden.');
    }
    return groupItemsBy(list, by)
        .map(
            (g) => `<section class="srdp-group">
      <header class="srdp-group__head"><h3>${esc(g.key)}</h3><span>${g.items.length}</span></header>
      ${renderFlatTable(g.items, state, { ...o, hideCol: by })}
    </section>`
        )
        .join('');
}

function renderBars(buckets, total) {
    const max = Math.max(1, total || 1);
    return `<div class="srdp-bars">${(buckets || [])
        .map((b) => {
            const pct = Math.round((b.count / max) * 100);
            return `<div class="srdp-bar">
          <div>
            <div>${esc(b.key)} · ${b.count}</div>
            <div class="srdp-bar__track"><div class="srdp-bar__fill" style="width:${pct}%"></div></div>
          </div>
          <strong>${b.count}</strong>
        </div>`;
        })
        .join('')}</div>`;
}

function renderLehrerDashboard(state) {
    const mine = itemsForCurrentRole(state);
    const kv = kvKlasseCodesForState(state);
    const headline =
        mine.length === 1
            ? '<span class="fb-num">1</span> Prüfung ist dir zugeordnet.'
            : `<span class="fb-num">${mine.length}</span> Prüfungen sind dir zugeordnet.`;
    const sub = mine.length
        ? 'Zuordnungen als Betreuung, LFS, BKO oder Wahlfach.'
        : 'Aktuell ist keine Prüfung bei dir offen.';
    return `
    <h1 class="fb-greet">${headline}</h1>
    <p class="fb-sub">${esc(sub)}</p>
    ${
        mine.length
            ? `<div class="fb-sec-label">MEINE ZUORDNUNGEN</div>${renderTable(mine.slice(0, 6), state, {
                  showRoles: true,
                  groupBy: ''
              })}`
            : quietEmpty('<strong>Alles erledigt</strong> · Keine offenen Zuordnungen.')
    }
    <div class="fb-actions-row">
      <button type="button" class="fb-btn" data-srdp-view-jump="meine-pruefungen">Meine Prüfungen</button>
      <button type="button" class="fb-btn fb-btn--ghost" data-srdp-view-jump="liste">Alle Anmeldungen</button>
      ${
          kv.length
              ? '<button type="button" class="fb-btn fb-btn--ghost" data-srdp-view-jump="meine-klasse">Meine Klasse</button>'
              : ''
      }
    </div>`;
}

function renderLeitungDashboard(state) {
    const items = state.items || [];
    const o = state.options || {};
    const fehlend = klassenOhneAnmeldung(items, o.klassen || []);
    const nacharbeit = countNacharbeitPending(items);
    const stats = computeAnmeldungStats(items);
    const headline =
        fehlend.length === 1
            ? '<span class="fb-num">1</span> Klasse ohne Anmeldung.'
            : `<span class="fb-num">${fehlend.length}</span> Klassen ohne Anmeldung.`;
    const sub = fehlend.length
        ? `${fehlend.slice(0, 4).join(', ')}${fehlend.length > 4 ? ' …' : ''}`
        : nacharbeit
          ? `${nacharbeit} bestätigte Anmeldung${nacharbeit === 1 ? '' : 'en'} warten auf die Prüfungsliste.`
          : 'Aktuell wartet keine Klasse ohne Anmeldung.';

    return `
    <h1 class="fb-greet">${headline}</h1>
    <p class="fb-sub">${esc(sub)}</p>
    <div class="fb-sec-label">STAND ${esc(o.jahr || '')}</div>
    <p class="muted" style="margin:6px 0 0;">${stats.total} Anmeldungen · ${stats.byKlasse.length} Klassen · ${
        stats.byVariante.length
    } Varianten · ${nacharbeit} Nacharbeit offen.</p>
    <div class="fb-actions-row">
      ${
          fehlend.length
              ? `<button type="button" class="fb-btn" data-srdp-jump-liste data-srdp-jump-klasse="${esc(
                    fehlend[0] || ''
                )}">Liste öffnen</button>`
              : '<button type="button" class="fb-btn" data-srdp-view-jump="liste">Anmeldungen</button>'
      }
      <button type="button" class="fb-btn fb-btn--ghost" data-srdp-view-jump="auswertung">Auswertung</button>
    </div>`;
}

function renderDashboard(state) {
    if (isSrdpStaffRole(state.role)) return renderLeitungDashboard(state);
    return renderLehrerDashboard(state);
}

function renderMeine(state) {
    const mine = itemsForCurrentRole(state);
    if (!mine.length) {
        return quietEmpty(
            '<strong>Noch keine Anmeldung</strong> · Du kannst dich jetzt für den Haupttermin anmelden.',
            `<div class="srdp-actions"><button type="button" class="btn btn-success" data-srdp-view-jump="anmelden"><i class="bi bi-pencil-square"></i>Jetzt anmelden</button></div>`
        );
    }
    return mine
        .map((it) =>
            renderSummaryCard(it, state.options.profileId, { jahr: state.options.jahr })
        )
        .join('');
}

function renderAuswertung(state) {
    const items = filterItemsForView(state.items || [], state);
    const stats = computeAnmeldungStats(items);
    const by =
        state.groupBy === 'variante' || state.groupBy === 'wahlfach' || state.groupBy === 'klasse'
            ? state.groupBy
            : '';
    const groups = by ? groupItemsBy(items, by) : [{ key: 'Alle', items }];
    const cards = groups.length
        ? `<div class="srdp-eval">${groups
              .map((g) => {
                  const rows = g.items
                      .map((it) => {
                          const name = it.title || [it.nachname, it.vorname].filter(Boolean).join(', ');
                          const meta = [
                              by === 'klasse' ? '' : it.klasse,
                              by === 'variante' ? '' : it.variante,
                              by === 'wahlfach' ? '' : it.wahlfachMuendlich
                          ]
                              .map((x) => String(x || '').trim())
                              .filter(Boolean)
                              .join(' · ');
                          return `<li><button type="button" data-srdp-open-detail="${esc(it.itemId)}">
                <strong>${esc(name || 'Anmeldung')}</strong>
                <span>${esc(meta || '–')}</span>
              </button></li>`;
                      })
                      .join('');
                  return `<article class="srdp-eval-card">
            <header><h3>${esc(g.key)}</h3><strong>${g.items.length}</strong></header>
            <ul>${rows}</ul>
          </article>`;
              })
              .join('')}</div>`
        : quietEmpty('<strong>Keine Anmeldungen</strong> · In dieser Auswahl ist nichts vorhanden.');
    return `${renderGroupBar(state)}
    <div class="srdp-kpis">
      <div class="srdp-kpi"><strong>${stats.total}</strong><span>Anmeldungen</span></div>
      <div class="srdp-kpi"><strong>${stats.byKlasse.length}</strong><span>Klassen</span></div>
      <div class="srdp-kpi"><strong>${stats.mitSeminar}</strong><span>mit Seminar</span></div>
    </div>
    ${cards}
    <section class="srdp-panel srdp-panel--calm"><h3>Verteilung</h3>
      <div class="srdp-eval-dist">
        <div><h4>Klassen</h4>${renderBars(stats.byKlasse, stats.total)}</div>
        <div><h4>Varianten</h4>${renderBars(stats.byVariante, stats.total)}</div>
        <div><h4>Wahlfächer</h4>${renderBars(stats.byWahlfach, stats.total)}</div>
      </div>
    </section>`;
}

function formatSyncPullAt(iso) {
    const s = String(iso || '').trim();
    if (!s) return '';
    const d = new Date(s);
    if (Number.isNaN(d.getTime())) return s;
    try {
        return d.toLocaleString('de-AT', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch {
        return s;
    }
}

function renderSystem(state) {
    const o = state.options || {};
    const live = state.stammdatenLive || {};
    const sync = state.syncMeta || {};
    const syncSource = state.localDemoOnly
        ? 'demo'
        : sync.source || (sync.connected ? 'sharepoint' : 'none');
    const withItemId = (state.items || []).filter((it) => String(it.itemId || '').trim()).length;
    const connected = !!(sync.connected || (state.siteUrl && !state.error && !state.localDemoOnly));
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
    const liveOk = !!(live.ok || (Number(live.classCount) || 0) > 0);
    const liveLine = liveOk
        ? `SharePoint: ${Number(live.classCount) || 0} Klassen, ${Number(live.teacherCount) || 0} Lehrkräfte`
        : live.error
          ? `Live-Laden: ${live.error}`
          : `Quelle: ${live.source || '–'} · Klassen ${Number(live.classCount) || 0}`;
    const syncLine = `Sync: ${syncSource}${
        sync.listTitle ? ' · Liste „' + sync.listTitle + '“' : ''
    }${sync.lastPullAt ? ' · zuletzt ' + formatSyncPullAt(sync.lastPullAt) : ''}${
        state.localDemoOnly ? ' · Demo local-first' : ''
    } · ${withItemId}/${(state.items || []).length} mit itemId`;
    const perms = loadEffectiveSrdpPermissionsConfig();
    const adminSum = formatSrdpAudienceSummary({
        groups: perms.adminGroups,
        users: perms.adminUsers,
        id: perms.groupAdminId,
        label: perms.groupAdmin
    });
    const lehrerSum = formatSrdpAudienceSummary({
        groups: perms.lehrerGroups,
        users: perms.lehrerUsers,
        id: perms.groupLehrerId,
        label: perms.groupLehrer
    });
    const kandSum = formatSrdpAudienceSummary({
        groups: perms.kandidatenGroups,
        users: perms.kandidatenUsers,
        id: perms.groupKandidatenId,
        label: perms.groupKandidaten
    });
    const permsOk = adminSum !== '— nicht gesetzt —' || lehrerSum !== '— nicht gesetzt —';
    const tab = normalizeSystemTab(state.systemTab);
    const noSite = !String(state.siteUrl || '').trim();
    const panelOpen = (id) => systemPanelHiddenAttr(tab, id);
    return `
    <div class="ms-sys">
      ${renderSystemTabBar({
          prefix: 'srdp',
          activeTab: tab,
          dataAttr: 'data-srdp-sys-tab',
          ariaLabel: 'Systemadministration sRDP'
      })}

      <section class="ms-sys-card" id="srdpSysPanel-site" role="tabpanel" aria-labelledby="srdpSysTab-site"${panelOpen(
          'site'
      )}>
        <header class="ms-sys-card__head">
          <span class="ms-sys-card__step">1</span>
          <div>
            <h3>Datenquelle</h3>
            <p>SharePoint-Liste für den Haupttermin verbinden – Lehrer sehen diesen Bereich nicht.</p>
          </div>
        </header>
        ${
            noSite
                ? `<div class="ms-sys-inline-warn" role="status">
          <strong>Keine SharePoint-URL in diesem Browser.</strong>
          <p class="muted" style="margin:6px 0 0">Site hier eintragen oder im vollständigen Setup hinterlegen.</p>
        </div>`
                : ''
        }
        <div class="ms-sys-health" role="status">
          <span class="ms-sys-status ms-sys-status--${statusMod}"><span class="ms-sys-status__dot" aria-hidden="true"></span>${esc(
              statusLabel
          )}</span>
          <span class="ms-sys-status ms-sys-status--${liveOk ? 'ok' : 'warn'}"><span class="ms-sys-status__dot" aria-hidden="true"></span>${esc(
              liveLine
          )}</span>
          <span class="ms-sys-status ms-sys-status--${
              sync.connected && !state.localDemoOnly ? 'ok' : 'warn'
          }"><span class="ms-sys-status__dot" aria-hidden="true"></span>${esc(syncLine)}</span>
        </div>
        <div class="ms-sys-field">
          <label for="srdpSiteUrl">SharePoint-URL</label>
          <div class="ms-sys-field__row">
            <input type="url" id="srdpSiteUrl" class="ms-sys-input" value="${esc(
                state.siteUrl || ''
            )}" placeholder="https://…sharepoint.com/sites/…" spellcheck="false" autocomplete="off">
            <button type="button" class="btn ms-sys-btn-primary" data-srdp-save-site><i class="bi bi-plug"></i>Verbindung prüfen</button>
          </div>
        </div>
        <div class="ms-sys-field__grid">
          <div class="ms-sys-field">
            <label>Profil</label>
            <input type="text" class="ms-sys-input" value="${esc(o.profileLabel || '')}" readonly>
          </div>
          <div class="ms-sys-field">
            <label>Terminjahr</label>
            <input type="text" class="ms-sys-input" value="${esc(o.jahr || '')}" readonly>
          </div>
        </div>
        <div class="ms-sys-field">
          <label>Listenname</label>
          <input type="text" class="ms-sys-input" value="${esc(
              o.listTitle || sync.listTitle || ''
          )}" readonly>
        </div>
        <div class="ms-sys-actions">
          <button type="button" class="btn ms-sys-btn-secondary" data-srdp-reload><i class="bi bi-arrow-repeat"></i>Von SharePoint laden</button>
          <a class="btn ms-sys-btn-secondary" href="sharepoint-liste-srdp.html"><i class="bi bi-list-ul"></i>Listen verwalten</a>
          <a class="btn ms-sys-btn-ghost" href="srdp-mail-setup.html"><i class="bi bi-envelope-check"></i>Bestätigungsmail-Flow</a>
        </div>
      </section>

      <section class="ms-sys-card" id="srdpSysPanel-perms" role="tabpanel" aria-labelledby="srdpSysTab-perms"${panelOpen(
          'perms'
      )}>
        <header class="ms-sys-card__head">
          <span class="ms-sys-card__step">2</span>
          <div>
            <h3>Berechtigungen</h3>
            <p>Vollzugriff, Lesende Rechte und Kandidat:innen – Detailpflege im Listen-Assistenten.</p>
          </div>
        </header>
        <div class="ms-sys-callout" role="note">
          <strong>Übersicht</strong> aus der gespeicherten Konfiguration.
          Gruppen und Listenrechte setzt du im vollständigen Setup.
        </div>
        <div class="ms-sys-health" role="status">
          <span class="ms-sys-status ms-sys-status--${permsOk ? 'ok' : 'warn'}"><span class="ms-sys-status__dot" aria-hidden="true"></span>${
              permsOk ? 'Rollen teilweise gesetzt' : 'Noch keine Gruppen gesetzt'
          }</span>
        </div>
        <dl class="ms-sys-dl">
          <dt>Vollzugriff</dt><dd>${esc(adminSum)}</dd>
          <dt>Lesende Rechte</dt><dd>${esc(lehrerSum)}</dd>
          <dt>Kandidat:innen</dt><dd>${esc(kandSum)}</dd>
        </dl>
        <div class="ms-sys-links">
          <a class="btn ms-sys-btn-secondary" href="sharepoint-liste-srdp.html"><i class="bi bi-shield-check"></i>Vollständiges Setup · Rechte</a>
        </div>
      </section>

      <section class="ms-sys-card" id="srdpSysPanel-kalender" role="tabpanel" aria-labelledby="srdpSysTab-kalender"${panelOpen(
          'kalender'
      )}>
        <header class="ms-sys-card__head">
          <span class="ms-sys-card__step">3</span>
          <div>
            <h3>Kalender &amp; Fristen</h3>
            <p>Anmeldefrist und Haupttermin steuern Phase und Kalenderansicht.</p>
          </div>
        </header>
        <div class="ms-sys-field__grid">
          <div class="ms-sys-field">
            <label for="srdpAnmeldeFrist">Anmeldefrist</label>
            <input type="date" id="srdpAnmeldeFrist" class="ms-sys-input" value="${esc(
                String(o.anmeldeFrist || '').slice(0, 10)
            )}">
          </div>
          <div class="ms-sys-field">
            <label for="srdpHaupttermin">Haupttermin</label>
            <input type="date" id="srdpHaupttermin" class="ms-sys-input" value="${esc(
                String(o.haupttermin || '').slice(0, 10)
            )}">
          </div>
        </div>
        <div class="ms-sys-actions">
          <button type="button" class="btn ms-sys-btn-primary" data-srdp-save-termine><i class="bi bi-calendar-check"></i>Frist &amp; Termin speichern</button>
          <button type="button" class="btn ms-sys-btn-secondary" data-srdp-view-jump="kalender"><i class="bi bi-calendar3"></i>Kalender öffnen</button>
        </div>
      </section>

      <section class="ms-sys-card" id="srdpSysPanel-wartung" role="tabpanel" aria-labelledby="srdpSysTab-wartung"${panelOpen(
          'wartung'
      )}>
        <header class="ms-sys-card__head">
          <span class="ms-sys-card__step">4</span>
          <div>
            <h3>Wartung &amp; Diagnose</h3>
            <p>Stammdaten-Diagnose, Demo und Setup – getrennt vom Schulalltag.</p>
          </div>
        </header>
        <div class="ms-sys-split">
          <div class="ms-sys-split__block">
            <h4>Stammdaten</h4>
            <p class="muted">Quelle: ${esc(live.source || '–')} · Klassen ${
                Number(live.classCount) || 0
            } · Lehrkräfte ${Number(live.teacherCount) || 0}</p>
            <dl class="ms-sys-dl">
              <dt>Klassen</dt><dd>${esc((o.klassen || []).join(', ') || '–')}</dd>
              <dt>LFS</dt><dd>${esc((o.lfs || []).join(', ') || '–')}</dd>
              <dt>Wahlfächer</dt><dd>${esc(
                  (o.wahlfaecher || []).slice(0, 8).join(', ') || '–'
              )}${(o.wahlfaecher || []).length > 8 ? ' …' : ''}</dd>
            </dl>
            <div class="ms-sys-actions">
              <a class="btn ms-sys-btn-secondary" href="../tenant.html"><i class="bi bi-gear"></i>Stammdaten</a>
              <a class="btn ms-sys-btn-ghost" href="sharepoint-liste-srdp.html"><i class="bi bi-sliders"></i>Vollständiges Setup</a>
            </div>
          </div>
          <div class="ms-sys-split__block">
            <h4>Demo &amp; Hilfe</h4>
            <p class="muted">Vorführung ohne SharePoint – Seeds bleiben lokal.</p>
            <div class="ms-sys-actions">
              <button type="button" class="btn ms-sys-btn-secondary" data-srdp-demo><i class="bi bi-database"></i>Demo lokal</button>
              <a class="btn ms-sys-btn-ghost" href="../hilfe.html#tool-srdp-anmeldung"><i class="bi bi-life-preserver"></i>Hilfe</a>
            </div>
          </div>
        </div>
        <div class="ms-sys-split" style="margin-top:1rem">
          <div class="ms-sys-split__block ms-sys-split__block--wide">
            <h4>Abschlussarbeiten (ABA-Portal)</h4>
            <p class="muted">Excel-Export vom ABA-Portal hochladen – Titel und Betreuer werden bei der Anmeldung vorausgefüllt. Pro Schule und Terminjahr.</p>
            <p class="muted">Aktuell: ${Number((state.abschlussarbeitenMeta && state.abschlussarbeitenMeta.count) || (state.abschlussarbeiten || []).length) || 0} Einträge${
                state.abschlussarbeitenMeta && state.abschlussarbeitenMeta.listTitle
                    ? ' · Liste „' + esc(state.abschlussarbeitenMeta.listTitle) + '“'
                    : ''
            }${
                state.abschlussarbeitenMeta && state.abschlussarbeitenMeta.lastImportAt
                    ? ' · Import ' + esc(String(state.abschlussarbeitenMeta.lastImportAt).slice(0, 19))
                    : ''
            }</p>
            <div class="ms-sys-actions">
              <input type="file" id="srdpAbaFile" accept=".xlsx,.xls,.csv" hidden>
              <button type="button" class="btn ms-sys-btn-primary" data-srdp-aba-pick-file><i class="bi bi-upload"></i>ABA-Excel wählen</button>
              <button type="button" class="btn ms-sys-btn-secondary" data-srdp-aba-reload${
                  state.localDemoOnly ? ' disabled' : ''
              }><i class="bi bi-arrow-repeat"></i>Aus SharePoint laden</button>
            </div>
            <p class="srdp-field__hint">Erwartete Spalten: Titel, Schüler/in, Klasse, Betreuer/in, Status. Team-Zeilen werden pro Person aufgeteilt. Bestehende Import-Daten des Terminjahrs werden ersetzt. Muster:
              <a href="../assets/srdp/ABA-Portal-Demo-Diplomarbeiten.xlsx" download>ABA-Portal-Demo-Diplomarbeiten.xlsx</a>
              · IT-Setup: <a href="sharepoint-liste-srdp.html">Listen anlegen &amp; ABA hochladen</a>.</p>
          </div>
        </div>
      </section>

      <section class="ms-sys-card ms-sys-card--danger" id="srdpSysPanel-danger" role="tabpanel" aria-labelledby="srdpSysTab-danger"${panelOpen(
          'danger'
      )}>
        <header class="ms-sys-card__head">
          <span class="ms-sys-card__step ms-sys-card__step--danger" aria-hidden="true"><i class="bi bi-exclamation-triangle"></i></span>
          <div>
            <h3>Danger Zone</h3>
            <p>Disruptive Aktionen – nur bewusst ausführen.</p>
          </div>
        </header>
        <button type="button" class="btn ms-sys-btn-danger" data-srdp-demo-reset><i class="bi bi-trash"></i>Demo zurücksetzen</button>
      </section>
    </div>`;
}

function examItemsForCalendar(state) {
    const items = Array.isArray(state.items) ? state.items : [];
    if (state.role === 'lehrer') return filterItemsForTeacher(items, state.teacherLabel);
    if (state.role === 'schueler') return itemsForCurrentRole(state);
    return items;
}

function teacherOwnsWrittenPart(part, roles) {
    const p = String(part || '').toLowerCase();
    return (roles || []).some((role) => {
        if (role.role === 'lfs' && p.includes('lfs')) return true;
        if (role.role === 'vierte' && (p.includes('klausur') || p.includes('englisch'))) return true;
        return false;
    });
}

function buildCalendarEvents(state) {
    const o = state.options || {};
    const jahr = o.jahr || '';
    /** @type {object[]} */
    const events = [];
    if (o.anmeldeFrist) {
        events.push({
            id: 'frist',
            title: 'Anmeldefrist sRDP',
            start: o.anmeldeFrist,
            allDay: true,
            tone: 'warn',
            subtitle: 'Haupttermin ' + jahr
        });
    }
    const items = examItemsForCalendar(state);
    if (state.role === 'admin') {
        /** @type {Map<string, { iso: string, fach: string, names: string[] }>} */
        const buckets = new Map();
        items.forEach((it) => {
            const plan = buildPruefplanDetail(it, o.profileId, {
                terminJahr: jahr,
                bfkDatum: o.bfkDatum || ''
            });
            plan.schriftlich.forEach((row) => {
                if (!row.datumIso || !row.fach) return;
                const key = row.datumIso + '|' + row.fachKey;
                if (!buckets.has(key)) buckets.set(key, { iso: row.datumIso, fach: row.fach, names: [] });
                const name = it.title || [it.nachname, it.vorname].filter(Boolean).join(' ');
                if (name) buckets.get(key).names.push(name);
            });
        });
        buckets.forEach((bucket, key) => {
            const n = bucket.names.length;
            events.push({
                id: 'exam-' + key,
                title: bucket.fach + ' schriftlich',
                start: bucket.iso,
                allDay: true,
                tone: 'info',
                subtitle:
                    'Haupttermin ' +
                    jahr +
                    ' · ' +
                    n +
                    (n === 1 ? ' Kandidat:in' : ' Kandidat:innen')
            });
        });
        return events;
    }
    items.forEach((it) => {
        const plan = buildPruefplanDetail(it, o.profileId, { terminJahr: jahr, bfkDatum: o.bfkDatum || '' });
        const roles = state.role === 'lehrer' ? teacherRolesOnItem(it, state.teacherLabel) : null;
        plan.schriftlich.forEach((row) => {
            if (!row.datumIso || !row.fach) return;
            if (roles && row.fachKey !== 'bfk' && !teacherOwnsWrittenPart(row.part, roles)) return;
            const who = it.title || [it.nachname, it.vorname].filter(Boolean).join(' ');
            events.push({
                id: String(it.itemId || who) + '-' + row.fachKey,
                title: row.fach + ' schriftlich',
                start: row.datumIso,
                allDay: true,
                tone: 'info',
                subtitle: [it.klasse, who, 'Haupttermin ' + jahr].filter(Boolean).join(' · '),
                openAttr: it.itemId ? 'data-srdp-open-detail="' + esc(it.itemId) + '"' : ''
            });
        });
    });
    return events;
}

export function firstCalendarExamDate(state) {
    const exams = buildCalendarEvents(state)
        .filter((ev) => ev.id !== 'frist')
        .map((ev) => String(ev.start || ''))
        .filter((iso) => /^\d{4}-\d{2}-\d{2}$/.test(iso))
        .sort();
    return exams[0] || '';
}

function renderTermine(state) {
    const jahr = (state.options && state.options.jahr) || '';
    const bundle = periodenForTerminJahr(jahr);
    const bfkMatch = String((state.options && state.options.bfkDatum) || '').match(/20\d{2}-\d{2}-\d{2}/);
    const bfkIso = bfkMatch ? bfkMatch[0] : '';
    const blocks = bundle.perioden
        .map((periode) => {
            const rows = (periode.block.faecher || []).map((fach) => {
                const daten = (fach.daten || []).slice().filter(Boolean).sort();
                return {
                    sort: daten[0] || '9999-99-99',
                    label: fach.fach,
                    when: daten.map((iso) => formatGermanDate(iso)).filter(Boolean).join(' und ') || '–',
                    school: false
                };
            });
            if (periode.id === 'haupt') {
                rows.push({
                    sort: bfkIso || '9999-99-99',
                    label: 'BFK schriftlich',
                    when: bfkIso ? formatGermanDate(bfkIso) : 'noch nicht festgelegt',
                    school: true
                });
            }
            rows.sort((a, b) => a.sort.localeCompare(b.sort) || (a.school ? 1 : 0) - (b.school ? 1 : 0));
            const body = rows
                .map(
                    (row) =>
                        `<tr><th scope="row">${esc(row.label)}</th><td>${esc(row.when)}</td></tr>`
                )
                .join('');
            const neben = periode.neben ? `<p class="srdp-termine__neben">${esc(periode.neben)}</p>` : '';
            return `<section class="srdp-termine__block">
        <header>
          <h3>${esc(periode.label)} ${esc(periode.block.jahr)}</h3>
          ${neben}
        </header>
        <table class="srdp-termine__table">
          <tbody>${body}</tbody>
        </table>
      </section>`;
        })
        .join('');
    const empty = bundle.perioden.length
        ? ''
        : state.termineSource === 'live' || state.termineSource === 'cache'
          ? `<p class="srdp-panel__hint">Für ${esc(jahr || 'dieses Jahr')} stehen auf matura.gv.at noch keine Termine.</p>`
          : `<p class="srdp-panel__hint">Für ${esc(jahr || 'dieses Jahr')} liegt noch kein Termin vor. Die Seite matura.gv.at wird beim Öffnen gelesen, sobald das Jahr dort veröffentlicht ist.</p>`;
    return `<section class="srdp-panel srdp-panel--calm srdp-termine">
      <h2 class="srdp-sr-only">sRDP-Termine</h2>
      <p class="srdp-panel__hint">Schuljahr ${esc(bundle.schuljahr || '–')} · Die Anmeldung gilt für den Haupttermin. Die schriftliche BFK steht dort mit den zentralen Klausuren. Herbst- und Wintertermin sind die Nebentermine.</p>
      ${empty}
      ${blocks}
      <p class="srdp-termine__source">${
          state.termineSource === 'live' || state.termineSource === 'cache'
              ? 'Gerade von matura.gv.at gelesen. '
              : 'Bis die Seite gelesen ist, gilt der letzte mitgelieferte Stand. '
      }Quelle: <a href="${esc(MATURA_TERMIN_URL)}" target="_blank" rel="noopener noreferrer">matura.gv.at/pruefungstermine</a></p>
    </section>`;
}

function renderKalender(state) {
    const mode = state.calMode || 'month';
    const y = Number(state.calYear) || new Date().getFullYear();
    const m = Number(state.calMonth) || new Date().getMonth() + 1;
    const weekAnchor = state.calWeekAnchor || toIsoDateOnly(new Date());
    const shell = renderMs365CalendarShell({
        mode,
        events: buildCalendarEvents(state),
        weekAnchor,
        year: y,
        month0: m - 1,
        prefix: 'srdp-cal',
        hint: 'Schriftliche Klausuren am Haupttermin. BFK, BKO und Wahlfach legt die Schule fest. Alle Fächer unter sRDP-Termine.',
        actionsHtml: ''
    });
    return `<section class="srdp-panel srdp-panel--calm srdp-cal-layout">
      <h2 class="srdp-sr-only">Kalender</h2>
      ${renderFilters(state)}
      ${shell}
    </section>`;
}

function renderView(state) {
    if (state.submitSuccess) {
        return renderSubmitSuccess(state.submitSuccess, state, {
            heading: 'Deine Anmeldung',
            sub: 'Bitte prüfe die Angaben. Änderungen später über das Sekretariat.'
        });
    }
    const view = state.view === 'einrichtung' ? 'system' : state.view;
    switch (view) {
        case 'anmelden': {
            const existing = findAnmeldungBySchuelerEmail(
                state.items,
                effectiveAccountEmail(state)
            );
            if (existing) {
                return (
                    quietEmpty(
                        '<strong>Du bist bereits angemeldet</strong> · Eine zweite Anmeldung ist nicht nötig.',
                        `<div class="srdp-actions"><button type="button" class="btn btn-success" data-srdp-view-jump="meine">Zur Zusammenfassung →</button></div>`
                    ) +
                    renderSummaryCard(existing, state.options.profileId, { jahr: state.options.jahr })
                );
            }
            if (!isSrdpAnmeldeFensterOffen(state.options)) {
                const frist = String((state.options && state.options.anmeldeFrist) || '').slice(
                    0,
                    10
                );
                return `<section class="srdp-calm-banner srdp-calm-banner--ops" role="status">
          <div class="srdp-calm-banner__text">
            <p class="srdp-calm-banner__eyebrow">Anmeldephase</p>
            <p><strong>Das Anmeldefenster ist geschlossen.</strong></p>
            <p class="muted">${
                frist
                    ? 'Frist war ' + esc(frist) + '. Bitte das Sekretariat fragen, falls noch etwas fehlt.'
                    : 'Bitte später erneut versuchen oder das Sekretariat fragen.'
            }</p>
          </div>
        </section>`;
            }
            return renderAntragWizard(state);
        }
        case 'meine':
            return renderMeine(state);
        case 'meine-pruefungen':
            return renderTable(itemsForCurrentRole(state), state, { showRoles: true, groupBy: '' });
        case 'meine-klasse': {
            const kvCodes = kvKlasseCodesForState(state);
            const rows = itemsForCurrentRole(state);
            if (!rows.length) {
                return (
                    `<p class="srdp-panel__hint">KV-Klassen: ${esc(kvCodes.join(', ') || '–')}</p>` +
                    quietEmpty(
                        `<strong>Noch keine Anmeldungen</strong> · In ${esc(
                            kvCodes.join(', ') || 'deinen Klassen'
                        )} ist noch nichts eingegangen.`,
                        `<div class="srdp-actions"><button type="button" class="btn" data-srdp-view-jump="liste">Zur Gesamtliste</button></div>`
                    )
                );
            }
            return (
                `<p class="srdp-panel__hint">KV-Klassen: ${esc(kvCodes.join(', ') || '–')}</p>` +
                renderTable(rows, state, { showRoles: false, groupBy: '' })
            );
        }
        case 'liste':
            return (
                renderGroupBar(state) +
                renderTable(filterItemsForView(state.items || [], state), state, {
                    showRoles: state.role === 'lehrer'
                })
            );
        case 'termine':
            return renderTermine(state);
        case 'kalender':
            return renderKalender(state);
        case 'auswertung':
            return renderAuswertung(state);
        case 'system':
            return renderSystem(state);
        case 'dashboard':
        default:
            return renderDashboard(state);
    }
}

export function syncSrdpPageChrome(state) {
    const staff = isSrdpStaffRole(state.role) && !state.demoPersonaEmail;
    const itOnly = document.querySelectorAll('[data-ms365-frontend-it-only]');
    itOnly.forEach((el) => {
        el.hidden = !staff;
    });
}

export function renderApp(state, root) {
    if (!root) return;
    syncSrdpPageChrome(state);
    const kv = kvKlasseCodesForState(state);
    state.kvKlassen = kv.length ? kv : state.kvKlassen || [];
    if (state.view === 'einrichtung') state.view = 'system';
    const sections = navSectionsForRole(state.role, state.planerRoles, { kvKlassen: state.kvKlassen });
    const roleBlock = renderRoleBlock(state);
    const staff = isSrdpStaffRole(state.role);
    const teacherLike = !staff;
    const nacharbeit = countNacharbeitPending(state.items);
    const detail = findItemById(state, state.detailItemId);
    const navHtml = sections
        .map((sec) => {
            const label = sec.label
                ? `<span class="srdp-nav__section-label">${esc(sec.label)}</span>`
                : '';
            const btns = (sec.items || [])
                .map((v) => {
                    const badge =
                        staff && v.id === 'liste' && nacharbeit
                            ? `<span class="srdp-nav__badge">${nacharbeit}</span>`
                            : '';
                    return `
            <button type="button" class="srdp-nav__btn${state.view === v.id ? ' is-active' : ''}${
                v.id === 'anmelden' ? ' srdp-nav__btn--cta' : ''
            }" data-srdp-view="${esc(v.id)}">
              <i class="bi ${v.icon}" aria-hidden="true"></i><span class="srdp-nav__btn-label">${esc(
                        v.label
                    )}</span>${badge}
            </button>`;
                })
                .join('');
            return `<div class="srdp-nav__section">${label}${btns}</div>`;
        })
        .join('');

    const showHardError =
        state.error &&
        !(state.localDemoOnly && /403|Zugriff verweigert|Forbidden/i.test(String(state.error)));

    root.innerHTML = `
    <div class="srdp-shell${teacherLike ? ' srdp-shell--lehrer' : ''}${staff ? ' srdp-shell--staff' : ''}">
      <aside class="srdp-nav" aria-label="Navigation">
        <div class="srdp-nav__brand">
          <i class="bi bi-mortarboard-fill" aria-hidden="true"></i>
          <div>
            <strong>sRDP-Anmeldung</strong>
            <span>${
                state.role === 'schueler'
                    ? 'Matura anmelden'
                    : state.role === 'lehrer'
                      ? 'Prüfungen'
                      : 'Verwaltung'
            }</span>
          </div>
        </div>
        <nav class="srdp-nav__list">${navHtml}</nav>
        ${roleBlock ? `<div class="srdp-nav__role">${roleBlock}</div>` : ''}
      </aside>
      <div class="srdp-main">
        ${renderTop(state)}
        ${
            state.demoPersonaEmail
                ? `<div class="srdp-alert srdp-alert--info" role="status">IT-Vorschau als <strong>${esc(
                      state.demoPersonaName || 'Lehrkraft'
                  )}</strong> (${esc(state.demoPersonaEmail)}). Angemeldet bleibt ${esc(
                      state.accountEmail || 'IT'
                  )}.</div>`
                : ''
        }
        ${
            state.localDemoOnly && state.view === 'system'
                ? '<div class="srdp-alert srdp-alert--info" role="status">Lokaler Demo-Modus – Anzeige ohne SharePoint (Freitag-Workshop).</div>'
                : ''
        }
        ${showHardError ? `<div class="srdp-alert srdp-alert--bad" role="alert">${esc(state.error)}</div>` : ''}
        ${state.softError ? `<div class="srdp-alert srdp-alert--info" role="status">${esc(state.softError)}</div>` : ''}
        ${state.loading ? '<div class="srdp-alert srdp-alert--info">Lade Anmeldungen…</div>' : ''}
        <div id="srdpView" class="srdp-view">
          ${state.loading ? '' : renderView(state)}
        </div>
      </div>
    </div>`;

    let host = document.getElementById('srdpDrawerHost');
    if (!host) {
        host = document.createElement('div');
        host.id = 'srdpDrawerHost';
        document.body.appendChild(host);
    }
    host.innerHTML = detail ? renderAnmeldungDrawer(detail, state) : '';
}

/**
 * @param {object} [prev] bisheriger Formularzustand (Hidden-Felder / andere Wizard-Schritte)
 */
function readTeacherMulti(fieldName, fallback) {
    const boxes = document.querySelectorAll(
        'input[type="checkbox"][data-srdp-teacher-field="' + fieldName + '"]'
    );
    if (boxes && boxes.length) {
        const picked = [];
        boxes.forEach((el) => {
            if (el.checked) picked.push(String(el.value || '').trim());
        });
        return picked.filter(Boolean).join('; ');
    }
    const hidden = document.querySelector('input[name="' + fieldName + '"]');
    if (hidden && hidden.value != null && String(hidden.value).trim()) {
        return String(hidden.value).trim();
    }
    return fallback;
}

export function readFormFromDom(prev) {
    const g = (id) => document.getElementById(id);
    const p = prev && typeof prev === 'object' ? prev : {};
    const val = (id, fallback) => {
        const el = g(id);
        if (!el) return fallback;
        return el.value;
    };
    const best = g('srdpBestaetigung');
    const bestaetigung = best
        ? best.type === 'checkbox'
            ? !!best.checked
            : String(best.value || '') === '1' || String(best.value || '').toLowerCase() === 'true'
        : !!p.bestaetigung;
    const formEl = document.getElementById('srdpForm');
    const step = formEl
        ? Number(formEl.getAttribute('data-srdp-wizard-step') || p.wizardStep || 1)
        : Number(p.wizardStep) || 1;
    const radio = document.querySelector('input[type="radio"][name="variante"]:checked');
    const variante = radio
        ? String(radio.value || '').trim()
        : val('srdpVariante', p.variante);
    const hatAbaEl = g('srdpHatAba');
    const betreuung = val('srdpBetreuung', '');
    const daTitle = val('srdpDA', '');
    const isAbaBetreuung =
        !!(g('srdpBetreuung') && g('srdpBetreuung').getAttribute('name') === 'betreuungslehrerABA');
    return emptyForm({
        klasse: val('srdpKlasse', p.klasse),
        nachname: val('srdpNachname', p.nachname),
        vorname: val('srdpVorname', p.vorname),
        titelDiplomarbeit: isAbaBetreuung
            ? p.titelDiplomarbeit || ''
            : daTitle || p.titelDiplomarbeit,
        betreuungslehrerDA: isAbaBetreuung
            ? p.betreuungslehrerDA || ''
            : betreuung || p.betreuungslehrerDA,
        hatABA: hatAbaEl ? !!hatAbaEl.checked : !!p.hatABA,
        titelABA: isAbaBetreuung ? daTitle || p.titelABA : p.titelABA || '',
        betreuungslehrerABA: isAbaBetreuung
            ? betreuung || p.betreuungslehrerABA
            : p.betreuungslehrerABA || '',
        abaPrefill: !!p.abaPrefill,
        variante,
        lfs: val('srdpLfs', p.lfs),
        lehrerLFS: readTeacherMulti('lehrerLFS', p.lehrerLFS),
        lehrerBKO: readTeacherMulti('lehrerBKO', p.lehrerBKO),
        lehrerAM: readTeacherMulti('lehrerAM', p.lehrerAM),
        wahlfachMuendlich: val('srdpWahlfach', p.wahlfachMuendlich || p.wahlfach),
        seminar: (function () {
            const typed = String((g('srdpSeminarText') && g('srdpSeminarText').value) || '').trim();
            const picked = String((g('srdpSeminar') && g('srdpSeminar').value) || '').trim();
            return typed || picked || String(p.seminar || '').trim();
        })(),
        lehrerWahlfach: readTeacherMulti('lehrerWahlfach', p.lehrerWahlfach),
        schwerpunktFach: val('srdpSchwerpunkt', p.schwerpunktFach),
        lehrerSchwerpunkt: readTeacherMulti('lehrerSchwerpunkt', p.lehrerSchwerpunkt),
        fachkolloquium: val('srdpFachkolloquium', p.fachkolloquium),
        lehrerFachkolloquium: readTeacherMulti('lehrerFachkolloquium', p.lehrerFachkolloquium),
        fachtheorie: val('srdpFachtheorie', p.fachtheorie),
        lehrerFachtheorie: readTeacherMulti('lehrerFachtheorie', p.lehrerFachtheorie),
        klausurKombi: val('srdpKlausurKombi', p.klausurKombi),
        muendlichExtra: val('srdpMuendlichExtra', p.muendlichExtra),
        lehrerMuendlichExtra: readTeacherMulti('lehrerMuendlichExtra', p.lehrerMuendlichExtra),
        vierteKlausur: val('srdpVierteKlausur', p.vierteKlausur),
        muendlich1: val('srdpMuendlich1', p.muendlich1),
        muendlich2: val('srdpMuendlich2', p.muendlich2),
        muendlich3: val('srdpMuendlich3', p.muendlich3),
        lehrerMuendlich1: readTeacherMulti('lehrerMuendlich1', p.lehrerMuendlich1),
        lehrerMuendlich2: readTeacherMulti('lehrerMuendlich2', p.lehrerMuendlich2),
        lehrerMuendlich3: readTeacherMulti('lehrerMuendlich3', p.lehrerMuendlich3),
        bestaetigung,
        wizardStep: step
    });
}

export { isSeminarWahlfach, buildPruefplanKurz, optionsHtml };
