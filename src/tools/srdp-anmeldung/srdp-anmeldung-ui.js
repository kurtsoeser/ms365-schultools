/**
 * UI-Rendering sRDP-Anmeldung (Calm UX, sektioniertes Nav, Multi-Filter, Kalender).
 */
import {
    navSectionsForRole,
    roleLabel,
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
    buildAnmeldungSummary,
    teacherRolesOnItem,
    buildPruefplanKurz,
    isSrdpAnmeldeFensterOffen,
    findAnmeldungBySchuelerEmail
} from './srdp-anmeldung-logic.js';
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
    if (state.role === 'schueler') {
        return 'Melde dich für den Haupttermin ' + jahr + ' an (' + profile + ').';
    }
    if (state.role === 'lehrer') {
        return 'Prüfungen und Klassen, in denen du beteiligt bist · ' + profile + ' ' + jahr;
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
    return `<header class="srdp-top">
      <div>
        <h2>${esc(topTitle(state))}</h2>
        <p>${esc(topHint(state))}</p>
      </div>
      <div class="srdp-top__user">
        <span>Angemeldet als <span class="srdp-badge srdp-badge--info">${esc(roleLabel(state.role))}</span></span>
        ${phaseBadge(state)}
        ${
            state.localDemoOnly
                ? '<span class="srdp-badge srdp-badge--warn">Demo lokal</span>'
                : ''
        }
        ${
            state.demoPersonaEmail
                ? `<small>Persona ${esc(state.demoPersonaName || state.demoPersonaEmail)}</small>`
                : state.accountEmail
                  ? `<small>${esc(state.accountName || state.accountEmail)}</small>`
                  : '<small>Bitte oben rechts anmelden</small>'
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

function renderSummaryCard(item, profileId, opts) {
    const o = opts || {};
    const summary = buildAnmeldungSummary(item, profileId);
    return `<button type="button" class="srdp-summary srdp-summary--clickable" data-srdp-open-detail="${esc(
        item.itemId
    )}">
      <div class="srdp-summary__head">
        <h3>${esc(o.heading || summary.title || 'Anmeldung')}</h3>
        <p>${esc(o.sub || summary.pruefplan || 'Deine gewählten Prüfungsteile')}</p>
      </div>
      <dl>
        ${summary.rows
            .slice(0, 6)
            .map((r) => `<dt>${esc(r.label)}</dt><dd>${esc(r.value)}</dd>`)
            .join('')}
      </dl>
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

function renderTable(items, state, opts) {
    const o = opts || {};
    const list = Array.isArray(items) ? items : [];
    if (!list.length) {
        return quietEmpty('<strong>Keine Anmeldungen</strong> · In dieser Ansicht ist nichts vorhanden.');
    }
    return `<div class="srdp-table-wrap"><table class="srdp-table">
      <thead><tr>
        <th>Name</th><th>Klasse</th><th>Variante</th><th>LFS</th><th>Wahlfach</th>
        ${o.showRoles ? '<th>Deine Rolle</th>' : ''}
        <th>Betreuung DA</th>
        <th></th>
      </tr></thead>
      <tbody>
        ${list
            .map((it) => {
                const roles = o.showRoles
                    ? teacherRolesOnItem(it, state.teacherLabel)
                          .map((r) => `<span class="srdp-pill srdp-pill--role">${esc(r.label)}</span>`)
                          .join('')
                    : '';
                return `<tr>
            <td><strong>${esc(it.title || it.nachname + ', ' + it.vorname)}</strong></td>
            <td>${esc(it.klasse)}</td>
            <td><span class="srdp-pill">${esc(it.variante)}</span></td>
            <td>${esc(it.lfs)}${it.lehrerLFS ? '<br><small>' + esc(it.lehrerLFS) + '</small>' : ''}</td>
            <td>${esc(it.wahlfachMuendlich)}${
                    it.seminar ? '<br><small>Seminar: ' + esc(it.seminar) + '</small>' : ''
                }</td>
            ${o.showRoles ? `<td>${roles || '–'}</td>` : ''}
            <td>${esc(it.betreuungslehrerDA || it.betreuungslehrerABA || '–')}</td>
            <td><button type="button" class="srdp-text-link" data-srdp-open-detail="${esc(
                it.itemId
            )}">Details</button></td>
          </tr>`;
            })
            .join('')}
      </tbody>
    </table></div>`;
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
    let action = quietEmpty(
        '<strong>Alles erledigt</strong> · Aktuell ist keine Prüfung bei dir offen.',
        `<div class="srdp-actions"><button type="button" class="btn" data-srdp-view-jump="liste">Alle Anmeldungen</button></div>`
    );
    if (mine.length) {
        const first = mine[0];
        action = `<section class="srdp-calm-banner srdp-calm-banner--action" role="status">
      <div class="srdp-calm-banner__text">
        <p class="srdp-calm-banner__eyebrow">Was braucht deine Aufmerksamkeit?</p>
        <p><strong>${
            mine.length === 1
                ? '1 Prüfung ist dir zugeordnet'
                : mine.length + ' Prüfungen sind dir zugeordnet'
        }</strong></p>
        <p class="muted">${esc(first.title || first.nachname)} · ${esc(first.klasse)} · ${esc(
            first.variante
        )}</p>
      </div>
      <button type="button" class="btn btn-success" data-srdp-view-jump="meine-pruefungen">Meine Prüfungen →</button>
    </section>`;
    }
    return `${action}
    <section class="srdp-panel srdp-panel--calm">
      <h3>Als Nächstes</h3>
      <p class="srdp-panel__hint">Zuordnungen als Betreuung, LFS, BKO oder Wahlfach.</p>
      ${
          mine.length
              ? renderTable(mine.slice(0, 6), state, { showRoles: true })
              : '<p class="muted">Keine kommenden Zuordnungen.</p>'
      }
      <div class="srdp-actions">
        <button type="button" class="srdp-text-link" data-srdp-view-jump="meine-pruefungen">Alle meine Prüfungen →</button>
        ${
            kv.length
                ? '<button type="button" class="srdp-text-link" data-srdp-view-jump="meine-klasse">Meine Klasse →</button>'
                : ''
        }
      </div>
    </section>`;
}

function renderLeitungDashboard(state) {
    const items = state.items || [];
    const o = state.options || {};
    const fehlend = klassenOhneAnmeldung(items, o.klassen || []);
    const nacharbeit = countNacharbeitPending(items);
    const stats = computeAnmeldungStats(items);

    const attention = fehlend.length
        ? `<section class="srdp-calm-banner srdp-calm-banner--action" role="status">
      <div class="srdp-calm-banner__text">
        <p class="srdp-calm-banner__eyebrow">Was braucht deine Aufmerksamkeit?</p>
        <p><strong>${
            fehlend.length === 1
                ? '1 Klasse ohne Anmeldung'
                : fehlend.length + ' Klassen ohne Anmeldung'
        }</strong></p>
        <p class="muted">${esc(fehlend.slice(0, 4).join(', '))}${fehlend.length > 4 ? ' …' : ''}</p>
      </div>
      <button type="button" class="btn btn-success" data-srdp-jump-liste data-srdp-jump-klasse="${esc(
          fehlend[0] || ''
      )}">Liste öffnen →</button>
    </section>`
        : quietEmpty('<strong>Aktuell wartet keine Klasse ohne Anmeldung.</strong>');

    const ops = nacharbeit
        ? `<section class="srdp-calm-banner srdp-calm-banner--ops" role="status">
      <div class="srdp-calm-banner__icon" aria-hidden="true"><i class="bi bi-clipboard-check"></i></div>
      <div class="srdp-calm-banner__text">
        <p class="srdp-calm-banner__eyebrow">Organisatorische Nacharbeit</p>
        <p><strong>${
            nacharbeit === 1
                ? '1 bestätigte Anmeldung wartet auf die Prüfungsliste'
                : nacharbeit + ' bestätigte Anmeldungen warten auf die Prüfungsliste'
        }</strong></p>
        <p class="muted">Bestätigung (Schüler) und Prüfungsliste sind getrennt.</p>
      </div>
      <button type="button" class="btn" data-srdp-view-jump="liste">Anmeldungen →</button>
    </section>`
        : '';

    return `
    <section class="srdp-panel srdp-panel--calm">
      <h2 class="srdp-sr-only">Aufmerksamkeit</h2>
      ${attention}
      ${ops}
    </section>
    <section class="srdp-panel srdp-panel--calm">
      <h3>Stand ${esc(o.jahr || '')}</h3>
      <div class="srdp-kpis">
        <div class="srdp-kpi"><strong>${stats.total}</strong><span>Anmeldungen</span></div>
        <div class="srdp-kpi"><strong>${stats.byKlasse.length}</strong><span>Klassen</span></div>
        <div class="srdp-kpi"><strong>${stats.byVariante.length}</strong><span>Varianten</span></div>
        <div class="srdp-kpi"><strong>${nacharbeit}</strong><span>Nacharbeit offen</span></div>
      </div>
      <button type="button" class="srdp-text-link" data-srdp-view-jump="auswertung">Zur Auswertung →</button>
    </section>`;
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
    return mine.map((it) => renderSummaryCard(it, state.options.profileId)).join('');
}

function renderAuswertung(state) {
    const stats = computeAnmeldungStats(filterItemsForView(state.items || [], state));
    return `${renderFilters(state)}
    <div class="srdp-kpis">
      <div class="srdp-kpi"><strong>${stats.total}</strong><span>gefiltert</span></div>
      <div class="srdp-kpi"><strong>${stats.mitSeminar}</strong><span>Seminar</span></div>
    </div>
    <section class="srdp-panel srdp-panel--calm"><h3>Klassen</h3>${renderBars(
        stats.byKlasse,
        stats.total
    )}</section>
    <section class="srdp-panel srdp-panel--calm"><h3>Varianten</h3>${renderBars(
        stats.byVariante,
        stats.total
    )}</section>
    <section class="srdp-panel srdp-panel--calm"><h3>Wahlfächer</h3>${renderBars(
        stats.byWahlfach,
        stats.total
    )}</section>`;
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
    return `<section class="srdp-panel srdp-panel--it">
      <h3>Listen &amp; Jahr</h3>
      <p class="srdp-panel__hint">IT / Administration: SharePoint-Liste für den Haupttermin verbinden. Lehrer sehen diesen Bereich nicht.</p>
      <div class="srdp-form-grid">
        <div class="srdp-field srdp-field--full">
          <label for="srdpSiteUrl">SharePoint-Website</label>
          <input type="url" id="srdpSiteUrl" value="${esc(state.siteUrl || '')}" placeholder="https://….sharepoint.com/sites/…">
        </div>
        <div class="srdp-field">
          <label>Profil</label>
          <input type="text" value="${esc(o.profileLabel || '')}" readonly>
        </div>
        <div class="srdp-field">
          <label>Terminjahr</label>
          <input type="text" value="${esc(o.jahr || '')}" readonly>
        </div>
        <div class="srdp-field srdp-field--full">
          <label>Listenname</label>
          <input type="text" value="${esc(o.listTitle || sync.listTitle || '')}" readonly>
        </div>
      </div>
      <p class="srdp-panel__hint">Sync: <strong>${esc(syncSource)}</strong>${
          sync.listTitle ? ' · Liste „' + esc(sync.listTitle) + '“' : ''
      }${
          sync.lastPullAt ? ' · zuletzt ' + esc(formatSyncPullAt(sync.lastPullAt)) : ''
      }${sync.connected ? ' · verbunden' : ' · nicht verbunden'}${
          state.localDemoOnly ? ' · Demo local-first' : ''
      } · ${withItemId}/${(state.items || []).length} mit itemId</p>
      <div class="srdp-actions">
        <button type="button" class="btn btn-success" data-srdp-save-site><i class="bi bi-save"></i>Site speichern &amp; laden</button>
        <button type="button" class="btn" data-srdp-reload><i class="bi bi-arrow-repeat"></i>Von SharePoint laden</button>
        <button type="button" class="btn" data-srdp-demo><i class="bi bi-magic"></i>Demo lokal</button>
        <button type="button" class="btn" data-srdp-demo-reset><i class="bi bi-arrow-counterclockwise"></i>Demo zurücksetzen</button>
        <a class="btn" href="sharepoint-liste-srdp.html" style="text-decoration:none;"><i class="bi bi-sliders"></i>Vollständiges Setup</a>
        <a class="btn" href="srdp-mail-setup.html" style="text-decoration:none;"><i class="bi bi-envelope-check"></i>Bestätigungsmail-Flow</a>
      </div>
    </section>
    <section class="srdp-panel srdp-panel--it">
      <h3>Stammdaten / Diagnose</h3>
      <p class="srdp-panel__hint">Quelle: ${esc(live.source || '–')} · Klassen ${
          Number(live.classCount) || 0
      } · Lehrkräfte ${Number(live.teacherCount) || 0}${
          live.ok ? '' : live.error ? ' · ' + esc(live.error) : ''
      }</p>
      <dl class="srdp-summary" style="border:0;border-radius:0;box-shadow:none;">
        <dt>Klassen (Setup)</dt><dd>${esc((o.klassen || []).join(', ') || '–')}</dd>
        <dt>LFS</dt><dd>${esc((o.lfs || []).join(', ') || '–')}</dd>
        <dt>Wahlfächer</dt><dd>${esc((o.wahlfaecher || []).slice(0, 8).join(', ') || '–')}${
            (o.wahlfaecher || []).length > 8 ? ' …' : ''
        }</dd>
      </dl>
      <div class="srdp-form-grid" style="margin-top:12px">
        <div class="srdp-field">
          <label for="srdpAnmeldeFrist">Anmeldefrist</label>
          <input type="date" id="srdpAnmeldeFrist" value="${esc(
              String(o.anmeldeFrist || '').slice(0, 10)
          )}">
        </div>
        <div class="srdp-field">
          <label for="srdpHaupttermin">Haupttermin</label>
          <input type="date" id="srdpHaupttermin" value="${esc(
              String(o.haupttermin || '').slice(0, 10)
          )}">
        </div>
      </div>
      <div class="srdp-actions">
        <button type="button" class="btn btn-success" data-srdp-save-termine><i class="bi bi-calendar-check"></i>Frist &amp; Termin speichern</button>
      </div>
    </section>`;
}

function buildCalendarEvents(state) {
    const o = state.options || {};
    /** @type {object[]} */
    const events = [];
    if (o.anmeldeFrist) {
        events.push({
            id: 'frist',
            title: 'Anmeldefrist sRDP',
            start: o.anmeldeFrist,
            allDay: true,
            tone: 'warn',
            subtitle: 'Haupttermin ' + (o.jahr || '')
        });
    }
    if (o.haupttermin) {
        events.push({
            id: 'haupt',
            title: 'Haupttermin Matura/sRDP',
            start: o.haupttermin,
            allDay: true,
            tone: 'ok',
            subtitle: o.profileLabel || 'HAK'
        });
    }
    const list = filterItemsForView(state.items || [], state);
    list.forEach((it) => {
        const start = toIsoDateOnly(it.created);
        if (!start) return;
        events.push({
            id: String(it.itemId || ''),
            title: it.title || buildAnmeldungSummary(it).title || 'Anmeldung',
            start,
            allDay: true,
            tone: it.nacharbeitErledigt ? 'ok' : 'info',
            subtitle: [it.klasse, it.variante].filter(Boolean).join(' · '),
            openAttr: 'data-srdp-open-detail="' + esc(it.itemId) + '"'
        });
    });
    return events;
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
        hint: 'Anmeldefrist, Haupttermin und eingegangene Anmeldungen · Agenda, Woche, Monat.',
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
                    ) + renderSummaryCard(existing, state.options.profileId)
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
            return renderTable(itemsForCurrentRole(state), state, { showRoles: true });
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
                renderTable(rows, state, { showRoles: false })
            );
        }
        case 'liste':
            return (
                renderFilters(state) +
                renderTable(filterItemsForView(state.items || [], state), state, {
                    showRoles: state.role === 'lehrer'
                })
            );
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
      ${detail ? renderAnmeldungDrawer(detail, state) : ''}
    </div>`;
}

export function readFormFromDom() {
    const g = (id) => document.getElementById(id);
    const best = g('srdpBestaetigung');
    const bestaetigung = best
        ? best.type === 'checkbox'
            ? !!best.checked
            : String(best.value || '') === '1' || String(best.value || '').toLowerCase() === 'true'
        : false;
    const formEl = document.getElementById('srdpForm');
    const step = formEl ? Number(formEl.getAttribute('data-srdp-wizard-step') || 1) : 1;
    return emptyForm({
        klasse: g('srdpKlasse') && g('srdpKlasse').value,
        nachname: g('srdpNachname') && g('srdpNachname').value,
        vorname: g('srdpVorname') && g('srdpVorname').value,
        titelDiplomarbeit: g('srdpDA') && g('srdpDA').value,
        betreuungslehrerDA: g('srdpBetreuung') && g('srdpBetreuung').value,
        variante: g('srdpVariante') && g('srdpVariante').value,
        lfs: g('srdpLfs') && g('srdpLfs').value,
        lehrerLFS: g('srdpLehrerLfs') && g('srdpLehrerLfs').value,
        lehrerBKO: g('srdpLehrerBko') && g('srdpLehrerBko').value,
        wahlfachMuendlich: g('srdpWahlfach') && g('srdpWahlfach').value,
        seminar: g('srdpSeminar') && g('srdpSeminar').value,
        lehrerWahlfach: g('srdpLehrerWahl') && g('srdpLehrerWahl').value,
        bestaetigung,
        wizardStep: step
    });
}

export { isSeminarWahlfach, buildPruefplanKurz, optionsHtml };
