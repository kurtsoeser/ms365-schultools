/**
 * UI-Rendering für den Schularbeiten-Planer.
 */
import {
    validateSchularbeit,
    computeDashboardKpis,
    buildWeeklyDistribution,
    toIsoDateOnly,
    dateInWindow
} from './schularbeiten-planer-logic.js';
import {
    viewsForRole,
    filterSchularbeiten,
    labelMaps,
    formatDeDate,
    statusLabel,
    roleLabel,
    emptyForm,
    canEditSchularbeit,
    canAdminDecide,
    fachMetaForCode,
    planerPublicUrl,
    buildEmbedSnippet,
    scopeFromState,
    resolveStudentKlasseCode
} from './schularbeiten-planer-state.js';
import { printTableCaption } from './schularbeiten-planer-export.js';

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
        <div class="sa-nav__role">
          <span class="sa-nav__role-label">Rolle (Demo)</span>
          <div class="sa-nav__role-btns">
            <button type="button" class="sa-chip${state.role === 'lehrer' ? ' is-active' : ''}" data-sa-role="lehrer">Lehrer</button>
            <button type="button" class="sa-chip${state.role === 'admin' ? ' is-active' : ''}" data-sa-role="admin">Admin</button>
            <button type="button" class="sa-chip${state.role === 'schueler' ? ' is-active' : ''}" data-sa-role="schueler">Schüler</button>
          </div>
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
                    : 'Demo: Klasse wählen, wenn keine Schüler-E-Mail in den Stammdaten.'
            }</p>`
                  : '<p class="sa-nav__hint">Später Entra-Gruppen. Demo-Umschalter nur zur Vorschau.</p>'
          }
        </div>
      </aside>
      <div class="sa-main">
        <header class="sa-top">
          <div class="sa-top__site">
            <label for="saSiteUrl">SharePoint-Site</label>
            <div class="sa-top__row">
              <input type="url" id="saSiteUrl" value="${esc(state.siteUrl)}" placeholder="https://…sharepoint.com/sites/Intranet" spellcheck="false">
              <button type="button" class="btn" id="saBtnLoad"><i class="bi bi-arrow-repeat"></i>Laden</button>
              ${
                  isSchueler
                      ? ''
                      : `<label class="btn" for="saImportJson">
                <i class="bi bi-filetype-json"></i>JSON importieren
              </label>
              <input type="file" id="saImportJson" accept=".json,application/json" hidden>
              <a class="btn" href="sharepoint-liste-schularbeiten.html" style="text-decoration:none"><i class="bi bi-list-ul"></i>Listen</a>`
              }
            </div>
            ${
                state.localDemoOnly
                    ? '<p class="sa-import-hint">Lokaler Demo-Import aktiv – Anzeige ohne SharePoint. Zum Schreiben: Site laden + erneut importieren und „Auf SharePoint schreiben“ wählen.</p>'
                    : ''
            }
          </div>
          <div class="sa-top__user">
            <span>Angemeldet als <span class="sa-badge sa-badge--info">${esc(roleLabel(state.role))}</span>
            ${isSchueler && klasseCode ? `<span class="sa-badge">${esc(klasseLabel)}</span>` : ''}</span>
            ${
                state.accountEmail
                    ? `<small>${esc(state.accountName || state.accountEmail)}</small>`
                    : '<small class="muted">Bitte oben rechts anmelden</small>'
            }
          </div>
        </header>
        ${state.error ? `<div class="sa-alert sa-alert--bad" role="alert">${esc(state.error)}</div>` : ''}
        ${
            state.roleHint
                ? `<div class="sa-alert sa-alert--info" role="status">${esc(state.roleHint)}
            ${
                state.role !== 'schueler'
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

function filterBar(state, opts) {
    const sd = state.stammdaten;
    const scopeAll = typeof opts === 'boolean' ? opts : !!(opts && opts.scopeAll);
    const onlyMine = typeof opts === 'object' && opts ? !!opts.onlyMine : false;
    const isSchueler = state.role === 'schueler';
    const scope = scopeFromState(state, { scopeAll: scopeAll && !isSchueler, onlyMine: onlyMine && !isSchueler });
    const filteredCount = filterSchularbeiten(state.items, state.filters, scope).length;
    if (isSchueler) {
        return `
    <div class="sa-filters" data-sa-filters>
      <select id="saFilterFach" aria-label="Fach">${optionList(sd.subjects, state.filters.fach, 'Alle Fächer')}</select>
      <button type="button" class="btn" id="saFilterReset"><i class="bi bi-arrow-counterclockwise"></i>Zurücksetzen</button>
      <span class="sa-filters__meta">${filteredCount} fixierte Termine</span>
    </div>`;
    }
    return `
    <div class="sa-filters" data-sa-filters>
      <select id="saFilterKlasse" aria-label="Klasse">${optionList(sd.classes, state.filters.klasse, 'Alle Klassen')}</select>
      <select id="saFilterFach" aria-label="Fach">${optionList(sd.subjects, state.filters.fach, 'Alle Fächer')}</select>
      <select id="saFilterLehrer" aria-label="Lehrer">${optionList(sd.teachers, state.filters.lehrer, 'Alle Lehrer:innen')}</select>
      <select id="saFilterStatus" aria-label="Status">
        <option value="">Alle Status</option>
        <option value="beantragt"${state.filters.status === 'beantragt' ? ' selected' : ''}>Beantragt</option>
        <option value="fixiert"${state.filters.status === 'fixiert' ? ' selected' : ''}>Fixiert</option>
        <option value="abgelehnt"${state.filters.status === 'abgelehnt' ? ' selected' : ''}>Abgelehnt</option>
      </select>
      <button type="button" class="btn" id="saFilterReset"><i class="bi bi-arrow-counterclockwise"></i>Zurücksetzen</button>
      <span class="sa-filters__meta">${filteredCount} Treffer</span>
    </div>`;
}

function renderView(state) {
    if (!state.bootstrapped && !state.ctx && !state.localDemoOnly) {
        return `<section class="tm-panel"><p>Site-URL eintragen und <strong>Laden</strong> – oder zuerst die Listen einrichten.</p>
          <p><a href="sharepoint-liste-schularbeiten.html">→ Schularbeiten-Listen anlegen</a></p></section>`;
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
    const kpis = computeDashboardKpis(items, undefined, state.rules);
    const labels = labelMaps(state.stammdaten);
    const dist = buildWeeklyDistribution(items, { weekCount: 8 });
    const listItems = items.slice().sort((a, b) => String(a.datum).localeCompare(String(b.datum)));
    const upcoming = listItems
        .filter((s) => s.status === 'fixiert' || s.status === 'beantragt')
        .slice(0, 8);

    return `
    ${filterBar(state, { scopeAll: state.role === 'admin' })}
    <section class="sa-hero-panel">
      <h2>Willkommen zurück</h2>
      <p>Schularbeitstermine nach SchUG &amp; LBVO – Daten in SharePoint, Stammdaten aus den Schultools.</p>
    </section>
    <div class="sa-kpi-grid">
      <div class="sa-kpi"><span>Offene Anträge</span><strong>${kpis.offen}</strong></div>
      <div class="sa-kpi"><span>Fixiert (2 Wochen)</span><strong>${kpis.fixiertNaechste2Wochen}</strong></div>
      <div class="sa-kpi${kpis.konflikte ? ' sa-kpi--warn' : ''}"><span>Konflikte</span><strong>${kpis.konflikte}</strong></div>
      <div class="sa-kpi"><span>Diese Woche</span><strong>${kpis.dieseWoche}</strong></div>
    </div>
    <div class="sa-split sa-dash-split">
      <section class="tm-panel">
        <div class="tm-panel__head"><i class="bi bi-bar-chart"></i>
          <div><h3>Verteilung pro Kalenderwoche</h3><p>Nächste 8 Wochen · gestapelt nach Fach</p></div>
        </div>
        ${renderWeeklyChart(dist, labels, state.fachMeta)}
      </section>
      <section class="tm-panel">
        <div class="tm-panel__head"><i class="bi bi-calendar-event"></i><div><h3>Nächste Termine</h3><p>Nach aktuellem Filter / Rolle</p></div></div>
        ${
            upcoming.length
                ? `<div class="sa-table-wrap"><table class="sa-table"><thead><tr><th>Datum</th><th>Fach</th><th>Klasse</th><th>Thema</th><th>Status</th></tr></thead><tbody>
            ${upcoming
                .map(
                    (sa) => `<tr class="sa-row" style="--sa-fach:${esc(fachColor(sa.fachCode, -1, state.fachMeta))}">
              <td><button type="button" class="sa-link" data-sa-detail="${esc(sa.itemId)}">${esc(formatDeDate(sa.datum))}</button></td>
              <td>${fachChip(sa.fachCode, labels, state.fachMeta)}</td>
              <td>${esc(labels.klasse[sa.klasseCode] || sa.klasseCode)}</td>
              <td>${esc(sa.thema)}</td>
              <td>${statusBadge(sa.status)}</td>
            </tr>`
                )
                .join('')}
          </tbody></table></div>`
                : '<p class="muted">Keine bevorstehenden Schularbeiten.</p>'
        }
      </section>
    </div>
    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-list-ul"></i>
        <div><h3>Listenansicht</h3><p>${listItems.length} Einträge nach aktuellem Filter</p></div>
      </div>
      ${fachLegend(listItems, labels, state.fachMeta)}
      ${tableSchularbeiten(listItems, labels, {
          showActions: state.role === 'admin' || state.role === 'lehrer',
          fachMeta: state.fachMeta,
          role: state.role
      })}
    </section>
    ${renderIntranetPanel(state)}`;
}

function renderSchuelerDashboard(state) {
    const klasseCode = resolveStudentKlasseCode(state);
    const labels = labelMaps(state.stammdaten);
    const items = scopedItems(state, false)
        .slice()
        .sort((a, b) => String(a.datum).localeCompare(String(b.datum)));
    const upcoming = items.slice(0, 12);
    const klasseName = labels.klasse[klasseCode] || klasseCode || '–';

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
    <section class="sa-hero-panel">
      <h2>Schularbeiten · ${esc(klasseName)}</h2>
      <p>Nur freigegebene (fixierte) Termine Ihrer Klasse – ohne Anträge oder Entwürfe.</p>
    </section>
    <div class="sa-kpi-grid">
      <div class="sa-kpi"><span>Fixierte Termine</span><strong>${items.length}</strong></div>
      <div class="sa-kpi"><span>Klasse</span><strong>${esc(klasseName)}</strong></div>
    </div>
    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-calendar-check"></i>
        <div><h3>Termine der Klasse</h3><p>Chronologisch</p></div>
      </div>
      ${fachLegend(items, labels, state.fachMeta)}
      ${
          upcoming.length
              ? tableSchularbeiten(upcoming, labels, { showActions: false, fachMeta: state.fachMeta })
              : '<p class="muted">Noch keine fixierten Schularbeiten für diese Klasse.</p>'
      }
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

function fachLegend(items, labels, fachMeta) {
    const seen = new Set();
    const codes = [];
    (items || []).forEach((sa) => {
        const c = sa && sa.fachCode;
        if (!c || seen.has(c)) return;
        seen.add(c);
        codes.push(c);
    });
    if (!codes.length) return '';
    return `<div class="sa-fach-legend" aria-label="Fachfarben">
      ${codes.map((c) => fachChip(c, labels, fachMeta)).join('')}
    </div>`;
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
        ${state.listsMissingFachMeta ? ' · <strong>SA-FachMeta fehlt noch</strong> – Setup erneut starten.' : ''}
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
                maxDauer: Math.max(150, m.standardDauer || 100)
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
          <div class="tm-field"><label for="saThema">Thema</label>
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
            <div class="tm-field"><label for="saDauer">Dauer (Min.)</label>
              <input id="saDauer" type="number" min="50" max="300" step="25" value="${esc(f.dauerMinuten)}"></div>
          </div>
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
      ${items.length ? tableSchularbeiten(items, labels, { showActions: true, role: state.role, fachMeta: state.fachMeta }) : emptyHint}
    </section>`;
}

function tableSchularbeiten(items, labels, opts) {
    if (!items.length) return '<p class="muted">Keine Einträge.</p>';
    const showActions = opts && opts.showActions;
    const fachMeta = (opts && opts.fachMeta) || [];
    return `<div class="sa-table-wrap"><table class="sa-table"><thead><tr>
      <th>Datum</th><th>Fach</th><th>Klasse</th><th>Thema</th><th>Dauer</th><th>Sem.</th><th>Status</th>
      ${showActions ? '<th>Aktionen</th>' : ''}
    </tr></thead><tbody>
    ${items
        .map((sa) => {
            const canEdit = sa.status === 'beantragt';
            const color = fachColor(sa.fachCode, -1, fachMeta);
            return `<tr class="sa-row" style="--sa-fach:${esc(color)}">
        <td><button type="button" class="sa-link" data-sa-detail="${esc(sa.itemId)}">${esc(formatDeDate(sa.datum))}</button></td>
        <td>${fachChip(sa.fachCode, labels, fachMeta)}</td>
        <td>${esc(labels.klasse[sa.klasseCode] || sa.klasseCode)}</td>
        <td>${esc(sa.thema)}</td>
        <td>${esc(sa.dauerMinuten)}'</td>
        <td>${esc(sa.semester)}</td>
        <td>${statusBadge(sa.status)}</td>
        ${
            showActions
                ? `<td class="sa-actions">
            ${
                canEdit
                    ? `<button type="button" class="btn btn-sm" data-sa-edit="${esc(sa.itemId)}" title="Bearbeiten"><i class="bi bi-pencil"></i></button>
               <button type="button" class="btn btn-sm" data-sa-del="${esc(sa.itemId)}" title="Löschen"><i class="bi bi-trash"></i></button>`
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

function renderAdmin(state) {
    if (state.role !== 'admin') {
        return `<section class="tm-panel"><p>Nur für Administrator:innen.</p></section>`;
    }
    const labels = labelMaps(state.stammdaten);
    const open = state.items.filter((s) => s.status === 'beantragt');
    const rw = state.rules;

    return `
    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-inbox"></i>
        <div><h3>Offene Anträge</h3><p>Fixieren oder ablehnen</p></div>
      </div>
      ${
          open.length
              ? `<div class="sa-table-wrap"><table class="sa-table"><thead><tr>
          <th>Datum</th><th>Fach</th><th>Klasse</th><th>Lehrer:in</th><th>Thema</th><th>Status</th><th>Aktionen</th>
        </tr></thead><tbody>
        ${open
            .map(
                (sa) => `<tr class="sa-row" style="--sa-fach:${esc(fachColor(sa.fachCode, -1, state.fachMeta))}">
          <td>${esc(formatDeDate(sa.datum))}</td>
          <td>${fachChip(sa.fachCode, labels, state.fachMeta)}</td>
          <td>${esc(labels.klasse[sa.klasseCode] || sa.klasseCode)}</td>
          <td>${esc(labels.lehrer[sa.lehrerCode] || sa.lehrerCode)}</td>
          <td>${esc(sa.thema)}</td>
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
        <div><h3>SA-FachMeta</h3><p>Farbe, Kontingent und Standarddauer (optional)</p></div>
      </div>
      ${
          state.ctx && state.ctx.lists && state.ctx.lists.fachMeta
              ? `<div class="sa-form">
          <div class="sa-form__row">
            <div class="tm-field"><label for="saFmCode">Fach-Code</label>
              <select id="saFmCode">${optionList(state.stammdaten.subjects, '', 'Bitte wählen')}</select></div>
            <div class="tm-field"><label for="saFmFarbe">Farbe (Hex)</label>
              <div class="sa-color-input">
                <input id="saFmFarbePicker" type="color" value="#6366f1" aria-label="Farbe wählen">
                <input id="saFmFarbe" type="text" placeholder="#6366f1" maxlength="20" value="#6366f1">
              </div></div>
          </div>
          <div class="sa-form__row">
            <div class="tm-field"><label for="saFmPro">Pro Semester</label>
              <input id="saFmPro" type="number" min="0" max="6" value="2"></div>
            <div class="tm-field"><label for="saFmDauer">Standard-Dauer</label>
              <input id="saFmDauer" type="number" min="50" max="300" value="100"></div>
          </div>
          <button type="button" class="btn" id="saBtnAddFachMeta"><i class="bi bi-plus"></i>Fach-Meta speichern</button>
        </div>
        <ul class="sa-fenster-list">
          ${
              (state.fachMeta || []).length
                  ? state.fachMeta
                        .map(
                            (m) => `<li>
              <div><strong>${fachChip(m.fachCode, labels, state.fachMeta)}</strong>
              <span>${esc(m.farbe || '–')} · ${esc(m.proSemester)}/Sem. · ${esc(m.standardDauer)} Min.</span></div>
              <button type="button" class="btn btn-sm" data-sa-del-fachmeta="${esc(m.itemId)}" title="Löschen"><i class="bi bi-trash"></i></button>
            </li>`
                        )
                        .join('')
                  : '<li class="muted">Noch keine Fach-Meta – Codes aus Stammdaten wählen.</li>'
          }
        </ul>`
              : '<p class="muted">Liste SA-FachMeta fehlt. <a href="sharepoint-liste-schularbeiten.html">Listen-Paket erneut ausführen</a>.</p>'
      }
    </section>

    <section class="tm-panel">
      <div class="tm-panel__head"><i class="bi bi-sliders2"></i>
        <div><h3>Automatisierung</h3><p>Lokal im Browser gespeichert</p></div>
      </div>
      <label class="sa-check">
        <input type="checkbox" id="saSyncTermine"${state.settings && state.settings.syncSchultermine ? ' checked' : ''}>
        Bei Fixierung Eintrag in Schultermine schreiben (Kategorie Prüfung)
      </label>
      <div class="tm-field" style="margin-top:10px;max-width:320px;">
        <label for="saSyncList">Schultermine-Listenname</label>
        <input id="saSyncList" type="text" value="${esc((state.settings && state.settings.schultermineList) || 'Schultermine')}">
      </div>
      <button type="button" class="btn" id="saBtnSaveSettings" style="margin-top:10px;"><i class="bi bi-save"></i>Einstellungen speichern</button>
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
    ${renderIntranetPanel(state)}`;
}

function renderKalender(state) {
    const labels = labelMaps(state.stammdaten);
    const items =
        state.role === 'admin'
            ? filterSchularbeiten(state.items, state.filters, scopeFromState(state, { scopeAll: true }))
            : scopedItems(state, state.role !== 'schueler');

    const y = state.calYear;
    const m = state.calMonth;
    const first = new Date(y, m, 1);
    const startPad = (first.getDay() + 6) % 7; // Montag = 0
    const daysInMonth = new Date(y, m + 1, 0).getDate();
    const monthName = first.toLocaleDateString('de-AT', { month: 'long', year: 'numeric' });
    const today = toIsoDateOnly(new Date());

    const byDay = {};
    items.forEach((sa) => {
        const d = sa.datum;
        if (!d) return;
        if (!byDay[d]) byDay[d] = [];
        byDay[d].push(sa);
    });

    const cells = [];
    for (let i = 0; i < startPad; i++) cells.push('<div class="sa-cal__cell sa-cal__cell--empty"></div>');
    for (let day = 1; day <= daysInMonth; day++) {
        const iso = `${y}-${String(m + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const blocked = (state.windows || []).some((w) => w.typ === 'gesperrt' && dateInWindow(iso, w));
        const list = byDay[iso] || [];
        cells.push(`<div class="sa-cal__cell${iso === today ? ' is-today' : ''}${blocked ? ' is-blocked' : ''}">
          <div class="sa-cal__day">${day}</div>
          <div class="sa-cal__events">
            ${list
                .slice(0, 3)
                .map((sa) => {
                    const color = fachColor(sa.fachCode, -1, state.fachMeta);
                    const fg = contrastOn(color);
                    const fachLabel = labels.fach[sa.fachCode] || sa.fachCode || '?';
                    return `<button type="button" class="sa-cal__ev sa-cal__ev--${esc(
                        sa.status
                    )}" style="--sa-fach:${esc(color)};--sa-fach-fg:${esc(fg)}" data-sa-detail="${esc(
                        sa.itemId
                    )}" title="${esc(fachLabel + ' · ' + (sa.thema || '') + ' · ' + statusLabel(sa.status))}">${esc(
                        fachLabel
                    )}</button>`;
                })
                .join('')}
            ${list.length > 3 ? `<span class="sa-cal__more">+${list.length - 3}</span>` : ''}
          </div>
        </div>`);
    }

    return `
    ${filterBar(state, { scopeAll: state.role === 'admin' })}
    <section class="tm-panel">
      <div class="sa-cal__head">
        <h3>${esc(monthName)}</h3>
        <div class="sa-cal__nav">
          <button type="button" class="btn" id="saCalPrev" aria-label="Vorheriger Monat"><i class="bi bi-chevron-left"></i></button>
          <button type="button" class="btn" id="saCalToday">Heute</button>
          <button type="button" class="btn" id="saCalNext" aria-label="Nächster Monat"><i class="bi bi-chevron-right"></i></button>
        </div>
      </div>
      <div class="sa-cal__weekdays"><span>Mo</span><span>Di</span><span>Mi</span><span>Do</span><span>Fr</span><span>Sa</span><span>So</span></div>
      <div class="sa-cal__grid">${cells.join('')}</div>
      <div class="sa-cal__legend">
        <span><i class="sa-dot sa-dot--warn"></i>Beantragt (gestrichelt)</span>
        <span><i class="sa-dot sa-dot--ok"></i>Fixiert</span>
        <span><i class="sa-dot sa-dot--bad"></i>Abgelehnt (blass)</span>
        <span><i class="sa-blocked-swatch"></i>Gesperrter Tag</span>
      </div>
      ${fachLegend(items, labels, state.fachMeta)}
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
    </div>
    <section class="tm-panel sa-print-area" id="saPrintArea">
      <div class="tm-panel__head"><i class="bi bi-eye"></i>
        <div><h3>Vorschau (${fixed.length} Termine)</h3><p>${esc(printTableCaption(fixed))}</p></div>
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
    const canDecide = canAdminDecide(sa, scope);
    const check = validateSchularbeit({
        draft: sa,
        existing: state.items,
        rules: state.rules,
        windows: state.windows
    });

    return `
    <div class="sa-modal" role="dialog" aria-modal="true" aria-labelledby="saDetailTitle">
      <div class="sa-modal__card">
        <header>
          <h3 id="saDetailTitle">${esc(sa.thema || 'Schularbeit')}</h3>
          <button type="button" class="btn btn-sm" id="saDetailClose" aria-label="Schließen"><i class="bi bi-x-lg"></i></button>
        </header>
        <dl class="sa-dl">
          <div><dt>Datum</dt><dd>${esc(formatDeDate(sa.datum))}</dd></div>
          <div><dt>Fach</dt><dd>${fachChip(sa.fachCode, labels, state.fachMeta)}</dd></div>
          <div><dt>Klasse</dt><dd>${esc(labels.klasse[sa.klasseCode] || sa.klasseCode)}</dd></div>
          <div><dt>Lehrer:in</dt><dd>${esc(labels.lehrer[sa.lehrerCode] || sa.lehrerCode)}</dd></div>
          <div><dt>Dauer</dt><dd>${esc(sa.dauerMinuten)} Min.</dd></div>
          <div><dt>Status</dt><dd>${statusBadge(sa.status)}</dd></div>
        </dl>
        ${sa.notiz ? `<p><strong>Notiz:</strong> ${esc(sa.notiz)}</p>` : ''}
        ${sa.ablehnungsGrund ? `<p><strong>Ablehnung:</strong> ${esc(sa.ablehnungsGrund)}</p>` : ''}
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
                  ? `<button type="button" class="btn btn-success" data-sa-fix="${esc(sa.itemId)}"><i class="bi bi-check-lg"></i>Fixieren</button>
             <button type="button" class="btn" data-sa-reject="${esc(sa.itemId)}"><i class="bi bi-x-lg"></i>Ablehnen</button>`
                  : ''
          }
          ${
              canEdit
                  ? `<button type="button" class="btn" data-sa-edit="${esc(sa.itemId)}"><i class="bi bi-pencil"></i>Bearbeiten</button>
             <button type="button" class="btn" data-sa-del="${esc(sa.itemId)}"><i class="bi bi-trash"></i>Löschen</button>`
                  : ''
          }
          <button type="button" class="btn" id="saDetailClose2">Schließen</button>
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
        thema: val('saThema'),
        fachCode: val('saFach'),
        klasseCode: val('saKlasse'),
        lehrerCode: val('saLehrer'),
        lehrerEmail,
        datum: val('saDatum'),
        dauerMinuten: Number(val('saDauer')) || 100,
        semester: val('saSemester') || 'WS',
        notiz: val('saNotiz')
    };
}

export function readFiltersFromDom() {
    const g = (id) => {
        const el = document.getElementById(id);
        return el ? String(el.value || '') : '';
    };
    return {
        klasse: g('saFilterKlasse'),
        fach: g('saFilterFach'),
        lehrer: g('saFilterLehrer'),
        status: g('saFilterStatus')
    };
}

export function readFensterForm() {
    const g = (id) => {
        const el = document.getElementById(id);
        return el ? String(el.value || '').trim() : '';
    };
    return {
        titel: g('saTfTitel'),
        startdatum: g('saTfVon'),
        enddatum: g('saTfBis'),
        beschreibung: g('saTfDesc'),
        typ: 'gesperrt'
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
