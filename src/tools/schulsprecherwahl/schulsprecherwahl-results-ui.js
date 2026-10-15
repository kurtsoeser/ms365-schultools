/**
 * Auswertung / Kundmachung – Layout analog Praxis-Screenshot.
 */

import {
    ROUND_MAIN,
    ROUND_RUNOFF,
    mainResultRows,
    runoffResultRows,
    turnoutPercent,
    sharePercent
} from './schulsprecherwahl-logic.js';
import { normalizeElection, statusLabel } from './schulsprecherwahl-model.js';
import { evaluateElectionAsync, ballotsForRoundAsync } from './schulsprecherwahl-service.js';

/**
 * @param {string} name
 */
function initials(name) {
    const parts = String(name || '')
        .trim()
        .split(/\s+/)
        .filter(Boolean);
    if (parts.length >= 2) return (parts[0][0] + parts[1][0]).toUpperCase();
    return (parts[0] || '?').slice(0, 2).toUpperCase();
}

/**
 * @param {string} s
 */
function esc(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * @param {{ id: string, name: string, className?: string, email?: string }|null} cand
 * @param {string} role
 * @param {string} tag
 */
function officeCard(cand, role, tag) {
    if (!cand) {
        return `<article class="ssw-office"><div class="ssw-office__role">${esc(role)}</div>
      <div class="ssw-avatar">?</div>
      <div class="ssw-office__name">Noch offen</div></article>`;
    }
    const photoKey = cand.email || cand.id;
    return `<article class="ssw-office">
      <div class="ssw-office__role">${esc(role)}</div>
      <div class="ssw-avatar ms365-user-photo-thumb" data-ms365-user-photo="${esc(photoKey)}">
        <img alt="" hidden />
        <span class="ms365-user-photo-thumb__initials">${esc(initials(cand.name))}</span>
      </div>
      <div class="ssw-office__name">${esc(cand.name)}</div>
      <div class="ssw-office__class">${esc(cand.className || '')}</div>
      ${tag ? `<span class="ssw-badge">${esc(tag)}</span>` : ''}
    </article>`;
}

/**
 * @param {ReturnType<typeof normalizeElection>} election
 * @param {'main'|'runoff'} activeRound
 * @param {Awaited<ReturnType<typeof evaluateElectionAsync>>} evaled
 * @param {number} castCount
 */
export function renderResultsHtml(election, activeRound, evaled, castCount) {
    const e = normalizeElection(election);
    const byId = new Map(e.candidates.map((c) => [c.id, c]));
    const mainRound = e.rounds.find((r) => r.kind === ROUND_MAIN) || e.rounds[0];
    const runoffRound = e.rounds.find((r) => r.kind === ROUND_RUNOFF);
    const hasRunoff = !!runoffRound;
    const showRunoff = hasRunoff && activeRound === 'runoff';

    const ss = byId.get(e.offices.schulsprecherId || '');
    const s1 = byId.get(e.offices.stv1Id || '');
    const s2 = byId.get(e.offices.stv2Id || '');

    const ssTag =
        e.offices.decidedAfter === 'runoff'
            ? 'In der Stichwahl gewählt'
            : e.offices.decidedAfter === 'main'
              ? 'Im 1. Wahlgang gewählt'
              : '';
    const stvTag = 'Im 1. Wahlgang bestätigt';

    const statusBadge =
        e.status === 'finalized'
            ? e.offices.decidedAfter === 'runoff'
                ? '✓ Nach Stichwahl entschieden'
                : '✓ Im 1. Wahlgang entschieden'
            : statusLabel(e);

    const round = showRunoff ? runoffRound : mainRound;
    const stats = round ? round.stats : { cast: 0, valid: 0, invalid: 0, eligible: e.eligibleCount };
    const turnout = turnoutPercent(stats.cast || castCount, stats.eligible || e.eligibleCount);

    let tableHtml = '';
    if (showRunoff && evaled.runoffTally) {
        const rows = runoffResultRows(evaled.runoffTally, e.candidates, e.offices.schulsprecherId);
        tableHtml = rows
            .map((row) => {
                const pct = Math.round((row.barValue / row.barMax) * 100);
                return `<tr>
          <td><div class="ssw-cand">
            <div class="ssw-avatar ssw-avatar--sm ms365-user-photo-thumb" data-ms365-user-photo="${esc(row.email || row.id)}">
              <img alt="" hidden /><span class="ms365-user-photo-thumb__initials">${esc(initials(row.name))}</span>
            </div>
            <div><div class="ssw-cand__name">${esc(row.name)}</div>
            <div class="ssw-cand__class">${esc(row.className)}</div></div>
          </div></td>
          <td><div class="ssw-bar-wrap"><div class="ssw-bar" aria-hidden="true"><i style="width:${pct}%"></i></div>
            <strong>${row.votes}</strong></div>
            <div class="ssw-cand__class">${sharePercent(row.votes, evaled.runoffTally.valid)}&nbsp;%</div></td>
          <td>${row.badge ? `<span class="ssw-result-tag">${esc(row.badge)}</span>` : '—'}</td>
        </tr>`;
            })
            .join('');
    } else {
        const rows = mainResultRows(evaled.main, e.candidates, e.offices);
        tableHtml = rows
            .map((row) => {
                const pct = Math.round((row.barValue / row.barMax) * 100);
                return `<tr>
          <td><div class="ssw-cand">
            <div class="ssw-avatar ssw-avatar--sm ms365-user-photo-thumb" data-ms365-user-photo="${esc(row.email || row.id)}">
              <img alt="" hidden /><span class="ms365-user-photo-thumb__initials">${esc(initials(row.name))}</span>
            </div>
            <div><div class="ssw-cand__name">${esc(row.name)}</div>
            <div class="ssw-cand__class">${esc(row.className)}</div></div>
          </div></td>
          <td><div class="ssw-bar-wrap"><div class="ssw-bar" aria-hidden="true"><i style="width:${pct}%"></i></div>
            <strong>${row.points}</strong></div>
            <div class="ssw-cand__class">${row.firstPlace}× Erst · ${row.points} Pkt.</div></td>
          <td>${row.badge ? `<span class="ssw-result-tag">${esc(row.badge)}</span>` : '—'}</td>
        </tr>`;
            })
            .join('');
    }

    const dateLabel =
        round && round.countedAt
            ? new Date(round.countedAt).toLocaleDateString('de-AT', {
                  day: '2-digit',
                  month: '2-digit',
                  year: 'numeric'
              })
            : '—';

    const foot = showRunoff
        ? `Prozentanteile bezogen auf ${stats.valid || (evaled.runoffTally && evaled.runoffTally.valid) || 0} gültige Stimmen in diesem Wahlgang.`
        : `Wahlpunkte: 6 (Schulsprecher:in) · 5 (1. STV) · 4 (2. STV). Erstplatzierungen für die absolute Mehrheit.`;

    return `<div class="ssw-wrap ssw-results" data-election-id="${esc(e.id)}">
      <header class="ssw-hero">
        <div>
          <h2>${esc(e.title)}</h2>
          <p class="ssw-meta">${esc(e.scopeLabel)} · ${e.candidates.length} Kandidaturen · ${
              hasRunoff ? '2 Wahlgänge' : '1 Wahlgang'
          }</p>
        </div>
        <span class="ssw-badge">${esc(statusBadge)}</span>
      </header>

      <section class="ssw-panel">
        <h3>Aktuelle Ämterbesetzung</h3>
        <div class="ssw-offices">
          ${officeCard(ss, 'Schulsprecher:in', ssTag)}
          ${officeCard(s1, '1. Stellvertretung', s1 ? stvTag : '')}
          ${officeCard(s2, '2. Stellvertretung', s2 ? stvTag : '')}
        </div>
      </section>

      <section class="ssw-panel">
        <h3>Wahlgänge</h3>
        <div class="ssw-round-tabs" role="tablist">
          <button type="button" data-ssw-round="main" aria-selected="${!showRunoff}">1. Wahlgang</button>
          ${
              hasRunoff
                  ? `<button type="button" data-ssw-round="runoff" aria-selected="${showRunoff}">2. Wahlgang · Stichwahl</button>`
                  : ''
          }
        </div>
        <div class="ssw-round-head">
          <strong>${esc(dateLabel)}</strong>
          <span class="ssw-badge">${round && round.status === 'counted' ? 'Ausgezählt' : 'Offen'}</span>
        </div>
        <div class="ssw-stats">
          <div class="ssw-stat"><strong>${stats.cast || castCount}</strong><span>Abgegeben</span></div>
          <div class="ssw-stat"><strong>${stats.invalid || 0}</strong><span>Ungültig</span></div>
          <div class="ssw-stat"><strong>${stats.eligible || e.eligibleCount || '–'}</strong><span>Wahlberechtigt</span></div>
          <div class="ssw-stat"><strong>${turnout == null ? '–' : Math.round(turnout) + '%'}</strong><span>Beteiligung</span></div>
        </div>
        <table class="ssw-table">
          <thead>
            <tr>
              <th>Kandidat:in</th>
              <th>${showRunoff ? 'Stimmen' : 'Wahlpunkte'}</th>
              <th>Ergebnis dieses Wahlgangs</th>
            </tr>
          </thead>
          <tbody>${tableHtml}</tbody>
        </table>
        <p class="ssw-foot-note">${esc(foot)}</p>
      </section>
    </div>`;
}

/**
 * @param {HTMLElement} host
 * @param {ReturnType<typeof normalizeElection>} election
 * @param {{ initialRound?: 'main'|'runoff' }} [opts]
 */
export async function mountResults(host, election, opts) {
    let round =
        opts && opts.initialRound
            ? opts.initialRound
            : election.rounds.some((r) => r.kind === ROUND_RUNOFF)
              ? 'runoff'
              : 'main';

    const paint = async () => {
        host.innerHTML = `<p class="ssw-meta">Auswertung wird geladen …</p>`;
        try {
            const evaled = await evaluateElectionAsync(election.id);
            const active =
                round === 'runoff'
                    ? election.rounds.find((r) => r.kind === ROUND_RUNOFF)
                    : election.rounds.find((r) => r.kind === ROUND_MAIN) || election.rounds[0];
            const castList = active ? await ballotsForRoundAsync(election.id, active.id) : [];
            host.innerHTML = renderResultsHtml(election, round, evaled, castList.length);
            host.querySelectorAll('[data-ssw-round]').forEach((btn) => {
                btn.addEventListener('click', () => {
                    round = /** @type {'main'|'runoff'} */ (btn.getAttribute('data-ssw-round') || 'main');
                    paint();
                });
            });
            try {
                if (window.ms365UserPhotoThumb && typeof window.ms365UserPhotoThumb.scan === 'function') {
                    window.ms365UserPhotoThumb.scan(host);
                }
            } catch {
                /* ignore */
            }
        } catch (err) {
            host.innerHTML = `<p class="ssw-badge ssw-badge--warn">${esc(
                err && err.message ? err.message : String(err)
            )}</p>`;
        }
    };
    await paint();
}
