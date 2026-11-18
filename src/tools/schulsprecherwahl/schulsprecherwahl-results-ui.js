/**
 * Auswertung / Kundmachung – Schulsprecher oder SGA Lehrer.
 * Visuell: Podium, Foto-Hydrate, moderne Stats & Ranking-Karten.
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
import {
    ELECTION_KIND_SGA_LEHRER,
    getElectionKindDef,
    normalizeElectionKind,
    resolveElectionBallotSchema,
    seatsForElection
} from './schulsprecherwahl-kinds.js';
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
 * @param {{ id?: string, name?: string, className?: string, email?: string }|null} cand
 * @param {'xl'|'lg'|'md'|'sm'} size
 */
function avatarHtml(cand, size) {
    if (!cand) {
        return `<div class="ssw-avatar ssw-avatar--${size}" aria-hidden="true"><span>?</span></div>`;
    }
    const photoKey = cand.email || cand.id || '';
    const sizeAttr =
        size === 'xl' ? '120x120' : size === 'lg' ? '96x96' : size === 'md' ? '64x64' : '48x48';
    const sizeClass =
        size === 'xl'
            ? ' ssw-avatar--xl'
            : size === 'lg'
              ? ' ssw-avatar--lg'
              : size === 'md'
                ? ' ssw-avatar--md'
                : ' ssw-avatar--sm';
    return `<div class="ssw-avatar${sizeClass} ms365-user-photo-thumb" data-ms365-user-photo="${esc(
        photoKey
    )}" data-ms365-user-photo-name="${esc(cand.name || '')}" data-ms365-user-photo-size="${sizeAttr}">
      <img alt="" hidden />
      <span class="ms365-user-photo-thumb__initials">${esc(initials(cand.name || ''))}</span>
    </div>`;
}

/**
 * SVG-Ring für Beteiligung.
 * @param {number|null} percent
 */
function turnoutRing(percent) {
    const p = percent == null ? 0 : Math.max(0, Math.min(100, percent));
    const r = 42;
    const c = 2 * Math.PI * r;
    const offset = c - (p / 100) * c;
    const label = percent == null ? '–' : Math.round(percent) + '%';
    return `<div class="ssw-ring" role="img" aria-label="Beteiligung ${esc(label)}">
      <svg viewBox="0 0 100 100" aria-hidden="true">
        <circle class="ssw-ring__track" cx="50" cy="50" r="${r}" />
        <circle class="ssw-ring__value" cx="50" cy="50" r="${r}"
          stroke-dasharray="${c.toFixed(2)}" stroke-dashoffset="${offset.toFixed(2)}" />
      </svg>
      <div class="ssw-ring__label">
        <strong>${esc(label)}</strong>
        <span>Beteiligung</span>
      </div>
    </div>`;
}

/**
 * @param {ReturnType<typeof normalizeElection>} election
 * @param {Map<string, any>} byId
 */
function renderPodium(election, byId) {
    const kind = normalizeElectionKind(election.electionKind);
    const seats = seatsForElection(election);
    const decidedTag =
        election.offices.decidedAfter === 'runoff'
            ? 'In der Stichwahl gewählt'
            : election.offices.decidedAfter === 'main'
              ? 'Im 1. Wahlgang gewählt'
              : election.status === 'finalized'
                ? 'Gewählt'
                : '';
    const stvTag = 'Im 1. Wahlgang bestätigt';

    const cards = seats.map((seat, idx) => {
        const cand = byId.get(election.offices[seat.key] || '') || null;
        const isPrimary =
            kind === ELECTION_KIND_SGA_LEHRER ? idx < 3 : seat.key === 'schulsprecherId';
        const isSgaDeputy = String(seat.key).indexOf('sgaDeputy') === 0;
        const tag = cand ? (isPrimary ? decidedTag : isSgaDeputy ? 'SGA · 3/2/1' : stvTag) : '';
        const place =
            kind === ELECTION_KIND_SGA_LEHRER
                ? idx + 1
                : seat.key === 'schulsprecherId'
                  ? 1
                  : seat.key === 'stv1Id'
                    ? 2
                    : seat.key === 'stv2Id'
                      ? 3
                      : idx + 1;
        const size = isPrimary && place === 1 ? 'xl' : isPrimary ? 'lg' : 'md';
        const placeClass =
            place === 1 ? 'ssw-podium__card--gold' : place === 2 ? 'ssw-podium__card--silver' : place === 3 ? 'ssw-podium__card--bronze' : '';

        return `<article class="ssw-podium__card ${placeClass}${isPrimary && place === 1 ? ' ssw-podium__card--lead' : ''}">
      <div class="ssw-podium__place" aria-hidden="true">${place}</div>
      ${avatarHtml(cand, /** @type {'xl'|'lg'|'md'} */ (size))}
      <p class="ssw-podium__role">${esc(seat.label)}</p>
      <h3 class="ssw-podium__name">${esc(cand ? cand.name : 'Noch offen')}</h3>
      <p class="ssw-podium__meta">${esc(cand ? cand.className || '' : '')}</p>
      ${tag ? `<span class="ssw-result-tag">${esc(tag)}</span>` : ''}
    </article>`;
    });

    // Schulsprecher: Lead in der Mitte (2, 1, 3) wenn genau 3 SV-Ämter
    if (kind !== ELECTION_KIND_SGA_LEHRER && cards.length >= 3 && !election.includeSgaDeputies) {
        return `<div class="ssw-podium ssw-podium--trio">${cards[1]}${cards[0]}${cards[2]}</div>`;
    }
    return `<div class="ssw-podium">${cards.join('')}</div>`;
}

/**
 * @param {object[]} rows
 * @param {{ mode: 'votes'|'points', validHint: string }} opts
 */
function renderRankList(rows, opts) {
    return `<ol class="ssw-rank">
      ${rows
          .map((row, i) => {
              const pct = Math.round((row.barValue / row.barMax) * 100);
              const share =
                  opts.mode === 'votes'
                      ? sharePercent(row.votes, row._valid || 0)
                      : null;
              const valueLabel =
                  opts.mode === 'votes' ? String(row.votes) + ' Stimmen' : String(row.points) + ' Pkt.';
              const sub =
                  opts.mode === 'votes'
                      ? share + ' %'
                      : row._sub || '';
              return `<li class="ssw-rank__row${row.badge ? ' is-winner' : ''}">
          <span class="ssw-rank__idx">${i + 1}</span>
          ${avatarHtml(row, 'md')}
          <div class="ssw-rank__body">
            <div class="ssw-rank__head">
              <div>
                <div class="ssw-rank__name">${esc(row.name)}</div>
                <div class="ssw-rank__class">${esc(row.className || '')}</div>
              </div>
              <div class="ssw-rank__nums">
                <strong>${esc(valueLabel)}</strong>
                ${sub ? `<span>${esc(String(sub))}</span>` : ''}
                ${row.badge ? `<span class="ssw-result-tag">${esc(row.badge)}</span>` : ''}
              </div>
            </div>
            <div class="ssw-bar ssw-bar--lg" aria-hidden="true"><i style="width:${pct}%"></i></div>
          </div>
        </li>`;
          })
          .join('')}
    </ol>`;
}

/**
 * @param {ReturnType<typeof normalizeElection>} election
 * @param {'main'|'runoff'} activeRound
 * @param {Awaited<ReturnType<typeof evaluateElectionAsync>>} evaled
 * @param {number} castCount
 */
export function renderResultsHtml(election, activeRound, evaled, castCount) {
    const e = normalizeElection(election);
    const kindDef = getElectionKindDef(e.electionKind);
    const ballotSchema = resolveElectionBallotSchema(e);
    const isSga = normalizeElectionKind(e.electionKind) === ELECTION_KIND_SGA_LEHRER;
    const byId = new Map(e.candidates.map((c) => [c.id, c]));
    const mainRound = e.rounds.find((r) => r.kind === ROUND_MAIN) || e.rounds[0];
    const runoffRound = e.rounds.find((r) => r.kind === ROUND_RUNOFF);
    const hasRunoff = !!runoffRound && !isSga;
    const showRunoff = hasRunoff && activeRound === 'runoff';

    const statusBadge =
        e.status === 'finalized'
            ? e.offices.decidedAfter === 'runoff'
                ? 'Nach Stichwahl entschieden'
                : e.publishedResults
                  ? 'Kundgemacht'
                  : 'Ausgezählt'
            : statusLabel(e);

    const round = showRunoff ? runoffRound : mainRound;
    const stats = round ? round.stats : { cast: 0, valid: 0, invalid: 0, eligible: e.eligibleCount };
    const turnout = turnoutPercent(stats.cast || castCount, stats.eligible || e.eligibleCount);
    const validCount = stats.valid || (showRunoff && evaled.runoffTally ? evaled.runoffTally.valid : evaled.main.valid) || 0;

    let rankHtml = '';
    if (showRunoff && evaled.runoffTally) {
        const rows = runoffResultRows(evaled.runoffTally, e.candidates, e.offices.schulsprecherId).map(
            (r) => Object.assign(r, { _valid: evaled.runoffTally.valid })
        );
        rankHtml = renderRankList(rows, { mode: 'votes', validHint: '' });
    } else {
        const rows = mainResultRows(evaled.main, e.candidates, e.offices, e).map((row) => {
            const sub = isSga
                ? ''
                : ballotSchema.includeSgaDeputies
                  ? `${row.firstPlace}× Erst · SGA ${row.sgaPoints || 0}`
                  : `${row.firstPlace}× Erstplatzierung`;
            return Object.assign(row, { _sub: sub });
        });
        rankHtml = renderRankList(rows, { mode: 'points', validHint: '' });
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
        ? `Prozentanteile bezogen auf ${validCount} gültige Stimmen in der Stichwahl.`
        : isSga
          ? `Wahlpunkte: ${kindDef.points.join(' · ')}. Top 3 = Lehrervertretung, Platz 4–6 = Stellvertretung.`
          : ballotSchema.includeSgaDeputies
            ? `Wahlpunkte: 6/5/4 für Schulsprecher:in & STV; 3/2/1 für SGA-Stellvertreter:innen.`
            : `Wahlpunkte: 6 · 5 · 4. Erstplatzierungen für die absolute Mehrheit.`;

    const cast = stats.cast || castCount;

    return `<div class="ssw-wrap ssw-results" data-election-id="${esc(e.id)}">
      <header class="ssw-results-hero">
        <div class="ssw-results-hero__glow" aria-hidden="true"></div>
        <div class="ssw-results-hero__text">
          <p class="ssw-results-hero__eyebrow">${esc(kindDef.shortLabel)}${
              ballotSchema.includeSgaDeputies ? ' · + SGA-STV' : ''
          }</p>
          <h2>${esc(e.title)}</h2>
          <p class="ssw-meta">${esc(e.scopeLabel)}${e.schoolYear ? ' · ' + esc(e.schoolYear) : ''} · ${
              e.candidates.length
          } Kandidaturen · ${hasRunoff ? '2 Wahlgänge' : '1 Wahlgang'}</p>
        </div>
        <span class="ssw-badge ssw-badge--hero">${esc(statusBadge)}</span>
      </header>

      <section class="ssw-panel ssw-panel--ceremony" aria-label="Ämterbesetzung">
        <div class="ssw-panel__head">
          <h3>${isSga ? 'Gewählte Vertreter:innen' : 'Ämterbesetzung'}</h3>
          <p class="ssw-meta">Offizielles Ergebnis der Wahl</p>
        </div>
        ${renderPodium(e, byId)}
      </section>

      <section class="ssw-insight" aria-label="Kennzahlen">
        ${turnoutRing(turnout)}
        <div class="ssw-insight__stats">
          <div class="ssw-stat-card">
            <span class="ssw-stat-card__icon" aria-hidden="true"><i class="bi bi-envelope-check"></i></span>
            <strong>${cast}</strong>
            <span>Abgegeben</span>
          </div>
          <div class="ssw-stat-card">
            <span class="ssw-stat-card__icon" aria-hidden="true"><i class="bi bi-check2-circle"></i></span>
            <strong>${validCount}</strong>
            <span>Gültig</span>
          </div>
          <div class="ssw-stat-card">
            <span class="ssw-stat-card__icon" aria-hidden="true"><i class="bi bi-x-circle"></i></span>
            <strong>${stats.invalid || 0}</strong>
            <span>Ungültig</span>
          </div>
          <div class="ssw-stat-card">
            <span class="ssw-stat-card__icon" aria-hidden="true"><i class="bi bi-people"></i></span>
            <strong>${stats.eligible || e.eligibleCount || '–'}</strong>
            <span>Wahlberechtigt</span>
          </div>
        </div>
      </section>

      <section class="ssw-panel ssw-panel--rank">
        <div class="ssw-panel__head ssw-panel__head--row">
          <div>
            <h3>Wahlgänge</h3>
            <p class="ssw-meta">${esc(dateLabel)} · ${
                round && round.status === 'counted' ? 'Ausgezählt' : 'Offen'
            }</p>
          </div>
          <div class="ssw-round-tabs" role="tablist">
            <button type="button" data-ssw-round="main" aria-selected="${!showRunoff}">1. Wahlgang</button>
            ${
                hasRunoff
                    ? `<button type="button" data-ssw-round="runoff" aria-selected="${showRunoff}">Stichwahl</button>`
                    : ''
            }
          </div>
        </div>
        ${rankHtml}
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

    const hydratePhotos = () => {
        try {
            const api = window.ms365UserPhotoThumb;
            if (!api) return;
            if (typeof api.hydrate === 'function') api.hydrate(host);
            else if (typeof api.scan === 'function') api.scan(host);
        } catch {
            /* ignore */
        }
    };

    const paint = async () => {
        host.innerHTML = `<p class="ssw-meta">Auswertung wird geladen …</p>`;
        try {
            const evaled = await evaluateElectionAsync(election.id);
            const active =
                round === 'runoff'
                    ? election.rounds.find((r) => r.kind === ROUND_RUNOFF)
                    : election.rounds.find((r) => r.kind === ROUND_MAIN) || election.rounds[0];
            const castList = active ? await ballotsForRoundAsync(election.id, active.id) : [];
            // Frische Meta (publishedResults kann sich geändert haben)
            const fresh = Object.assign({}, election);
            host.innerHTML = renderResultsHtml(fresh, round, evaled, castList.length);
            host.querySelectorAll('[data-ssw-round]').forEach((btn) => {
                btn.addEventListener('click', () => {
                    round = /** @type {'main'|'runoff'} */ (btn.getAttribute('data-ssw-round') || 'main');
                    paint();
                });
            });
            // Foto-Hydrate nach Paint + kurzer Delay (MSAL/Graph bereit)
            hydratePhotos();
            requestAnimationFrame(hydratePhotos);
            setTimeout(hydratePhotos, 400);
        } catch (err) {
            host.innerHTML = `<p class="ssw-badge ssw-badge--warn">${esc(
                err && err.message ? err.message : String(err)
            )}</p>`;
        }
    };
    await paint();
}
