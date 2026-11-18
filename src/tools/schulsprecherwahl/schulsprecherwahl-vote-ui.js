/**
 * Abstimmungsformular (QR-Zielseite) – Schulsprecher oder SGA Lehrer.
 * Schlank, Trust/Privacy, soft Auth-Fehler – kein IT-Setup.
 */

import { ROUND_RUNOFF } from './schulsprecherwahl-logic.js';
import { votingPhase, normalizeElection, statusLabel } from './schulsprecherwahl-model.js';
import {
    getElectionKindDef,
    ELECTION_KIND_SGA_LEHRER,
    rankingStepLabelsForElection,
    resolveElectionBallotSchema
} from './schulsprecherwahl-kinds.js';
import {
    castVoteAsync,
    getElectionAsync,
    hasVotedAsync,
    isRemoteMode
} from './schulsprecherwahl-service.js';
import {
    findStudentByEmailEffective,
    findTeacherByEmailEffective
} from './schulsprecherwahl-stammdaten.js';
import { voterKeyFor, hasVoted as localHasVoted } from './schulsprecherwahl-store.js';
import { softElectionMessage, classifyElectionError } from './schulsprecherwahl-errors.js';

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
 * @returns {{ email: string, name: string, oid: string }|null}
 */
function currentAccount() {
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const a = window.ms365AuthGetAccountInfo();
            if (a) {
                return {
                    email: String(a.upn || a.username || '')
                        .trim()
                        .toLowerCase(),
                    name: String(a.name || ''),
                    oid: String(a.oid || '')
                };
            }
        }
    } catch {
        /* ignore */
    }
    return null;
}

/**
 * @param {HTMLElement} host
 * @param {string} title
 * @param {string} body
 * @param {{ login?: boolean }} [opts]
 */
function renderSoftState(host, title, body, opts) {
    host.innerHTML = `<div class="ssw-vote ssw-panel">
      <h2>${esc(title)}</h2>
      <p>${esc(body)}</p>
      ${
          opts && opts.login
              ? `<p style="margin-top:1rem"><button type="button" class="btn btn-primary" id="sswLoginBtn"><i class="bi bi-box-arrow-in-right"></i> Anmelden</button></p>`
              : ''
      }
    </div>`;
    host.querySelector('#sswLoginBtn')?.addEventListener('click', () => {
        if (typeof window.ms365AuthLogin === 'function') window.ms365AuthLogin();
        else if (typeof window.ms365AuthLoginRedirect === 'function') window.ms365AuthLoginRedirect();
    });
}

/**
 * @param {HTMLElement} host
 * @param {{ electionId: string }} opts
 */
export async function mountVoteApp(host, opts) {
    const electionId = String(opts.electionId || '').trim();
    host.innerHTML = `<p class="ssw-meta">Laden …</p>`;

    if (!electionId) {
        renderSoftState(
            host,
            'Kein Wahl-Link',
            'Bitte den QR-Code bzw. Link Ihrer Schule verwenden (?wahl=…).'
        );
        return;
    }

    let election = null;
    try {
        election = await getElectionAsync(electionId);
    } catch (err) {
        const soft = classifyElectionError(err);
        renderSoftState(
            host,
            soft.code === 'auth_required' ? 'Anmeldung nötig' : 'Wahl nicht geladen',
            soft.message,
            soft.code === 'auth_required' ? { login: true } : undefined
        );
        return;
    }

    if (!election) {
        host.innerHTML = `<div class="ssw-empty"><p>Wahl nicht gefunden. Bitte den QR-Code bzw. Link der Schule verwenden${
            isRemoteMode() ? '' : ' (Demo: Wahl zuerst im Admin-Tool dieses Browsers anlegen)'
        }.</p></div>`;
        return;
    }

    election = normalizeElection(election);
    const kindDef = getElectionKindDef(election.electionKind);
    const ballotSchema = resolveElectionBallotSchema(election);
    const account = currentAccount();
    const loggedIn =
        typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
    const phaseBadge = `<span class="ssw-badge ssw-badge--muted">${esc(statusLabel(election))}</span>`;

    if (!loggedIn || !account) {
        host.innerHTML = `<div class="ssw-vote ssw-panel">
      <h2>${esc(election.title)}</h2>
      <p class="ssw-meta">${phaseBadge} · ${esc(kindDef.shortLabel)}</p>
      <p>Bitte mit dem <strong>Schul-Konto</strong> (Microsoft 365) anmelden, um anonym abzustimmen.</p>
      <p class="ssw-meta">Es wird nur geprüft, ob Sie noch nicht abgestimmt haben – Ihre Stimme wird getrennt gespeichert.</p>
      <p style="margin-top:1rem"><button type="button" class="btn btn-primary" id="sswLoginBtn"><i class="bi bi-box-arrow-in-right"></i> Anmelden</button></p>
    </div>`;
        host.querySelector('#sswLoginBtn')?.addEventListener('click', () => {
            if (typeof window.ms365AuthLogin === 'function') window.ms365AuthLogin();
            else if (typeof window.ms365AuthLoginRedirect === 'function') window.ms365AuthLoginRedirect();
        });
        return;
    }

    const identity = account.oid || account.email;
    const person =
        kindDef.candidatePool === 'teachers'
            ? findTeacherByEmailEffective(account.email)
            : findStudentByEmailEffective(account.email);
    const phase = votingPhase(election);

    if (!phase.canVote || !phase.round) {
        host.innerHTML = `<div class="ssw-vote ssw-panel">
      <h2>${esc(election.title)}</h2>
      <p class="ssw-meta">${phaseBadge}</p>
      <p>${esc(phase.reason || 'Derzeit keine Stimmabgabe möglich.')}</p>
      <p class="ssw-meta">Angemeldet als ${esc(account.name || account.email)}</p>
    </div>`;
        return;
    }

    let already = false;
    try {
        if (isRemoteMode()) {
            already = await hasVotedAsync(election.id, identity);
        } else {
            const voterKey = await voterKeyFor(election.id, phase.round.id, identity);
            already = localHasVoted(election.id, phase.round.id, voterKey);
        }
    } catch (err) {
        renderSoftState(host, election.title, softElectionMessage(err), {
            login: classifyElectionError(err).code === 'auth_required'
        });
        return;
    }

    if (already) {
        host.innerHTML = `<div class="ssw-vote ssw-panel">
      <h2>${esc(election.title)}</h2>
      <p><span class="ssw-badge">✓ Stimme gespeichert</span></p>
      <p>Sie haben in diesem Wahlgang bereits abgestimmt. Vielen Dank!</p>
    </div>`;
        return;
    }

    const candIds = phase.round.candidateIds.length
        ? phase.round.candidateIds
        : election.candidates.map((c) => c.id);
    const candidates = candIds
        .map((id) => election.candidates.find((c) => c.id === id))
        .filter(Boolean);

    /** @type {string[]} */
    let rankings = [];
    /** @type {string|null} */
    let runoffChoice = null;
    let busy = false;
    let error = '';
    const needRankings = ballotSchema.rankingCount;
    const stepLabels = rankingStepLabelsForElection(election);

    const paint = () => {
        const greet = person
            ? `${person.name}${person.className ? ' · ' + person.className : ''}`
            : account.name || account.email;

        if (phase.roundKind === ROUND_RUNOFF) {
            host.innerHTML = `<div class="ssw-vote">
        <div class="ssw-panel">
          <h2>${esc(election.title)}</h2>
          <p class="ssw-meta">${phaseBadge} · Stichwahl · angemeldet als ${esc(greet)}</p>
          <p>Bitte <strong>eine Person</strong> als Schulsprecher:in wählen.</p>
          <div class="ssw-vote-step">
            ${candidates
                .map((c) => {
                    const sel = runoffChoice === c.id;
                    return `<button type="button" class="ssw-vote-option${sel ? ' is-selected' : ''}" data-choice="${esc(c.id)}" aria-pressed="${sel}">
                <div class="ssw-avatar ssw-avatar--sm">${esc(initials(c.name))}</div>
                <div><div class="ssw-cand__name">${esc(c.name)}</div>
                <div class="ssw-cand__class">${esc(c.className)}</div></div>
              </button>`;
                })
                .join('')}
          </div>
          ${error ? `<p class="ssw-badge ssw-badge--warn">${esc(error)}</p>` : ''}
          <p style="margin-top:1rem">
            <button type="button" class="btn btn-primary" id="sswSubmit" ${!runoffChoice || busy ? 'disabled' : ''}>Stimme abgeben</button>
          </p>
        </div>
      </div>`;
        } else {
            const step = rankings.length;
            const stepMeta = stepLabels[step] || null;
            const taken = new Set(rankings);
            const intro =
                kindDef.id === ELECTION_KIND_SGA_LEHRER
                    ? 'Reihen Sie sechs unterschiedliche Kandidaturen (6 bis 1 Punkte) – Top 3 werden Mitglieder, Platz 4–6 Stellvertretung.'
                    : ballotSchema.includeSgaDeputies
                      ? 'Reihen Sie sechs unterschiedliche Kandidaturen: 6/5/4 für Schulsprecher:in und Stellvertretungen, 3/2/1 für SGA-Stellvertreter:innen.'
                      : 'Reihen Sie drei unterschiedliche Kandidaturen. Das entspricht dem amtlichen Stimmzettel mit Wahlpunkten.';

            host.innerHTML = `<div class="ssw-vote">
        <div class="ssw-panel">
          <h2>${esc(election.title)}</h2>
          <p class="ssw-meta">${phaseBadge} · ${esc(kindDef.shortLabel)} · angemeldet als ${esc(greet)}</p>
          <p>${esc(intro)}</p>
          ${
              stepMeta
                  ? `<p><span class="ssw-points">${stepMeta.pts}</span> &nbsp;<strong>${esc(stepMeta.label)}</strong></p>
            <div class="ssw-vote-step">
              ${candidates
                  .map((c) => {
                      const disabled = taken.has(c.id);
                      return `<button type="button" class="ssw-vote-option" data-pick="${esc(c.id)}" ${disabled ? 'disabled' : ''}>
                  <div class="ssw-avatar ssw-avatar--sm">${esc(initials(c.name))}</div>
                  <div><div class="ssw-cand__name">${esc(c.name)}</div>
                  <div class="ssw-cand__class">${esc(c.className)}</div></div>
                </button>`;
                  })
                  .join('')}
            </div>`
                  : `<div class="ssw-rank-summary">
              ${rankings
                  .map((id, i) => {
                      const c = candidates.find((x) => x.id === id);
                      const pts = ballotSchema.points[i];
                      return `<div><span class="ssw-points">${pts}</span> ${esc(c ? c.name : id)} <span class="ssw-cand__class">${esc(c ? c.className : '')}</span></div>`;
                  })
                  .join('')}
            </div>
            <p class="ssw-meta">Ihre Stimme wird anonym gespeichert. Die Anmeldung dient nur dem Einmal-Schutz.</p>`
          }
          ${error ? `<p class="ssw-badge ssw-badge--warn">${esc(error)}</p>` : ''}
          <p style="margin-top:1rem;display:flex;gap:.5rem;flex-wrap:wrap">
            ${
                rankings.length
                    ? `<button type="button" class="btn" id="sswBack">Zurück</button>`
                    : ''
            }
            ${
                rankings.length === needRankings
                    ? `<button type="button" class="btn btn-primary" id="sswSubmit" ${busy ? 'disabled' : ''}>Stimme abgeben</button>`
                    : ''
            }
          </p>
        </div>
      </div>`;
        }

        host.querySelectorAll('[data-pick]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-pick');
                if (!id || rankings.includes(id)) return;
                rankings = [...rankings, id];
                error = '';
                paint();
            });
        });
        host.querySelectorAll('[data-choice]').forEach((btn) => {
            btn.addEventListener('click', () => {
                runoffChoice = btn.getAttribute('data-choice');
                error = '';
                paint();
            });
        });
        host.querySelector('#sswBack')?.addEventListener('click', () => {
            rankings = rankings.slice(0, -1);
            error = '';
            paint();
        });
        host.querySelector('#sswSubmit')?.addEventListener('click', async () => {
            if (busy) return;
            busy = true;
            error = '';
            paint();
            try {
                const ballot =
                    phase.roundKind === ROUND_RUNOFF
                        ? { choice: runoffChoice || undefined }
                        : { rankings };
                await castVoteAsync({
                    electionId: election.id,
                    voterIdentity: identity,
                    ballot
                });
                host.innerHTML = `<div class="ssw-vote ssw-panel">
          <h2>Danke!</h2>
          <p><span class="ssw-badge">✓ Stimme anonym gespeichert</span></p>
          <p>Ihre Stimme wurde gezählt. Sie können dieses Fenster schließen.</p>
        </div>`;
            } catch (err) {
                error = softElectionMessage(err);
                busy = false;
                paint();
            }
        });
    };

    paint();
}
