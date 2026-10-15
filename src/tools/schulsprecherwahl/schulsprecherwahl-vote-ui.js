/**
 * Schüler-Abstimmungsformular (einfach, QR-Zielseite).
 */

import {
    POINTS_SCHULSPRECHER,
    POINTS_STV1,
    POINTS_STV2,
    ROUND_RUNOFF
} from './schulsprecherwahl-logic.js';
import { votingPhase, normalizeElection } from './schulsprecherwahl-model.js';
import {
    castVoteAsync,
    getElectionAsync,
    hasVotedAsync,
    isRemoteMode
} from './schulsprecherwahl-service.js';
import { findStudentByEmail } from './schulsprecherwahl-students.js';
import { voterKeyFor, hasVoted as localHasVoted } from './schulsprecherwahl-store.js';

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
 * @param {{ electionId: string }} opts
 */
export async function mountVoteApp(host, opts) {
    const electionId = String(opts.electionId || '').trim();
    host.innerHTML = `<p class="ssw-meta">Laden …</p>`;
    let election = await getElectionAsync(electionId);

    if (!election) {
        host.innerHTML = `<div class="ssw-empty"><p>Wahl nicht gefunden. Bitte den QR-Code bzw. Link der Schule verwenden${
            isRemoteMode() ? '' : ' (Demo: Wahl zuerst im Admin-Tool dieses Browsers anlegen)'
        }.</p></div>`;
        return;
    }

    election = normalizeElection(election);
    const account = currentAccount();
    const loggedIn =
        typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();

    if (!loggedIn || !account) {
        host.innerHTML = `<div class="ssw-vote ssw-panel">
      <h2>${esc(election.title)}</h2>
      <p>Bitte mit dem <strong>Schul-Konto</strong> (Microsoft 365) anmelden, um anonym abzustimmen.</p>
      <p class="ssw-meta">Es wird nur geprüft, ob Sie wahlberechtigt sind und noch nicht abgestimmt haben – Ihre Stimme wird getrennt gespeichert.</p>
      <p style="margin-top:1rem"><button type="button" class="btn btn-primary" id="sswLoginBtn"><i class="bi bi-box-arrow-in-right"></i> Anmelden</button></p>
    </div>`;
        host.querySelector('#sswLoginBtn')?.addEventListener('click', () => {
            if (typeof window.ms365AuthLogin === 'function') window.ms365AuthLogin();
            else if (typeof window.ms365AuthLoginRedirect === 'function') window.ms365AuthLoginRedirect();
        });
        return;
    }

    const identity = account.oid || account.email;
    const student = findStudentByEmail(account.email);
    const phase = votingPhase(election);

    if (!phase.canVote || !phase.round) {
        host.innerHTML = `<div class="ssw-vote ssw-panel">
      <h2>${esc(election.title)}</h2>
      <p>${esc(phase.reason || 'Derzeit keine Stimmabgabe möglich.')}</p>
      <p class="ssw-meta">Angemeldet als ${esc(account.name || account.email)}</p>
    </div>`;
        return;
    }

    let already = false;
    if (isRemoteMode()) {
        already = await hasVotedAsync(election.id, identity);
    } else {
        const voterKey = await voterKeyFor(election.id, phase.round.id, identity);
        already = localHasVoted(election.id, phase.round.id, voterKey);
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

    const paint = () => {
        const greet = student
            ? `${student.name} · ${student.className}`
            : account.name || account.email;

        if (phase.roundKind === ROUND_RUNOFF) {
            host.innerHTML = `<div class="ssw-vote">
        <div class="ssw-panel">
          <h2>${esc(election.title)}</h2>
          <p class="ssw-meta">Stichwahl · angemeldet als ${esc(greet)}</p>
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
            const labels = [
                { pts: POINTS_SCHULSPRECHER, label: 'Schulsprecher:in (6 Punkte)' },
                { pts: POINTS_STV1, label: '1. Stellvertretung (5 Punkte)' },
                { pts: POINTS_STV2, label: '2. Stellvertretung (4 Punkte)' }
            ];
            const step = rankings.length;
            const stepMeta = labels[step] || null;
            const taken = new Set(rankings);

            host.innerHTML = `<div class="ssw-vote">
        <div class="ssw-panel">
          <h2>${esc(election.title)}</h2>
          <p class="ssw-meta">1. Wahlgang · angemeldet als ${esc(greet)}</p>
          <p>Reihen Sie drei unterschiedliche Kandidaturen. Das entspricht dem amtlichen Stimmzettel mit Wahlpunkten.</p>
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
                      const pts = [POINTS_SCHULSPRECHER, POINTS_STV1, POINTS_STV2][i];
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
                rankings.length === 3
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
                error = err && err.message ? String(err.message) : String(err);
                busy = false;
                paint();
            }
        });
    };

    paint();
}
