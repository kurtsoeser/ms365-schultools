/**
 * Admin-Backend: Wahl anlegen, Kandidaturen, Zeitfenster, QR, Auszählung.
 */

import { mountResults } from './schulsprecherwahl-results-ui.js';
import { searchStudentPool, loadStudentPool } from './schulsprecherwahl-students.js';
import { statusLabel, newId } from './schulsprecherwahl-model.js';
import {
    createElection,
    normalizeElection,
    ELECTION_STATUS,
    listElections,
    getElectionAsync,
    upsertElectionAsync,
    deleteElectionAsync,
    openElectionAsync,
    closeAndCountMainAsync,
    openRunoffAsync,
    closeAndCountRunoffAsync,
    seedDemoElectionAsync,
    evaluateElectionAsync,
    storageModeLabel,
    isRemoteMode
} from './schulsprecherwahl-service.js';

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
 * @param {string} iso
 */
function toLocalInput(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return '';
    const pad = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/**
 * @param {string} local
 */
function fromLocalInput(local) {
    if (!local) return null;
    const d = new Date(local);
    return Number.isNaN(d.getTime()) ? null : d.toISOString();
}

/**
 * @param {string} electionId
 */
function voteUrl(electionId) {
    const u = new URL('schulsprecherwahl-stimme.html', window.location.href);
    u.searchParams.set('wahl', electionId);
    return u.href;
}

/**
 * @param {string} url
 */
function qrImgUrl(url) {
    return (
        'https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=' + encodeURIComponent(url)
    );
}

/**
 * @param {HTMLElement} host
 */
export function mountSchulsprecherwahlAdmin(host) {
    /** @type {'list'|'edit'|'results'} */
    let view = 'list';
    /** @type {string|null} */
    let editId = null;
    let studentQuery = '';
    let flash = '';

    const refresh = () => paint();

    const paint = async () => {
        if (view === 'results' && editId) {
            host.innerHTML = `<p class="ssw-meta">Laden …</p>`;
            const election = await getElectionAsync(editId);
            if (!election) {
                view = 'list';
                return paint();
            }
            host.innerHTML = `<div class="ssw-wrap">
        <div class="ssw-toolbar">
          <button type="button" class="btn" data-ssw-nav="list"><i class="bi bi-arrow-left"></i> Übersicht</button>
          <button type="button" class="btn" data-ssw-nav="edit"><i class="bi bi-pencil"></i> Verwaltung</button>
        </div>
        <div id="sswResultsHost"></div>
      </div>`;
            bindNav(host);
            await mountResults(
                /** @type {HTMLElement} */ (host.querySelector('#sswResultsHost')),
                election
            );
            return;
        }

        if (view === 'edit') {
            await paintEdit();
            return;
        }

        await paintList();
    };

    const bindNav = (root) => {
        root.querySelectorAll('[data-ssw-nav]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const n = btn.getAttribute('data-ssw-nav');
                if (n === 'list') {
                    view = 'list';
                    editId = null;
                } else if (n === 'edit') view = 'edit';
                else if (n === 'results') view = 'results';
                paint();
            });
        });
    };

    const paintList = async () => {
        host.innerHTML = `<p class="ssw-meta">Wahlen werden geladen …</p>`;
        let elections = [];
        try {
            elections = await listElections();
        } catch (err) {
            flash = err && err.message ? String(err.message) : String(err);
        }
        const rows = elections
            .map((e) => {
                return `<article class="ssw-card-row">
          <div>
            <h4>${esc(e.title)}</h4>
            <p class="ssw-meta">${esc(statusLabel(e))} · ${e.candidates.length} Kandidaturen · ${
                e.eligibleCount || '–'
            } Wahlberechtigte</p>
          </div>
          <div class="ssw-toolbar">
            <button type="button" class="btn" data-open="${esc(e.id)}">Verwalten</button>
            <button type="button" class="btn" data-results="${esc(e.id)}">Auswertung</button>
          </div>
        </article>`;
            })
            .join('');

        host.innerHTML = `<div class="ssw-wrap">
      <header class="ssw-hero">
        <div>
          <h2>Schulsprecher:innenwahl</h2>
          <p class="ssw-meta">Online-Wahl nach österreichischem Wahlpunktesystem (6 / 5 / 4), mit Zeitfenster, QR-Link und Stichwahl.</p>
        </div>
        <span class="ssw-badge ssw-badge--muted">${esc(storageModeLabel())}</span>
      </header>
      ${flash ? `<p class="ssw-badge">${esc(flash)}</p>` : ''}
      <div class="ssw-toolbar">
        <button type="button" class="btn btn-primary" id="sswNew"><i class="bi bi-plus-lg"></i> Neue Wahl</button>
        ${
            isRemoteMode()
                ? ''
                : `<button type="button" class="btn" id="sswDemo"><i class="bi bi-stars"></i> Demo-Auswertung laden</button>`
        }
      </div>
      <section class="ssw-panel">
        <h3>Wahlen</h3>
        <div class="ssw-list">${rows || '<p class="ssw-empty">Noch keine Wahl angelegt.</p>'}</div>
      </section>
      <section class="ssw-panel">
        <h3>So funktioniert’s</h3>
        <ol style="margin:0;padding-left:1.2rem;line-height:1.55">
          <li>Kandidaturen aus dem Schüler-Pool wählen und Zeitfenster setzen.</li>
          <li>Wahl eröffnen – QR-Code / Link an die Schülerschaft.</li>
          <li>Schüler:innen melden sich mit Schul-Konto an und reihen 3 Personen (6/5/4).</li>
          <li>Auszählen: absolute Mehrheit → fertig; sonst Stichwahl, STV bleiben aus dem 1. Gang.</li>
        </ol>
        <p class="ssw-foot-note">${
            isRemoteMode()
                ? 'Produktion: Stimmen liegen anonym in Azure Table Storage (Wähler-Hash getrennt vom Stimmzettel) – nicht im Schul-SharePoint.'
                : 'Demo: lokaler Browser-Store. Für echte schulweite Anonymität MS365_ELECTION_API.baseUrl setzen (election-api).'
        }</p>
      </section>
    </div>`;

        flash = '';
        host.querySelector('#sswNew')?.addEventListener('click', async () => {
            const pool = loadStudentPool();
            const e = createElection({
                eligibleCount: pool.length,
                schoolYear: ''
            });
            await upsertElectionAsync(e);
            editId = e.id;
            view = 'edit';
            paint();
        });
        host.querySelector('#sswDemo')?.addEventListener('click', async () => {
            const e = await seedDemoElectionAsync();
            editId = e.id;
            view = 'results';
            flash = 'Demo-Auswertung geladen.';
            paint();
        });
        host.querySelectorAll('[data-open]').forEach((btn) => {
            btn.addEventListener('click', () => {
                editId = btn.getAttribute('data-open');
                view = 'edit';
                paint();
            });
        });
        host.querySelectorAll('[data-results]').forEach((btn) => {
            btn.addEventListener('click', () => {
                editId = btn.getAttribute('data-results');
                view = 'results';
                paint();
            });
        });
    };

    const paintEdit = async () => {
        host.innerHTML = `<p class="ssw-meta">Laden …</p>`;
        let election = editId ? await getElectionAsync(editId) : null;
        if (!election) {
            view = 'list';
            return paint();
        }
        election = normalizeElection(election);
        const url = voteUrl(election.id);
        const hits = searchStudentPool(studentQuery, 50);

        let evalHint = '';
        try {
            if (election.status !== ELECTION_STATUS.draft && election.status !== ELECTION_STATUS.open) {
                const ev = await evaluateElectionAsync(election.id);
                if (ev.needsRunoff) {
                    evalHint = `Stichwahl nötig zwischen: ${ev.runoffCandidateIds
                        .map((id) => {
                            const c = election.candidates.find((x) => x.id === id);
                            return c ? c.name : id;
                        })
                        .join(' · ')}`;
                } else if (ev.offices.schulsprecherId) {
                    const c = election.candidates.find((x) => x.id === ev.offices.schulsprecherId);
                    evalHint = `Schulsprecher:in: ${c ? c.name : ev.offices.schulsprecherId}`;
                }
            }
        } catch {
            /* ignore */
        }

        host.innerHTML = `<div class="ssw-wrap">
      <div class="ssw-toolbar">
        <button type="button" class="btn" data-ssw-nav="list"><i class="bi bi-arrow-left"></i> Übersicht</button>
        <button type="button" class="btn" data-ssw-nav="results"><i class="bi bi-bar-chart"></i> Auswertung</button>
        <span class="ssw-badge ssw-badge--muted">${esc(statusLabel(election))}</span>
      </div>
      ${flash ? `<p class="ssw-badge">${esc(flash)}</p>` : ''}

      <section class="ssw-panel">
        <h3>Wahl-Einstellungen</h3>
        <div class="ssw-form-grid">
          <label>Titel<input id="sswTitle" value="${esc(election.title)}" /></label>
          <label>Kennzeichnung<input id="sswScope" value="${esc(election.scopeLabel)}" placeholder="Schulweit" /></label>
          <label>Wahlbeginn<input id="sswStart" type="datetime-local" value="${esc(toLocalInput(election.window.start))}" /></label>
          <label>Wahlende<input id="sswEnd" type="datetime-local" value="${esc(toLocalInput(election.window.end))}" /></label>
          <label>Wahlberechtigte (Anzahl)<input id="sswEligible" type="number" min="0" value="${election.eligibleCount || 0}" /></label>
          <label>Schuljahr<input id="sswYear" value="${esc(election.schoolYear)}" placeholder="2026/27" /></label>
        </div>
        <p style="margin-top:.85rem" class="ssw-toolbar">
          <button type="button" class="btn btn-primary" id="sswSave">Speichern</button>
          <button type="button" class="btn" id="sswFillEligible">Anzahl aus Stammdaten</button>
        </p>
      </section>

      <section class="ssw-panel">
        <h3>Kandidaturen (${election.candidates.length})</h3>
        <div class="ssw-selected" id="sswSelected">
          ${
              election.candidates.length
                  ? election.candidates
                        .map(
                            (c) => `<span class="ssw-chip">${esc(c.name)} <span class="ssw-cand__class">${esc(c.className)}</span>
              <button type="button" data-remove="${esc(c.id)}" aria-label="Entfernen">×</button></span>`
                        )
                        .join('')
                  : '<span class="ssw-meta">Noch keine Kandidaturen – aus dem Schüler-Pool wählen.</span>'
          }
        </div>
        <div class="ssw-cand-pick" style="margin-top:.75rem">
          <label>Schüler:in suchen<input id="sswSearch" value="${esc(studentQuery)}" placeholder="Name, Klasse oder E-Mail" /></label>
          <div class="ssw-cand-pick__list">
            ${
                hits.length
                    ? hits
                          .map((s) => {
                              const disabled = election.candidates.some((c) => c.email && c.email === s.email);
                              return `<button type="button" data-add-email="${esc(s.email)}" data-add-name="${esc(s.name)}" data-add-class="${esc(s.className)}" ${disabled ? 'disabled' : ''}>
                  <strong>${esc(s.name)}</strong> · ${esc(s.className)} <span class="ssw-cand__class">${esc(s.email)}</span>
                </button>`;
                          })
                          .join('')
                    : '<p class="ssw-empty">Keine Treffer. Stammdaten unter Tenant → Schüler:innen pflegen.</p>'
            }
          </div>
        </div>
      </section>

      <section class="ssw-panel">
        <h3>QR-Code &amp; Abstimmungs-Link</h3>
        <div class="ssw-qr">
          <img alt="QR-Code zur Abstimmung" width="180" height="180" src="${esc(qrImgUrl(url))}" />
          <div>
            <p>Schüler:innen scannen den Code und melden sich mit dem Schul-Konto an.</p>
            <code>${esc(url)}</code>
            <p class="ssw-toolbar" style="margin-top:.65rem">
              <button type="button" class="btn" id="sswCopy">Link kopieren</button>
              <a class="btn" href="${esc(url)}" target="_blank" rel="noopener">Stimme öffnen</a>
            </p>
          </div>
        </div>
      </section>

      <section class="ssw-panel">
        <h3>Ablauf steuern</h3>
        ${evalHint ? `<p class="ssw-meta">${esc(evalHint)}</p>` : ''}
        <div class="ssw-toolbar">
          <button type="button" class="btn btn-primary" id="sswOpen" ${
              election.status === ELECTION_STATUS.draft ? '' : 'disabled'
          }>1. Wahlgang eröffnen</button>
          <button type="button" class="btn" id="sswCountMain" ${
              election.status === ELECTION_STATUS.open ? '' : 'disabled'
          }>1. Wahlgang auszählen</button>
          <button type="button" class="btn" id="sswOpenRunoff" ${
              election.status === ELECTION_STATUS.closed ? '' : 'disabled'
          }>Stichwahl eröffnen</button>
          <button type="button" class="btn" id="sswCountRunoff" ${
              election.status === ELECTION_STATUS.runoff_open ? '' : 'disabled'
          }>Stichwahl auszählen</button>
          <button type="button" class="btn" id="sswDelete" style="margin-left:auto">Wahl löschen</button>
        </div>
        <p class="ssw-foot-note">Rechtliche Basis: VO über die Wahl der Schülervertreter – Wahlpunkte 6/5/4; Schulsprecher:in mit absoluter Mehrheit der Erstgereihten, sonst Stichwahl; Stellvertretungen nach Punktezahl des 1. Wahlgangs.</p>
      </section>
    </div>`;

        flash = '';
        bindNav(host);

        const readFormInto = async () => {
            const cur = await getElectionAsync(election.id);
            if (!cur) return null;
            cur.title = /** @type {HTMLInputElement} */ (host.querySelector('#sswTitle')).value;
            cur.scopeLabel = /** @type {HTMLInputElement} */ (host.querySelector('#sswScope')).value;
            cur.schoolYear = /** @type {HTMLInputElement} */ (host.querySelector('#sswYear')).value;
            cur.eligibleCount =
                Number(/** @type {HTMLInputElement} */ (host.querySelector('#sswEligible')).value) || 0;
            cur.window = {
                start: fromLocalInput(
                    /** @type {HTMLInputElement} */ (host.querySelector('#sswStart')).value
                ),
                end: fromLocalInput(/** @type {HTMLInputElement} */ (host.querySelector('#sswEnd')).value)
            };
            return upsertElectionAsync(cur);
        };

        host.querySelector('#sswSave')?.addEventListener('click', async () => {
            await readFormInto();
            flash = 'Gespeichert.';
            paint();
        });
        host.querySelector('#sswFillEligible')?.addEventListener('click', async () => {
            const cur = await readFormInto();
            if (!cur) return;
            cur.eligibleCount = loadStudentPool().length;
            await upsertElectionAsync(cur);
            flash = `Wahlberechtigte: ${cur.eligibleCount}`;
            paint();
        });
        host.querySelector('#sswSearch')?.addEventListener('input', (ev) => {
            studentQuery = /** @type {HTMLInputElement} */ (ev.target).value;
            paint();
        });
        host.querySelectorAll('[data-add-email]').forEach((btn) => {
            btn.addEventListener('click', async () => {
                const cur = await readFormInto();
                if (!cur) return;
                const email = btn.getAttribute('data-add-email') || '';
                if (cur.candidates.some((c) => c.email === email)) return;
                cur.candidates.push({
                    id: newId('cand'),
                    name: btn.getAttribute('data-add-name') || email,
                    className: btn.getAttribute('data-add-class') || '',
                    email,
                    entraId: null
                });
                await upsertElectionAsync(cur);
                paint();
            });
        });
        host.querySelectorAll('[data-remove]').forEach((btn) => {
            btn.addEventListener('click', async () => {
                const cur = await readFormInto();
                if (!cur) return;
                const id = btn.getAttribute('data-remove');
                cur.candidates = cur.candidates.filter((c) => c.id !== id);
                await upsertElectionAsync(cur);
                paint();
            });
        });
        host.querySelector('#sswCopy')?.addEventListener('click', async () => {
            try {
                await navigator.clipboard.writeText(url);
                flash = 'Link kopiert.';
                paint();
            } catch {
                flash = url;
                paint();
            }
        });
        host.querySelector('#sswOpen')?.addEventListener('click', async () => {
            try {
                await readFormInto();
                await openElectionAsync(election.id);
                flash = 'Wahlgang ist offen – QR-Code teilen.';
                paint();
            } catch (err) {
                flash = err && err.message ? String(err.message) : String(err);
                paint();
            }
        });
        host.querySelector('#sswCountMain')?.addEventListener('click', async () => {
            try {
                const next = await closeAndCountMainAsync(election.id);
                flash =
                    next.status === ELECTION_STATUS.finalized
                        ? 'Ausgezählt – im 1. Wahlgang entschieden.'
                        : 'Ausgezählt – Stichwahl vorbereiten.';
                paint();
            } catch (err) {
                flash = err && err.message ? String(err.message) : String(err);
                paint();
            }
        });
        host.querySelector('#sswOpenRunoff')?.addEventListener('click', async () => {
            try {
                const now = new Date();
                const end = new Date(now.getTime() + 2 * 24 * 60 * 60 * 1000);
                await openRunoffAsync(election.id, {
                    start: now.toISOString(),
                    end: end.toISOString()
                });
                flash = 'Stichwahl eröffnet.';
                paint();
            } catch (err) {
                flash = err && err.message ? String(err.message) : String(err);
                paint();
            }
        });
        host.querySelector('#sswCountRunoff')?.addEventListener('click', async () => {
            try {
                await closeAndCountRunoffAsync(election.id);
                flash = 'Stichwahl ausgezählt – Ämter besetzt.';
                view = 'results';
                paint();
            } catch (err) {
                flash = err && err.message ? String(err.message) : String(err);
                paint();
            }
        });
        host.querySelector('#sswDelete')?.addEventListener('click', async () => {
            if (!confirm('Diese Wahl inkl. Stimmen wirklich löschen?')) return;
            await deleteElectionAsync(election.id);
            editId = null;
            view = 'list';
            flash = 'Wahl gelöscht.';
            paint();
        });
    };

    paint();
    return { refresh };
}
