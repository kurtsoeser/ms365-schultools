/**
 * Admin-Backend: Wahl anlegen, Kandidaturen, Zeitfenster, QR, Auszählung.
 * Familie wahl-portal – Calm Phasen/Action-Banner, Ownership Meta, soft Errors.
 */

import { mountResults } from './schulsprecherwahl-results-ui.js';
import { statusLabel, newId } from './schulsprecherwahl-model.js';
import {
    ELECTION_KIND_SCHULSPRECHER,
    ELECTION_KIND_SGA_LEHRER,
    ELECTION_KINDS,
    getElectionKindDef,
    normalizeElectionKind,
    resolveElectionBallotSchema
} from './schulsprecherwahl-kinds.js';
import {
    createElection,
    normalizeElection,
    ELECTION_STATUS,
    listElectionsSoft,
    getElectionAsync,
    upsertElectionAsync,
    deleteElectionAsync,
    openElectionAsync,
    closeAndCountMainAsync,
    openRunoffAsync,
    closeAndCountRunoffAsync,
    seedDemoElectionAsync,
    evaluateElectionAsync,
    publishResultsAsync,
    storageModeLabel,
    isRemoteMode,
    isLocalDemoOnly,
    enableLocalDemoMode,
    wantsElectionDemo
} from './schulsprecherwahl-service.js';
import { softElectionMessage } from './schulsprecherwahl-errors.js';
import {
    electionPhaseSteps,
    electionQuorumInfo,
    nextAdminAction
} from './schulsprecherwahl-phase.js';
import {
    hydrateSchulsprecherwahlStammdaten,
    getSchulsprecherwahlStammdatenMeta,
    loadStudentPoolEffective,
    loadTeacherPoolEffective,
    searchStudentPoolEffective,
    searchTeacherPoolEffective
} from './schulsprecherwahl-stammdaten.js';

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
 * @param {ReturnType<typeof normalizeElection>} election
 * @param {Awaited<ReturnType<typeof evaluateElectionAsync>>} ev
 */
function formatEvalHint(election, ev) {
    const kind = normalizeElectionKind(election.electionKind);
    if (kind === ELECTION_KIND_SGA_LEHRER) {
        const ids = [
            ev.offices.member1Id,
            ev.offices.member2Id,
            ev.offices.member3Id,
            ev.offices.deputy1Id,
            ev.offices.deputy2Id,
            ev.offices.deputy3Id
        ].filter(Boolean);
        if (!ids.length) return '';
        const names = ids.map((id) => {
            const c = election.candidates.find((x) => x.id === id);
            return c ? c.name : id;
        });
        return (
            'SGA Lehrer: ' +
            names.slice(0, 3).join(', ') +
            (names.length > 3 ? ' · STV: ' + names.slice(3).join(', ') : '')
        );
    }
    if (ev.needsRunoff) {
        return (
            'Stichwahl nötig zwischen: ' +
            ev.runoffCandidateIds
                .map((id) => {
                    const c = election.candidates.find((x) => x.id === id);
                    return c ? c.name : id;
                })
                .join(' · ')
        );
    }
    if (ev.offices.schulsprecherId) {
        const c = election.candidates.find((x) => x.id === ev.offices.schulsprecherId);
        return 'Schulsprecher:in: ' + (c ? c.name : ev.offices.schulsprecherId);
    }
    return '';
}

/**
 * @param {ReturnType<typeof normalizeElection>} election
 */
function isElectionAdmin(election) {
    if (!isRemoteMode()) return true;
    return !!(election && /** @type {{ isAdmin?: boolean }} */ (election).isAdmin);
}

/**
 * @param {ReturnType<typeof electionPhaseSteps>} steps
 */
function renderPhaseSteps(steps) {
    return `<ol class="ssw-phases" aria-label="Wahlphasen">
      ${steps
          .map(
              (s) =>
                  `<li class="ssw-phases__step ssw-phases__step--${esc(s.state)}"><span>${esc(
                      s.label
                  )}</span></li>`
          )
          .join('')}
    </ol>`;
}

/**
 * @param {ReturnType<typeof nextAdminAction>} action
 * @param {boolean} canAdmin
 */
function renderActionBanner(action, canAdmin) {
    if (!action || !canAdmin) return '';
    const cta =
        action.ctaAction && action.ctaLabel
            ? `<button type="button" class="btn btn-primary" data-ssw-action="${esc(
                  action.ctaAction
              )}">${esc(action.ctaLabel)}</button>`
            : '';
    return `<section class="ssw-calm-banner ssw-calm-banner--action" role="status">
      <div class="ssw-calm-banner__text">
        <p class="ssw-calm-banner__eyebrow">${esc(action.eyebrow)}</p>
        <p><strong>${esc(action.title)}</strong> – ${esc(action.hint)}</p>
      </div>
      ${cta}
    </section>`;
}

/**
 * @param {ReturnType<typeof electionQuorumInfo>} info
 */
function renderQuorumBanner(info) {
    if (!info) return '';
    const tone =
        info.tone === 'ok' ? 'ssw-calm-banner--ok' : info.tone === 'warn' ? 'ssw-calm-banner--warn' : '';
    return `<section class="ssw-calm-banner ${tone}" role="status">
      <div class="ssw-calm-banner__text">
        <p class="ssw-calm-banner__eyebrow">${esc(info.eyebrow)}</p>
        <p><strong>${esc(info.title)}</strong> – ${esc(info.hint)}</p>
      </div>
    </section>`;
}

/**
 * Deep-Link zur Auswertung (Kundmachung / Teilen).
 * @param {string} electionId
 */
function resultsUrl(electionId) {
    const u = new URL('schulsprecherwahl.html', window.location.href);
    u.searchParams.set('wahl', electionId);
    u.searchParams.set('view', 'results');
    return u.href;
}

/**
 * @param {string} html
 */
function quietEmpty(html) {
    return `<div class="ssw-quiet-empty" role="status">${html}</div>`;
}

/**
 * @param {HTMLElement} host
 */
export function mountSchulsprecherwahlAdmin(host) {
    /** @type {'list'|'edit'|'results'} */
    let view = 'list';
    /** @type {string|null} */
    let editId = null;
    let candidateQuery = '';
    let flash = '';
    /** @type {'ok'|'warn'|''} */
    let flashTone = '';
    let stammdatenReady = false;

    const setFlash = (msg, tone) => {
        flash = msg || '';
        flashTone = tone || '';
    };

    const ensureStammdaten = async () => {
        if (stammdatenReady) return;
        try {
            await hydrateSchulsprecherwahlStammdaten();
        } catch {
            /* soft – lokale Pools bleiben */
        }
        stammdatenReady = true;
    };

    const paint = async () => {
        if (view === 'results' && editId) {
            host.innerHTML = `<p class="ssw-meta">Laden …</p>`;
            let election = null;
            try {
                election = await getElectionAsync(editId);
            } catch (err) {
                setFlash(softElectionMessage(err), 'warn');
                view = 'list';
                return paint();
            }
            if (!election) {
                view = 'list';
                return paint();
            }
            const adminFromApi = !!(/** @type {{ isAdmin?: boolean }} */ (election).isAdmin);
            election = Object.assign(normalizeElection(election), {
                isAdmin: !isRemoteMode() || adminFromApi
            });
            const canAdmin = isElectionAdmin(election);
            const action = nextAdminAction(election);
            const quorum = electionQuorumInfo(election);
            host.innerHTML = `<div class="ssw-wrap">
        <div class="ssw-toolbar">
          <button type="button" class="btn" data-ssw-nav="list"><i class="bi bi-arrow-left"></i> Übersicht</button>
          <button type="button" class="btn" data-ssw-nav="edit"><i class="bi bi-pencil"></i> Verwaltung</button>
        </div>
        ${renderPhaseSteps(electionPhaseSteps(election))}
        ${
            flash
                ? `<p class="ssw-badge${flashTone === 'warn' ? ' ssw-badge--warn' : ''}">${esc(
                      flash
                  )}</p>`
                : ''
        }
        ${renderActionBanner(action, canAdmin)}
        ${renderQuorumBanner(quorum)}
        <div class="ssw-share-bar" role="region" aria-label="Kundmachung">
          <span class="ssw-share-bar__label">${
              election.publishedResults
                  ? 'Ergebnis freigegeben – Link teilen'
                  : 'Noch nicht kundgemacht'
          }</span>
          ${
              canAdmin && !election.publishedResults && election.status === ELECTION_STATUS.finalized
                  ? `<button type="button" class="btn btn-primary" data-ssw-action="publish">Jetzt kundmachen</button>`
                  : ''
          }
          <button type="button" class="btn" data-ssw-action="copy_results">Link kopieren</button>
        </div>
        <div id="sswResultsHost"></div>
      </div>`;
            setFlash('', '');
            bindNav(host);
            bindAdminActions(host, election, canAdmin);
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

    /**
     * 1-Klick-CTAs aus Action-Banner / Kundmachung.
     * @param {ParentNode} root
     * @param {ReturnType<typeof normalizeElection> & { isAdmin?: boolean }} election
     * @param {boolean} canAdmin
     */
    const bindAdminActions = (root, election, canAdmin) => {
        root.querySelectorAll('[data-ssw-action]').forEach((btn) => {
            btn.addEventListener('click', async () => {
                const action = btn.getAttribute('data-ssw-action') || '';
                if (action === 'copy_results') {
                    const link = resultsUrl(election.id);
                    try {
                        await navigator.clipboard.writeText(link);
                        setFlash('Auswertungs-Link kopiert.', 'ok');
                    } catch {
                        setFlash(link, 'ok');
                    }
                    paint();
                    return;
                }
                if (!canAdmin) {
                    setFlash('Keine Berechtigung für diese Aktion.', 'warn');
                    paint();
                    return;
                }
                try {
                    if (action === 'open') {
                        await openElectionAsync(election.id);
                        setFlash('Wahlgang ist offen – QR-Code teilen.', 'ok');
                    } else if (action === 'count') {
                        const next = await closeAndCountMainAsync(election.id);
                        if (normalizeElectionKind(next.electionKind) === ELECTION_KIND_SGA_LEHRER) {
                            setFlash('Ausgezählt – jetzt kundmachen.', 'ok');
                            view = 'results';
                        } else if (next.status === ELECTION_STATUS.finalized) {
                            setFlash('Ausgezählt – jetzt kundmachen.', 'ok');
                            view = 'results';
                        } else {
                            setFlash('Ausgezählt – Stichwahl vorbereiten.', 'ok');
                        }
                    } else if (action === 'runoff') {
                        const now = new Date();
                        const end = new Date(now.getTime() + 2 * 24 * 60 * 60 * 1000);
                        await openRunoffAsync(election.id, {
                            start: now.toISOString(),
                            end: end.toISOString()
                        });
                        setFlash('Stichwahl eröffnet.', 'ok');
                    } else if (action === 'count_runoff') {
                        await closeAndCountRunoffAsync(election.id);
                        setFlash('Stichwahl ausgezählt – jetzt kundmachen.', 'ok');
                        view = 'results';
                    } else if (action === 'publish') {
                        await publishResultsAsync(election.id);
                        setFlash('Ergebnis kundgemacht – Link teilen.', 'ok');
                        view = 'results';
                    } else if (action === 'results') {
                        view = 'results';
                    }
                } catch (err) {
                    setFlash(softElectionMessage(err), 'warn');
                }
                paint();
            });
        });
    };

    /**
     * @param {string} kind
     */
    const createNewElection = async (kind) => {
        await ensureStammdaten();
        const kindDef = getElectionKindDef(kind);
        const pool =
            kindDef.candidatePool === 'teachers'
                ? loadTeacherPoolEffective()
                : loadStudentPoolEffective();
        const e = createElection({
            electionKind: kindDef.id,
            eligibleCount: pool.length,
            schoolYear: ''
        });
        try {
            await upsertElectionAsync(e);
            editId = e.id;
            view = 'edit';
            setFlash(kindDef.label + ' angelegt.', 'ok');
        } catch (err) {
            setFlash(softElectionMessage(err), 'warn');
        }
        paint();
    };

    const paintList = async () => {
        host.innerHTML = `<p class="ssw-meta">Wahlen werden geladen …</p>`;
        await ensureStammdaten();
        const { elections, error } = await listElectionsSoft();
        if (error) setFlash(error.message, 'warn');

        const stamm = getSchulsprecherwahlStammdatenMeta();
        const rows = elections
            .map((e) => {
                const kindDef = getElectionKindDef(e.electionKind);
                const sgaBadge = e.includeSgaDeputies
                    ? ' <span class="ssw-badge ssw-badge--muted">+ SGA-STV</span>'
                    : '';
                const admin = isElectionAdmin(e);
                return `<article class="ssw-card-row">
          <div>
            <h4>${esc(e.title)}</h4>
            <p class="ssw-meta"><span class="ssw-badge ssw-badge--muted">${esc(kindDef.shortLabel)}</span>${sgaBadge}
              ${esc(statusLabel(e))} · ${e.candidates.length} Kandidaturen · ${
                    e.eligibleCount || '–'
                } Wahlberechtigte</p>
            ${renderPhaseSteps(electionPhaseSteps(e))}
          </div>
          <div class="ssw-toolbar">
            <button type="button" class="btn" data-open="${esc(e.id)}">${
                    admin ? 'Verwalten' : 'Ansehen'
                }</button>
            <button type="button" class="btn" data-results="${esc(e.id)}">Auswertung</button>
          </div>
        </article>`;
            })
            .join('');

        const emptyBlock = elections.length
            ? ''
            : quietEmpty(
                  error
                      ? '<strong>Keine Wahlen geladen.</strong> Anmeldung prüfen – lokale Demo mit ?demo=1 möglich.'
                      : '<strong>Noch keine Wahl.</strong> Schulsprecher:innenwahl oder SGA Lehrervertretung anlegen.'
              );

        const listBanner =
            !elections.length && !error
                ? `<section class="ssw-calm-banner ssw-calm-banner--action" role="status">
            <div class="ssw-calm-banner__text">
              <p class="ssw-calm-banner__eyebrow">Was jetzt?</p>
              <p><strong>Erste Wahl anlegen</strong> – Kandidaturen aus Stammdaten, Zeitfenster, dann QR teilen.</p>
            </div>
          </section>`
                : '';

        host.innerHTML = `<div class="ssw-wrap">
      <header class="ssw-hero">
        <div>
          <h2>Schulwahlen</h2>
          <p class="ssw-meta">Schulsprecher:innenwahl und SGA-Lehrervertretung – anonym, mit Zeitfenster, QR-Link und Auswertung.</p>
        </div>
        <span class="ssw-badge ssw-badge--muted">${esc(storageModeLabel())}</span>
      </header>
      ${
          flash
              ? `<p class="ssw-badge${flashTone === 'warn' ? ' ssw-badge--warn' : ''}">${esc(
                    flash
                )}</p>`
              : ''
      }
      ${listBanner}
      <div class="ssw-toolbar">
        <button type="button" class="btn btn-primary" data-new-kind="${ELECTION_KIND_SCHULSPRECHER}"><i class="bi bi-plus-lg"></i> Schulsprecher:innenwahl</button>
        <button type="button" class="btn btn-primary" data-new-kind="${ELECTION_KIND_SGA_LEHRER}"><i class="bi bi-plus-lg"></i> SGA Lehrervertretung</button>
        <button type="button" class="btn" id="sswDemo"><i class="bi bi-stars"></i> Demo-Auswertung</button>
      </div>
      <section class="ssw-panel">
        <h3>Wahlen</h3>
        <div class="ssw-list">${rows || emptyBlock}</div>
      </section>
      <section class="ssw-panel ssw-panel--system">
        <h3>Setup &amp; Speicher</h3>
        <p class="ssw-meta">Stimmen liegen <strong>nicht</strong> im Schul-SharePoint (Anonymität). Meta/Voters/Ballots: Azure Tables bzw. Demo localStorage.</p>
        <p class="ssw-meta">Personenpools: ${esc(String(stamm.studentCount))} Schüler:innen · ${esc(
            String(stamm.teacherCount)
        )} Lehrer:innen · Quelle ${esc(stamm.source)}${
            stamm.error ? ' – ' + esc(stamm.error) : ''
        }</p>
        <p class="ssw-foot-note">${
            isRemoteMode()
                ? 'Produktion: election-api (MS365_ELECTION_API.baseUrl). Site-URL nur für Stammdaten-Hydrate, nie für Stimmzettel.'
                : isLocalDemoOnly()
                  ? 'Demo lokal aktiv (?demo=1 oder Button). Remote-API bleibt unangetastet.'
                  : 'Kein Remote-Backend – lokaler Browser-Store. Für schulweite Anonymität MS365_ELECTION_API.baseUrl setzen.'
        }</p>
      </section>
    </div>`;

        setFlash('', '');
        host.querySelectorAll('[data-new-kind]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const kind = btn.getAttribute('data-new-kind') || ELECTION_KIND_SCHULSPRECHER;
                createNewElection(kind);
            });
        });
        host.querySelector('#sswDemo')?.addEventListener('click', async () => {
            enableLocalDemoMode();
            const e = await seedDemoElectionAsync();
            editId = e.id;
            view = 'results';
            setFlash('Demo-Auswertung geladen (lokal).', 'ok');
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
        await ensureStammdaten();
        let election = null;
        try {
            election = editId ? await getElectionAsync(editId) : null;
        } catch (err) {
            setFlash(softElectionMessage(err), 'warn');
            view = 'list';
            return paint();
        }
        if (!election) {
            view = 'list';
            return paint();
        }
        const adminFromApi = !!(/** @type {{ isAdmin?: boolean }} */ (election).isAdmin);
        election = Object.assign(normalizeElection(election), {
            isAdmin: !isRemoteMode() || adminFromApi
        });
        const canAdmin = isElectionAdmin(election);
        const kindDef = getElectionKindDef(election.electionKind);
        const ballotSchema = resolveElectionBallotSchema(election);
        const isSgaLehrer = kindDef.id === ELECTION_KIND_SGA_LEHRER;
        const isSchulsprecher = kindDef.id === ELECTION_KIND_SCHULSPRECHER;
        const url = voteUrl(election.id);
        const hits =
            kindDef.candidatePool === 'teachers'
                ? searchTeacherPoolEffective(candidateQuery, 50)
                : searchStudentPoolEffective(candidateQuery, 50);
        const poolLabel =
            kindDef.candidatePool === 'teachers' ? 'Lehrer:innen-Pool' : 'Schüler-Pool';
        const searchPlaceholder =
            kindDef.candidatePool === 'teachers'
                ? 'Name, Kürzel oder E-Mail'
                : 'Name, Klasse oder E-Mail';
        const emptyHits =
            kindDef.candidatePool === 'teachers'
                ? 'Keine Treffer. Stammdaten unter Tenant → Lehrer:innen pflegen.'
                : 'Keine Treffer. Stammdaten unter Tenant → Schüler:innen pflegen.';
        const voterHint = isSgaLehrer
            ? 'Lehrer:innen scannen den Code und melden sich mit dem Schul-Konto an.'
            : 'Schüler:innen scannen den Code und melden sich mit dem Schul-Konto an.';
        const action = nextAdminAction(election);
        const quorum = electionQuorumInfo(election);
        const readOnlyAttr = canAdmin ? '' : 'disabled';

        let evalHint = '';
        try {
            if (election.status !== ELECTION_STATUS.draft && election.status !== ELECTION_STATUS.open) {
                const ev = await evaluateElectionAsync(election.id);
                evalHint = formatEvalHint(election, ev);
            }
        } catch {
            /* ignore – Auszählung optional */
        }

        host.innerHTML = `<div class="ssw-wrap">
      <div class="ssw-toolbar">
        <button type="button" class="btn" data-ssw-nav="list"><i class="bi bi-arrow-left"></i> Übersicht</button>
        <button type="button" class="btn" data-ssw-nav="results"><i class="bi bi-bar-chart"></i> Auswertung</button>
        <span class="ssw-badge ssw-badge--muted">${esc(kindDef.shortLabel)}</span>
        <span class="ssw-badge ssw-badge--muted">${esc(statusLabel(election))}</span>
        ${
            election.publishedResults
                ? '<span class="ssw-badge">Kundgemacht</span>'
                : ''
        }
        ${canAdmin ? '' : '<span class="ssw-badge ssw-badge--warn">Nur Lesen</span>'}
      </div>
      ${renderPhaseSteps(electionPhaseSteps(election))}
      ${
          flash
              ? `<p class="ssw-badge${flashTone === 'warn' ? ' ssw-badge--warn' : ''}">${esc(
                    flash
                )}</p>`
              : ''
      }
      ${renderActionBanner(action, canAdmin)}
      ${renderQuorumBanner(quorum)}

      <section class="ssw-panel">
        <h3>Wahl-Einstellungen</h3>
        <p class="ssw-meta">${esc(kindDef.legalHint)}</p>
        <div class="ssw-form-grid">
          <label>Titel<input id="sswTitle" value="${esc(election.title)}" ${readOnlyAttr} /></label>
          <label>Kennzeichnung<input id="sswScope" value="${esc(election.scopeLabel)}" placeholder="Schulweit" ${readOnlyAttr} /></label>
          <label>Wahlbeginn<input id="sswStart" type="datetime-local" value="${esc(toLocalInput(election.window.start))}" ${readOnlyAttr} /></label>
          <label>Wahlende<input id="sswEnd" type="datetime-local" value="${esc(toLocalInput(election.window.end))}" ${readOnlyAttr} /></label>
          <label>Wahlberechtigte (Anzahl)<input id="sswEligible" type="number" min="0" value="${election.eligibleCount || 0}" ${readOnlyAttr} /></label>
          <label>Schuljahr<input id="sswYear" value="${esc(election.schoolYear)}" placeholder="2026/27" ${readOnlyAttr} /></label>
        </div>
        ${
            isSchulsprecher
                ? `<p style="margin-top:.85rem">
          <label class="ssw-check">
            <input type="checkbox" id="sswIncludeSga" ${election.includeSgaDeputies ? 'checked' : ''} ${
                      election.status === ELECTION_STATUS.draft && canAdmin ? '' : 'disabled'
                  } />
            Auch SGA-Schüler-Stellvertreter:innen wählen (3 / 2 / 1 Punkte am selben Stimmzettel)
          </label>
          <span class="ssw-meta" style="display:block;margin-top:.35rem">Nur im Entwurf änderbar. Bei aktivierter Option mind. 6 Kandidaturen und sechs Reihungen (6…1).</span>
        </p>`
                : ''
        }
        ${
            canAdmin
                ? `<p style="margin-top:.85rem" class="ssw-toolbar">
          <button type="button" class="btn btn-primary" id="sswSave">Speichern</button>
          <button type="button" class="btn" id="sswFillEligible">Anzahl aus Stammdaten</button>
        </p>`
                : ''
        }
      </section>

      <section class="ssw-panel">
        <h3>Kandidaturen (${election.candidates.length}) · ${esc(poolLabel)}</h3>
        <div class="ssw-selected" id="sswSelected">
          ${
              election.candidates.length
                  ? election.candidates
                        .map(
                            (c) => `<span class="ssw-chip">${esc(c.name)} <span class="ssw-cand__class">${esc(c.className)}</span>
              ${
                  canAdmin && election.status === ELECTION_STATUS.draft
                      ? `<button type="button" data-remove="${esc(c.id)}" aria-label="Entfernen">×</button>`
                      : ''
              }</span>`
                        )
                        .join('')
                  : `<span class="ssw-meta">Noch keine Kandidaturen – aus dem ${esc(poolLabel)} wählen (mind. ${ballotSchema.rankingCount}).</span>`
          }
        </div>
        ${
            canAdmin && election.status === ELECTION_STATUS.draft
                ? `<div class="ssw-cand-pick" style="margin-top:.75rem">
          <label>Suchen<input id="sswSearch" value="${esc(candidateQuery)}" placeholder="${esc(searchPlaceholder)}" /></label>
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
                    : quietEmpty(esc(emptyHits))
            }
          </div>
        </div>`
                : ''
        }
      </section>

      <section class="ssw-panel">
        <h3>QR-Code &amp; Abstimmungs-Link</h3>
        <div class="ssw-qr">
          <img alt="QR-Code zur Abstimmung" width="180" height="180" src="${esc(qrImgUrl(url))}" />
          <div>
            <p>${esc(voterHint)}</p>
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
        ${
            canAdmin
                ? `<div class="ssw-toolbar">
          <button type="button" class="btn btn-primary" id="sswOpen" ${
              election.status === ELECTION_STATUS.draft ? '' : 'disabled'
          }>Wahlgang eröffnen</button>
          <button type="button" class="btn" id="sswCountMain" ${
              election.status === ELECTION_STATUS.open ? '' : 'disabled'
          }>Auszählen</button>
          ${
              isSgaLehrer
                  ? ''
                  : `<button type="button" class="btn" id="sswOpenRunoff" ${
                        election.status === ELECTION_STATUS.closed ? '' : 'disabled'
                    }>Stichwahl eröffnen</button>
          <button type="button" class="btn" id="sswCountRunoff" ${
              election.status === ELECTION_STATUS.runoff_open ? '' : 'disabled'
          }>Stichwahl auszählen</button>`
          }
          <button type="button" class="btn btn-primary" id="sswPublish" ${
              election.status === ELECTION_STATUS.finalized && !election.publishedResults
                  ? ''
                  : 'disabled'
          }>Kundmachen</button>
          <button type="button" class="btn" id="sswDelete" style="margin-left:auto">Wahl löschen</button>
        </div>`
                : `<p class="ssw-meta">Sie können diese Wahl einsehen, aber nicht bearbeiten (Ownership).</p>`
        }
        <p class="ssw-foot-note">${esc(kindDef.legalHint)}</p>
      </section>
    </div>`;

        setFlash('', '');
        bindNav(host);
        bindAdminActions(host, election, canAdmin);

        if (!canAdmin) return;

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
            const sgaBox = /** @type {HTMLInputElement|null} */ (host.querySelector('#sswIncludeSga'));
            if (sgaBox && cur.status === ELECTION_STATUS.draft) {
                cur.includeSgaDeputies = !!sgaBox.checked;
            }
            return upsertElectionAsync(cur);
        };

        host.querySelector('#sswSave')?.addEventListener('click', async () => {
            try {
                await readFormInto();
                setFlash('Gespeichert.', 'ok');
            } catch (err) {
                setFlash(softElectionMessage(err), 'warn');
            }
            paint();
        });
        host.querySelector('#sswFillEligible')?.addEventListener('click', async () => {
            try {
                const cur = await readFormInto();
                if (!cur) return;
                cur.eligibleCount =
                    kindDef.candidatePool === 'teachers'
                        ? loadTeacherPoolEffective().length
                        : loadStudentPoolEffective().length;
                await upsertElectionAsync(cur);
                setFlash(`Wahlberechtigte: ${cur.eligibleCount}`, 'ok');
            } catch (err) {
                setFlash(softElectionMessage(err), 'warn');
            }
            paint();
        });
        host.querySelector('#sswSearch')?.addEventListener('input', (ev) => {
            candidateQuery = /** @type {HTMLInputElement} */ (ev.target).value;
            paint();
        });
        host.querySelectorAll('[data-add-email]').forEach((btn) => {
            btn.addEventListener('click', async () => {
                try {
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
                } catch (err) {
                    setFlash(softElectionMessage(err), 'warn');
                }
                paint();
            });
        });
        host.querySelectorAll('[data-remove]').forEach((btn) => {
            btn.addEventListener('click', async () => {
                try {
                    const cur = await readFormInto();
                    if (!cur) return;
                    const id = btn.getAttribute('data-remove');
                    cur.candidates = cur.candidates.filter((c) => c.id !== id);
                    await upsertElectionAsync(cur);
                } catch (err) {
                    setFlash(softElectionMessage(err), 'warn');
                }
                paint();
            });
        });
        host.querySelector('#sswCopy')?.addEventListener('click', async () => {
            try {
                await navigator.clipboard.writeText(url);
                setFlash('Link kopiert.', 'ok');
                paint();
            } catch {
                setFlash(url, 'ok');
                paint();
            }
        });
        host.querySelector('#sswOpen')?.addEventListener('click', async () => {
            try {
                await readFormInto();
                await openElectionAsync(election.id);
                setFlash('Wahlgang ist offen – QR-Code teilen.', 'ok');
                paint();
            } catch (err) {
                setFlash(softElectionMessage(err), 'warn');
                paint();
            }
        });
        host.querySelector('#sswCountMain')?.addEventListener('click', async () => {
            try {
                const next = await closeAndCountMainAsync(election.id);
                if (normalizeElectionKind(next.electionKind) === ELECTION_KIND_SGA_LEHRER) {
                    setFlash('Ausgezählt – jetzt kundmachen.', 'ok');
                    view = 'results';
                } else if (next.status === ELECTION_STATUS.finalized) {
                    setFlash('Ausgezählt – jetzt kundmachen.', 'ok');
                    view = 'results';
                } else {
                    setFlash('Ausgezählt – Stichwahl vorbereiten.', 'ok');
                }
                paint();
            } catch (err) {
                setFlash(softElectionMessage(err), 'warn');
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
                setFlash('Stichwahl eröffnet.', 'ok');
                paint();
            } catch (err) {
                setFlash(softElectionMessage(err), 'warn');
                paint();
            }
        });
        host.querySelector('#sswCountRunoff')?.addEventListener('click', async () => {
            try {
                await closeAndCountRunoffAsync(election.id);
                setFlash('Stichwahl ausgezählt – jetzt kundmachen.', 'ok');
                view = 'results';
                paint();
            } catch (err) {
                setFlash(softElectionMessage(err), 'warn');
                paint();
            }
        });
        host.querySelector('#sswPublish')?.addEventListener('click', async () => {
            try {
                await publishResultsAsync(election.id);
                setFlash('Ergebnis kundgemacht – Link teilen.', 'ok');
                view = 'results';
                paint();
            } catch (err) {
                setFlash(softElectionMessage(err), 'warn');
                paint();
            }
        });
        host.querySelector('#sswDelete')?.addEventListener('click', async () => {
            if (!confirm('Diese Wahl inkl. Stimmen wirklich löschen?')) return;
            try {
                await deleteElectionAsync(election.id);
                editId = null;
                view = 'list';
                setFlash('Wahl gelöscht.', 'ok');
            } catch (err) {
                setFlash(softElectionMessage(err), 'warn');
            }
            paint();
        });
    };

    // Deep-Link ?wahl=&view=results|edit
    try {
        const params = new URLSearchParams(window.location.search);
        const deepId = String(params.get('wahl') || params.get('e') || '').trim();
        const deepView = String(params.get('view') || '').trim();
        if (deepId) {
            editId = deepId;
            view = deepView === 'edit' ? 'edit' : deepView === 'results' ? 'results' : 'edit';
        }
    } catch {
        /* ignore */
    }

    // ?demo=1 → lokal seeden, wenn noch keine Demo-Wahl
    if (wantsElectionDemo()) {
        enableLocalDemoMode();
        void seedDemoElectionAsync().then((e) => {
            if (e && !editId) {
                editId = e.id;
                view = 'results';
                setFlash('Demo geladen (?demo=1).', 'ok');
            }
            paint();
        });
        return { refresh: paint };
    }

    paint();
    return { refresh: paint };
}

// re-export for tests / tooling
export { ELECTION_KINDS };
