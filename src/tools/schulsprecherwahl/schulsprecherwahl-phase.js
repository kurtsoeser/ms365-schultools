/**
 * Phasen, Quorum, Admin-Nächste-Schritte (wahl-portal) – DOM-frei.
 */

import {
    ELECTION_STATUS,
    normalizeElection,
    resolveElectionBallotSchema
} from './schulsprecherwahl-model.js';
import { ELECTION_KIND_SGA_LEHRER } from './schulsprecherwahl-kinds.js';
import { ROUND_MAIN, ROUND_RUNOFF, turnoutPercent } from './schulsprecherwahl-logic.js';

/** Schulpraxis-Hinweis (keine rechtliche Sperre der Auszählung). */
export const DEFAULT_QUORUM_PERCENT = 50;

/**
 * Phasenleiste für Admin. Stichwahl nur bei Schulsprecher.
 * @param {ReturnType<typeof normalizeElection>} election
 * @returns {{ id: string, label: string, state: 'done'|'current'|'todo' }[]}
 */
export function electionPhaseSteps(election) {
    const e = normalizeElection(election);
    const isSga = e.electionKind === ELECTION_KIND_SGA_LEHRER;
    const status = e.status;
    /** @type {{ id: string, label: string }[]} */
    const defs = isSga
        ? [
              { id: 'draft', label: 'Entwurf' },
              { id: 'open', label: 'Offen' },
              { id: 'finalized', label: 'Ausgezählt' },
              { id: 'publish', label: 'Kundmachung' }
          ]
        : [
              { id: 'draft', label: 'Entwurf' },
              { id: 'open', label: 'Offen' },
              { id: 'counted', label: 'Ausgezählt' },
              { id: 'runoff', label: 'Stichwahl' },
              { id: 'publish', label: 'Kundmachung' }
          ];

    const publishIdx = defs.length - 1;
    /** @type {Record<string, number>} */
    const currentIndexByStatus = isSga
        ? {
              [ELECTION_STATUS.draft]: 0,
              [ELECTION_STATUS.open]: 1,
              [ELECTION_STATUS.closed]: 2,
              [ELECTION_STATUS.finalized]: 2
          }
        : {
              [ELECTION_STATUS.draft]: 0,
              [ELECTION_STATUS.open]: 1,
              [ELECTION_STATUS.closed]: 2,
              [ELECTION_STATUS.runoff_closed]: 2,
              [ELECTION_STATUS.runoff_open]: 3,
              [ELECTION_STATUS.finalized]: e.offices && e.offices.decidedAfter === 'runoff' ? 3 : 2
          };

    let current = currentIndexByStatus[status];
    if (current == null) current = 0;

    // Kundmachung als eigener Schritt nach Finalisierung
    if (status === ELECTION_STATUS.finalized) {
        current = e.publishedResults ? defs.length : publishIdx;
    }

    return defs.map((d, i) => ({
        id: d.id,
        label: d.label,
        state: /** @type {'done'|'current'|'todo'} */ (
            i < current ? 'done' : i === current ? 'current' : 'todo'
        )
    }));
}

/**
 * Beteiligung / Quorum-Hinweis (informativ, blockiert Auszählung nicht).
 * @param {ReturnType<typeof normalizeElection>} election
 * @param {number} [castOverride]
 * @param {number} [thresholdPercent]
 * @returns {{
 *   eligible: number,
 *   cast: number,
 *   percent: number|null,
 *   thresholdPercent: number,
 *   reached: boolean|null,
 *   unknown: boolean,
 *   roundKind: string|null,
 *   eyebrow: string,
 *   title: string,
 *   hint: string,
 *   tone: 'ok'|'warn'|'muted'
 * }|null}
 */
export function electionQuorumInfo(election, castOverride, thresholdPercent = DEFAULT_QUORUM_PERCENT) {
    const e = normalizeElection(election);
    if (e.status === ELECTION_STATUS.draft) return null;

    const threshold = Math.max(1, Number(thresholdPercent) || DEFAULT_QUORUM_PERCENT);
    const main = e.rounds.find((r) => r.kind === ROUND_MAIN) || e.rounds[0];
    const runoff = e.rounds.find((r) => r.kind === ROUND_RUNOFF);
    const activeRound =
        e.status === ELECTION_STATUS.runoff_open || e.status === ELECTION_STATUS.runoff_closed
            ? runoff || main
            : main;
    const eligible = Math.max(
        0,
        Number((activeRound && activeRound.stats && activeRound.stats.eligible) || e.eligibleCount) || 0
    );
    const castFromStats = Number(activeRound && activeRound.stats && activeRound.stats.cast) || 0;
    const cast = castOverride != null ? Number(castOverride) || 0 : castFromStats;
    const percent = turnoutPercent(cast, eligible);
    const unknown = eligible <= 0;
    const reached = unknown || percent == null ? null : percent >= threshold;
    const roundKind = activeRound ? activeRound.kind : null;

    if (unknown) {
        return {
            eligible,
            cast,
            percent,
            thresholdPercent: threshold,
            reached: null,
            unknown: true,
            roundKind,
            eyebrow: 'Beteiligung',
            title: 'Wahlberechtigte noch offen',
            hint: 'Anzahl Wahlberechtigter setzen, damit die Beteiligung geprüft werden kann.',
            tone: 'muted'
        };
    }

    // Während offenem Gang oft noch keine Stats – kein Fehlalarm
    if (e.status === ELECTION_STATUS.open && castOverride == null && cast === 0) {
        return {
            eligible,
            cast,
            percent: null,
            thresholdPercent: threshold,
            reached: null,
            unknown: false,
            roundKind,
            eyebrow: 'Quorum',
            title: 'Beteiligung nach Auszählung',
            hint:
                'Richtwert ≥ ' +
                threshold +
                ' % von ' +
                eligible +
                ' Wahlberechtigten (Schulpraxis, keine Sperre).',
            tone: 'muted'
        };
    }

    if (reached) {
        return {
            eligible,
            cast,
            percent,
            thresholdPercent: threshold,
            reached: true,
            unknown: false,
            roundKind,
            eyebrow: 'Quorum',
            title: 'Beteiligung erreicht (' + percent + ' %)',
            hint:
                cast +
                ' von ' +
                eligible +
                ' Wahlberechtigten (≥ ' +
                threshold +
                ' %). Hinweis zur Schulpraxis – Auszählung bleibt möglich.',
            tone: 'ok'
        };
    }

    return {
        eligible,
        cast,
        percent,
        thresholdPercent: threshold,
        reached: false,
        unknown: false,
        roundKind,
        eyebrow: 'Quorum',
        title: 'Beteiligung unter ' + threshold + ' % (' + (percent == null ? '–' : percent + ' %') + ')',
        hint:
            cast +
            ' von ' +
            eligible +
            ' – Auszählung ist trotzdem möglich; bei Bedarf Wahlfenster verlängern.',
        tone: 'warn'
    };
}

/**
 * Action-Banner „Was jetzt?“ inkl. 1-Klick-CTA.
 * @param {ReturnType<typeof normalizeElection>} election
 * @returns {{
 *   code: string,
 *   eyebrow: string,
 *   title: string,
 *   hint: string,
 *   ctaLabel: string|null,
 *   ctaAction: string|null
 * }|null}
 */
export function nextAdminAction(election) {
    const e = normalizeElection(election);
    const isSga = e.electionKind === ELECTION_KIND_SGA_LEHRER;
    const schema = resolveElectionBallotSchema(e);

    switch (e.status) {
        case ELECTION_STATUS.draft: {
            const need = schema.rankingCount;
            const have = e.candidates.length;
            if (have < need) {
                return {
                    code: 'candidates',
                    eyebrow: 'Was jetzt?',
                    title: 'Kandidaturen ergänzen',
                    hint:
                        'Mindestens ' +
                        need +
                        ' Kandidaturen nötig (aktuell ' +
                        have +
                        '), dann Wahlgang eröffnen.',
                    ctaLabel: null,
                    ctaAction: null
                };
            }
            return {
                code: 'open',
                eyebrow: 'Was jetzt?',
                title: 'Wahlgang eröffnen',
                hint: 'Zeitfenster prüfen, dann eröffnen und QR-/Link an Wahlberechtigte teilen.',
                ctaLabel: 'Wahlgang eröffnen',
                ctaAction: 'open'
            };
        }
        case ELECTION_STATUS.open:
            return {
                code: 'count',
                eyebrow: 'Was jetzt?',
                title: 'Auszählen',
                hint: 'Wenn das Fenster zu ist bzw. alle abgestimmt haben: 1. Wahlgang auszählen.',
                ctaLabel: 'Jetzt auszählen',
                ctaAction: 'count'
            };
        case ELECTION_STATUS.closed:
            if (isSga) {
                return {
                    code: 'publish',
                    eyebrow: 'Was jetzt?',
                    title: 'Ergebnis kundmachen',
                    hint: 'Auswertung prüfen und Kundmachung freigeben / Link teilen.',
                    ctaLabel: 'Kundmachung',
                    ctaAction: 'publish'
                };
            }
            return {
                code: 'runoff',
                eyebrow: 'Was jetzt?',
                title: 'Stichwahl eröffnen',
                hint: 'Kein absolutes Mehr im 1. Gang – Stichwahl-Fenster setzen und eröffnen.',
                ctaLabel: 'Stichwahl eröffnen',
                ctaAction: 'runoff'
            };
        case ELECTION_STATUS.runoff_open:
            return {
                code: 'count_runoff',
                eyebrow: 'Was jetzt?',
                title: 'Stichwahl auszählen',
                hint: 'Nach dem Stichwahl-Fenster auszählen und Ämter besetzen.',
                ctaLabel: 'Stichwahl auszählen',
                ctaAction: 'count_runoff'
            };
        case ELECTION_STATUS.finalized:
            if (!e.publishedResults) {
                return {
                    code: 'publish',
                    eyebrow: 'Kundmachung',
                    title: 'Ergebnis kundmachen',
                    hint: 'Ergebnis freigeben und Link zur Auswertung teilen (ohne Rohstimmen).',
                    ctaLabel: 'Jetzt kundmachen',
                    ctaAction: 'publish'
                };
            }
            return {
                code: 'share',
                eyebrow: 'Kundmachung',
                title: 'Ergebnis teilen',
                hint: 'Auswertung öffnen oder Link kopieren. Optional später Ämter ins Organigramm.',
                ctaLabel: 'Link kopieren',
                ctaAction: 'copy_results'
            };
        default:
            return null;
    }
}

/**
 * Demo nur mit ?demo=1 oder explizitem Button – nie still in Produktiv-API.
 * @param {string} [search]
 */
export function wantsElectionDemo(search) {
    try {
        const q = search != null ? search : typeof window !== 'undefined' ? window.location.search : '';
        return new URLSearchParams(q).get('demo') === '1';
    } catch {
        return false;
    }
}
