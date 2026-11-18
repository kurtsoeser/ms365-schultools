/**
 * Persistenz: Wähler-Register und Stimmzettel getrennt (Anonymität im Demo-Store).
 * Produktion: später Azure-Function mit gleichem Vertrag.
 */

import {
    createElection,
    normalizeElection,
    normalizeStore,
    STORAGE_KEY,
    ELECTION_STATUS,
    newId,
    votingPhase
} from './schulsprecherwahl-model.js';
import {
    ROUND_MAIN,
    ROUND_RUNOFF,
    evaluateElection,
    validateMainBallot,
    validateRunoffBallot,
    normId
} from './schulsprecherwahl-logic.js';
import {
    ELECTION_KIND_SGA_LEHRER,
    normalizeElectionKind,
    resolveElectionBallotSchema
} from './schulsprecherwahl-kinds.js';

/**
 * @returns {ReturnType<typeof normalizeStore>}
 */
export function loadStore() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return normalizeStore({ elections: [], voters: [], ballots: [] });
        return normalizeStore(JSON.parse(raw));
    } catch {
        return normalizeStore({ elections: [], voters: [], ballots: [] });
    }
}

/**
 * @param {ReturnType<typeof normalizeStore>} store
 */
export function saveStore(store) {
    const next = normalizeStore(store);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
    return next;
}

/**
 * Stabiler Wähler-Schlüssel ohne Stimminhalt (SHA-256 Hex).
 * @param {string} electionId
 * @param {string} roundId
 * @param {string} voterIdentity oid oder E-Mail
 */
export async function voterKeyFor(electionId, roundId, voterIdentity) {
    const payload = `${normId(electionId)}|${normId(roundId)}|${String(voterIdentity || '').trim().toLowerCase()}`;
    if (typeof crypto !== 'undefined' && crypto.subtle) {
        const buf = new TextEncoder().encode(payload);
        const dig = await crypto.subtle.digest('SHA-256', buf);
        return [...new Uint8Array(dig)].map((b) => b.toString(16).padStart(2, '0')).join('');
    }
    // Fallback (nicht kryptographisch stark, nur Demo)
    let h = 0;
    for (let i = 0; i < payload.length; i++) h = (h * 31 + payload.charCodeAt(i)) >>> 0;
    return 'fb_' + h.toString(16) + '_' + payload.length;
}

/**
 * @param {string} electionId
 */
export function getElection(electionId) {
    const store = loadStore();
    return store.elections.find((e) => e.id === normId(electionId)) || null;
}

/**
 * @param {ReturnType<typeof normalizeElection>} election
 */
export function upsertElection(election) {
    const store = loadStore();
    const next = normalizeElection({ ...election, updatedAt: new Date().toISOString() });
    const idx = store.elections.findIndex((e) => e.id === next.id);
    if (idx >= 0) store.elections[idx] = next;
    else store.elections.unshift(next);
    saveStore(store);
    return next;
}

/**
 * @param {string} electionId
 */
export function deleteElection(electionId) {
    const store = loadStore();
    const id = normId(electionId);
    store.elections = store.elections.filter((e) => e.id !== id);
    store.voters = store.voters.filter((v) => v.electionId !== id);
    store.ballots = store.ballots.filter((b) => b.electionId !== id);
    saveStore(store);
}

/**
 * @param {string} electionId
 * @param {string} roundId
 * @param {string} voterKey
 */
export function hasVoted(electionId, roundId, voterKey) {
    const store = loadStore();
    return store.voters.some(
        (v) =>
            v.electionId === normId(electionId) &&
            v.roundId === normId(roundId) &&
            v.voterKey === String(voterKey || '')
    );
}

/**
 * Stimme abgeben: erst anonymen Stimmzettel, dann Wähler-Vermerk (getrennt).
 * @param {object} args
 * @param {string} args.electionId
 * @param {string} args.voterIdentity
 * @param {{ rankings?: string[], choice?: string }} args.ballot
 */
export async function castVote(args) {
    const election = getElection(args.electionId);
    if (!election) throw new Error('Wahl nicht gefunden.');
    const phase = votingPhase(election);
    if (!phase.canVote || !phase.round) throw new Error(phase.reason || 'Keine Stimmabgabe möglich.');

    const round = phase.round;
    const allow = new Set(
        (round.candidateIds.length ? round.candidateIds : election.candidates.map((c) => c.id)).map(normId)
    );

    if (phase.roundKind === ROUND_MAIN) {
        const check = validateMainBallot(args.ballot, allow, election);
        if (!check.ok) throw new Error(check.reason);
    } else {
        if (normalizeElectionKind(election.electionKind) === ELECTION_KIND_SGA_LEHRER) {
            throw new Error('Für die SGA-Lehrerwahl gibt es keine Stichwahl.');
        }
        const check = validateRunoffBallot(args.ballot, allow);
        if (!check.ok) throw new Error(check.reason);
    }

    const voterKey = await voterKeyFor(election.id, round.id, args.voterIdentity);
    if (hasVoted(election.id, round.id, voterKey)) {
        throw new Error('Sie haben in diesem Wahlgang bereits abgestimmt.');
    }

    const store = loadStore();
    const at = new Date().toISOString();
    /** @type {{ id: string, electionId: string, roundId: string, rankings?: string[], choice?: string, at: string }} */
    const ballotRow = {
        id: newId('ballot'),
        electionId: election.id,
        roundId: round.id,
        at
    };
    if (phase.roundKind === ROUND_MAIN) {
        ballotRow.rankings = /** @type {string[]} */ (
            validateMainBallot(args.ballot, allow, election).rankings
        );
    } else {
        ballotRow.choice = /** @type {string} */ (validateRunoffBallot(args.ballot, allow).choice);
    }

    store.ballots.push(ballotRow);
    store.voters.push({
        electionId: election.id,
        roundId: round.id,
        voterKey,
        at
    });
    saveStore(store);
    return { ok: true, at };
}

/**
 * @param {string} electionId
 * @param {string} roundId
 */
export function ballotsForRound(electionId, roundId) {
    const store = loadStore();
    return store.ballots.filter(
        (b) => b.electionId === normId(electionId) && b.roundId === normId(roundId)
    );
}

/**
 * @param {string} electionId
 */
export function evaluateStoredElection(electionId) {
    const election = getElection(electionId);
    if (!election) throw new Error('Wahl nicht gefunden.');
    const mainRound = election.rounds.find((r) => r.kind === ROUND_MAIN) || election.rounds[0];
    const runoffRound = election.rounds.find((r) => r.kind === ROUND_RUNOFF);
    const mainBallots = ballotsForRound(election.id, mainRound.id);
    const runoffBallots = runoffRound ? ballotsForRound(election.id, runoffRound.id) : [];

    return evaluateElection({
        electionKind: election.electionKind,
        includeSgaDeputies: !!election.includeSgaDeputies,
        mainBallots,
        candidateIds: election.candidates.map((c) => c.id),
        runoffBallots,
        runoffCandidateIds: runoffRound ? runoffRound.candidateIds : [],
        lotSchulsprecherId: election.lot && election.lot.schulsprecherId,
        lotStv1Id: election.lot && election.lot.stv1Id,
        lotStv2Id: election.lot && election.lot.stv2Id,
        lot: election.lot || {}
    });
}

/**
 * @param {ReturnType<typeof normalizeElection>} election
 * @param {ReturnType<typeof evaluateElection>} evaled
 */
function applySgaLehrerOffices(election, evaled) {
    const o = evaled.offices || {};
    election.offices.member1Id = o.member1Id || null;
    election.offices.member2Id = o.member2Id || null;
    election.offices.member3Id = o.member3Id || null;
    election.offices.deputy1Id = o.deputy1Id || null;
    election.offices.deputy2Id = o.deputy2Id || null;
    election.offices.deputy3Id = o.deputy3Id || null;
    election.offices.decidedAfter = 'main';
}

/**
 * 1. Wahlgang schließen / auszählen.
 * @param {string} electionId
 */
export function closeAndCountMain(electionId) {
    const election = getElection(electionId);
    if (!election) throw new Error('Wahl nicht gefunden.');
    const mainRound = election.rounds.find((r) => r.kind === ROUND_MAIN) || election.rounds[0];
    const evaled = evaluateStoredElection(election.id);
    const ballots = ballotsForRound(election.id, mainRound.id);

    mainRound.status = 'counted';
    mainRound.countedAt = new Date().toISOString();
    mainRound.stats = {
        cast: ballots.length,
        valid: evaled.main.valid,
        invalid: evaled.main.invalid,
        eligible: election.eligibleCount || 0
    };

    if (normalizeElectionKind(election.electionKind) === ELECTION_KIND_SGA_LEHRER) {
        applySgaLehrerOffices(election, evaled);
        election.status = ELECTION_STATUS.finalized;
        election.publishedResults = false;
        return upsertElection(election);
    }

    election.offices.stv1Id = evaled.offices.stv1Id;
    election.offices.stv2Id = evaled.offices.stv2Id;
    if (election.includeSgaDeputies) {
        election.offices.sgaDeputy1Id = evaled.offices.sgaDeputy1Id || null;
        election.offices.sgaDeputy2Id = evaled.offices.sgaDeputy2Id || null;
        election.offices.sgaDeputy3Id = evaled.offices.sgaDeputy3Id || null;
    }

    if (evaled.offices.schulsprecherId && evaled.offices.decidedAfter === 'main') {
        election.offices.schulsprecherId = evaled.offices.schulsprecherId;
        election.offices.decidedAfter = 'main';
        election.status = ELECTION_STATUS.finalized;
        election.publishedResults = false;
    } else if (evaled.needsRunoff || evaled.runoffCandidateIds.length >= 2) {
        let runoff = election.rounds.find((r) => r.kind === ROUND_RUNOFF);
        if (!runoff) {
            runoff = {
                id: newId('runoff'),
                kind: ROUND_RUNOFF,
                status: 'pending',
                candidateIds: evaled.runoffCandidateIds.slice(0, 2),
                window: { start: null, end: null },
                countedAt: null,
                stats: { cast: 0, valid: 0, invalid: 0, eligible: election.eligibleCount || 0 }
            };
            election.rounds.push(runoff);
        } else {
            runoff.candidateIds = evaled.runoffCandidateIds.slice(0, 2);
        }
        election.status = ELECTION_STATUS.closed;
        election.offices.schulsprecherId = null;
        election.offices.decidedAfter = null;
    } else {
        election.status = ELECTION_STATUS.closed;
    }

    return upsertElection(election);
}

/**
 * Stichwahl eröffnen.
 * @param {string} electionId
 * @param {{ start?: string|null, end?: string|null }} [window]
 */
export function openRunoff(electionId, window) {
    const election = getElection(electionId);
    if (!election) throw new Error('Wahl nicht gefunden.');
    if (normalizeElectionKind(election.electionKind) === ELECTION_KIND_SGA_LEHRER) {
        throw new Error('Für die SGA-Lehrerwahl gibt es keine Stichwahl.');
    }
    const runoff = election.rounds.find((r) => r.kind === ROUND_RUNOFF);
    if (!runoff || runoff.candidateIds.length < 2) {
        throw new Error('Keine Stichwahl-Kandidaten vorhanden. Zuerst 1. Wahlgang auszählen.');
    }
    runoff.status = 'open';
    if (window) {
        election.runoffWindow = {
            start: window.start || null,
            end: window.end || null
        };
        runoff.window = { ...election.runoffWindow };
    } else if (!election.runoffWindow.start) {
        const now = new Date();
        const end = new Date(now.getTime() + 3 * 24 * 60 * 60 * 1000);
        election.runoffWindow = { start: now.toISOString(), end: end.toISOString() };
        runoff.window = { ...election.runoffWindow };
    }
    election.status = ELECTION_STATUS.runoff_open;
    return upsertElection(election);
}

/**
 * Stichwahl auszählen und finalisieren.
 * @param {string} electionId
 */
export function closeAndCountRunoff(electionId) {
    const election = getElection(electionId);
    if (!election) throw new Error('Wahl nicht gefunden.');
    const runoff = election.rounds.find((r) => r.kind === ROUND_RUNOFF);
    if (!runoff) throw new Error('Keine Stichwahl.');
    const evaled = evaluateStoredElection(election.id);
    const ballots = ballotsForRound(election.id, runoff.id);

    runoff.status = 'counted';
    runoff.countedAt = new Date().toISOString();
    runoff.stats = {
        cast: ballots.length,
        valid: evaled.runoffTally ? evaled.runoffTally.valid : 0,
        invalid: evaled.runoffTally ? evaled.runoffTally.invalid : 0,
        eligible: election.eligibleCount || 0
    };

    if (!evaled.offices.schulsprecherId) {
        throw new Error(
            evaled.offices.tieNeedsLotSchulsprecher
                ? 'Stichwahl unentschieden – Losentscheid in den Einstellungen setzen.'
                : 'Schulsprecher:in konnte nicht ermittelt werden.'
        );
    }

    election.offices.schulsprecherId = evaled.offices.schulsprecherId;
    election.offices.stv1Id = evaled.offices.stv1Id;
    election.offices.stv2Id = evaled.offices.stv2Id;
    if (election.includeSgaDeputies) {
        election.offices.sgaDeputy1Id = evaled.offices.sgaDeputy1Id || null;
        election.offices.sgaDeputy2Id = evaled.offices.sgaDeputy2Id || null;
        election.offices.sgaDeputy3Id = evaled.offices.sgaDeputy3Id || null;
    }
    election.offices.decidedAfter = 'runoff';
    election.status = ELECTION_STATUS.finalized;
    election.publishedResults = false;
    return upsertElection(election);
}

/**
 * Kundmachung: Ergebnis freigeben (ohne Rohstimmen zu teilen).
 * @param {string} electionId
 */
export function publishResults(electionId) {
    const election = getElection(electionId);
    if (!election) throw new Error('Wahl nicht gefunden.');
    if (election.status !== ELECTION_STATUS.finalized) {
        throw new Error('Erst auszählen – dann kundmachen.');
    }
    election.publishedResults = true;
    return upsertElection(election);
}

/**
 * Wahlgang öffnen.
 * @param {string} electionId
 */
export function openElection(electionId) {
    const election = getElection(electionId);
    if (!election) throw new Error('Wahl nicht gefunden.');
    const schema = resolveElectionBallotSchema(election);
    if (election.candidates.length < schema.rankingCount) {
        throw new Error(
            'Mindestens ' +
                schema.rankingCount +
                ' Kandidaturen nötig (' +
                schema.points.join(' / ') +
                ' Punkte).'
        );
    }
    const main = election.rounds.find((r) => r.kind === ROUND_MAIN) || election.rounds[0];
    main.status = 'open';
    main.candidateIds = election.candidates.map((c) => c.id);
    election.status = ELECTION_STATUS.open;
    return upsertElection(election);
}

/**
 * Demo-Wahl mit fiktiven Personen (keine Realnamen).
 */
export function seedDemoElection() {
    const election = createElection({
        title: 'Schulsprecher:innenwahl',
        scopeLabel: 'Schulweit',
        schoolYear: '2026/27',
        status: ELECTION_STATUS.finalized,
        eligibleCount: 673,
        publishedResults: true,
        candidates: [
            {
                id: 'demo_lena',
                name: 'Lena Bergmann',
                className: '5AHW',
                email: 'lena.bergmann.demo@schule.example'
            },
            {
                id: 'demo_felix',
                name: 'Felix Ortner',
                className: '4AIT',
                email: 'felix.ortner.demo@schule.example'
            },
            {
                id: 'demo_mira',
                name: 'Mira Brandstetter',
                className: '3BHM',
                email: 'mira.brandstetter.demo@schule.example'
            },
            {
                id: 'demo_noah',
                name: 'Noah Weiss',
                className: '2CHW',
                email: 'noah.weiss.demo@schule.example'
            },
            {
                id: 'demo_ida',
                name: 'Ida Sommer',
                className: '5BIT',
                email: 'ida.sommer.demo@schule.example'
            }
        ],
        window: {
            start: '2026-10-01T07:00:00.000Z',
            end: '2026-10-08T16:00:00.000Z'
        },
        runoffWindow: {
            start: '2026-10-09T07:00:00.000Z',
            end: '2026-10-09T16:00:00.000Z'
        },
        offices: {
            schulsprecherId: 'demo_lena',
            stv1Id: 'demo_felix',
            stv2Id: 'demo_mira',
            decidedAfter: 'runoff'
        },
        rounds: [
            {
                id: 'main_demo',
                kind: ROUND_MAIN,
                status: 'counted',
                candidateIds: ['demo_lena', 'demo_felix', 'demo_mira', 'demo_noah', 'demo_ida'],
                countedAt: '2026-10-08T16:30:00.000Z',
                stats: { cast: 612, valid: 610, invalid: 2, eligible: 673 }
            },
            {
                id: 'runoff_demo',
                kind: ROUND_RUNOFF,
                status: 'counted',
                candidateIds: ['demo_lena', 'demo_felix'],
                countedAt: '2026-10-09T16:30:00.000Z',
                stats: { cast: 587, valid: 587, invalid: 0, eligible: 673 },
                window: {
                    start: '2026-10-09T07:00:00.000Z',
                    end: '2026-10-09T16:00:00.000Z'
                }
            }
        ]
    });

    const store = loadStore();
    // bestehende Demo ersetzen
    store.elections = store.elections.filter((e) => !String(e.id).startsWith('wahl_demo'));
    election.id = 'wahl_demo_screenshot';
    store.elections.unshift(election);

    // synthetische Stimmzettel für Auszählung / Balken
    store.ballots = store.ballots.filter((b) => b.electionId !== election.id);
    store.voters = store.voters.filter((v) => v.electionId !== election.id);

    const mainBallots = [];
    // Lena oft vorne, aber keine absolute Mehrheit → Stichwahl; STV-Punkte für Felix/Mira
    for (let i = 0; i < 280; i++) mainBallots.push({ rankings: ['demo_lena', 'demo_felix', 'demo_mira'] });
    for (let i = 0; i < 250; i++) mainBallots.push({ rankings: ['demo_felix', 'demo_lena', 'demo_mira'] });
    for (let i = 0; i < 50; i++) mainBallots.push({ rankings: ['demo_mira', 'demo_lena', 'demo_felix'] });
    for (let i = 0; i < 20; i++) mainBallots.push({ rankings: ['demo_noah', 'demo_lena', 'demo_mira'] });
    for (let i = 0; i < 10; i++) mainBallots.push({ rankings: ['demo_ida', 'demo_mira', 'demo_felix'] });

    for (const rankings of mainBallots) {
        store.ballots.push({
            id: newId('ballot'),
            electionId: election.id,
            roundId: 'main_demo',
            rankings,
            at: '2026-10-08T12:00:00.000Z'
        });
    }

    for (let i = 0; i < 373; i++) {
        store.ballots.push({
            id: newId('ballot'),
            electionId: election.id,
            roundId: 'runoff_demo',
            choice: 'demo_lena',
            at: '2026-10-09T12:00:00.000Z'
        });
    }
    for (let i = 0; i < 214; i++) {
        store.ballots.push({
            id: newId('ballot'),
            electionId: election.id,
            roundId: 'runoff_demo',
            choice: 'demo_felix',
            at: '2026-10-09T12:00:00.000Z'
        });
    }

    saveStore(store);
    return election;
}

export { ELECTION_STATUS, createElection, loadStore as loadElectionStore };
