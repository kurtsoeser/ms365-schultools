/**
 * Einheitliche Schicht: lokal (Demo) oder election-api (mandantenweit anonym).
 */

import { electionApiConfigured, apiListElections, apiGetElection, apiUpsertElection, apiDeleteElection, apiElectionMe, apiCastVote, apiListBallots } from './schulsprecherwahl-api.js';
import {
    loadStore,
    saveStore,
    getElection as localGet,
    upsertElection as localUpsert,
    deleteElection as localDelete,
    castVote as localCast,
    hasVoted as localHasVoted,
    voterKeyFor,
    ballotsForRound as localBallots,
    evaluateStoredElection as localEvaluate,
    closeAndCountMain as localCloseMain,
    openRunoff as localOpenRunoff,
    closeAndCountRunoff as localCloseRunoff,
    openElection as localOpen,
    seedDemoElection as localSeed,
    createElection
} from './schulsprecherwahl-store.js';
import { evaluateElection, ROUND_MAIN, ROUND_RUNOFF, normId } from './schulsprecherwahl-logic.js';
import { ELECTION_STATUS, normalizeElection, newId } from './schulsprecherwahl-model.js';

export function isRemoteMode() {
    return electionApiConfigured();
}

export function storageModeLabel() {
    return isRemoteMode()
        ? 'Azure Table Storage (anonym, mandantenweit)'
        : 'Lokaler Demo-Store (dieser Browser)';
}

/**
 * @returns {Promise<ReturnType<typeof normalizeElection>[]>}
 */
export async function listElections() {
    if (!isRemoteMode()) return loadStore().elections;
    const data = await apiListElections();
    return (data.elections || []).map((e) => normalizeElection(e));
}

/**
 * @param {string} electionId
 */
export async function getElectionAsync(electionId) {
    if (!isRemoteMode()) return localGet(electionId);
    const data = await apiGetElection(electionId);
    return data.election ? normalizeElection(data.election) : null;
}

/**
 * @param {ReturnType<typeof normalizeElection>} election
 */
export async function upsertElectionAsync(election) {
    if (!isRemoteMode()) return localUpsert(election);
    const data = await apiUpsertElection(normalizeElection(election));
    return normalizeElection(data.election);
}

/**
 * @param {string} electionId
 */
export async function deleteElectionAsync(electionId) {
    if (!isRemoteMode()) return localDelete(electionId);
    await apiDeleteElection(electionId);
}

/**
 * @param {object} args
 * @param {string} args.electionId
 * @param {string} args.voterIdentity
 * @param {{ rankings?: string[], choice?: string }} args.ballot
 */
export async function castVoteAsync(args) {
    if (!isRemoteMode()) return localCast(args);
    return apiCastVote(args.electionId, args.ballot);
}

/**
 * @param {string} electionId
 * @param {string} voterIdentity
 */
export async function hasVotedAsync(electionId, voterIdentity) {
    if (!isRemoteMode()) {
        const e = localGet(electionId);
        if (!e) return false;
        const { votingPhase } = await import('./schulsprecherwahl-model.js');
        const phase = votingPhase(e);
        if (!phase.round) return false;
        const key = await voterKeyFor(e.id, phase.round.id, voterIdentity);
        return localHasVoted(e.id, phase.round.id, key);
    }
    const me = await apiElectionMe(electionId);
    return !!me.voted;
}

/**
 * @param {string} electionId
 */
export async function electionMeAsync(electionId) {
    if (!isRemoteMode()) {
        const e = localGet(electionId);
        if (!e) return null;
        const { votingPhase } = await import('./schulsprecherwahl-model.js');
        const phase = votingPhase(e);
        return {
            canVote: phase.canVote,
            voted: false,
            reason: phase.reason,
            roundKind: phase.roundKind,
            roundId: phase.round ? phase.round.id : null,
            election: e
        };
    }
    return apiElectionMe(electionId);
}

/**
 * @param {string} electionId
 * @param {string} roundId
 */
export async function ballotsForRoundAsync(electionId, roundId) {
    if (!isRemoteMode()) return localBallots(electionId, roundId);
    const data = await apiListBallots(electionId, roundId);
    return data.ballots || [];
}

/**
 * @param {string} electionId
 */
export async function evaluateElectionAsync(electionId) {
    const election = await getElectionAsync(electionId);
    if (!election) throw new Error('Wahl nicht gefunden.');
    const mainRound = election.rounds.find((r) => r.kind === ROUND_MAIN) || election.rounds[0];
    const runoffRound = election.rounds.find((r) => r.kind === ROUND_RUNOFF);
    const mainBallots = await ballotsForRoundAsync(election.id, mainRound.id);
    const runoffBallots = runoffRound
        ? await ballotsForRoundAsync(election.id, runoffRound.id)
        : [];
    return evaluateElection({
        mainBallots,
        candidateIds: election.candidates.map((c) => c.id),
        runoffBallots,
        runoffCandidateIds: runoffRound ? runoffRound.candidateIds : [],
        lotSchulsprecherId: election.lot && election.lot.schulsprecherId,
        lotStv1Id: election.lot && election.lot.stv1Id,
        lotStv2Id: election.lot && election.lot.stv2Id
    });
}

/**
 * @param {string} electionId
 */
export async function openElectionAsync(electionId) {
    if (!isRemoteMode()) return localOpen(electionId);
    const election = await getElectionAsync(electionId);
    if (!election) throw new Error('Wahl nicht gefunden.');
    if (election.candidates.length < 3) {
        throw new Error('Mindestens drei Kandidaturen nötig (6 / 5 / 4 Punkte).');
    }
    const main = election.rounds.find((r) => r.kind === ROUND_MAIN) || election.rounds[0];
    main.status = 'open';
    main.candidateIds = election.candidates.map((c) => c.id);
    election.status = ELECTION_STATUS.open;
    return upsertElectionAsync(election);
}

/**
 * @param {string} electionId
 */
export async function closeAndCountMainAsync(electionId) {
    if (!isRemoteMode()) return localCloseMain(electionId);
    const election = await getElectionAsync(electionId);
    if (!election) throw new Error('Wahl nicht gefunden.');
    const mainRound = election.rounds.find((r) => r.kind === ROUND_MAIN) || election.rounds[0];
    const evaled = await evaluateElectionAsync(election.id);
    const ballots = await ballotsForRoundAsync(election.id, mainRound.id);

    mainRound.status = 'counted';
    mainRound.countedAt = new Date().toISOString();
    mainRound.stats = {
        cast: ballots.length,
        valid: evaled.main.valid,
        invalid: evaled.main.invalid,
        eligible: election.eligibleCount || 0
    };
    election.offices.stv1Id = evaled.offices.stv1Id;
    election.offices.stv2Id = evaled.offices.stv2Id;

    if (evaled.offices.schulsprecherId && evaled.offices.decidedAfter === 'main') {
        election.offices.schulsprecherId = evaled.offices.schulsprecherId;
        election.offices.decidedAfter = 'main';
        election.status = ELECTION_STATUS.finalized;
        election.publishedResults = true;
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
    return upsertElectionAsync(election);
}

/**
 * @param {string} electionId
 * @param {{ start?: string|null, end?: string|null }} [window]
 */
export async function openRunoffAsync(electionId, window) {
    if (!isRemoteMode()) return localOpenRunoff(electionId, window);
    const election = await getElectionAsync(electionId);
    if (!election) throw new Error('Wahl nicht gefunden.');
    const runoff = election.rounds.find((r) => r.kind === ROUND_RUNOFF);
    if (!runoff || runoff.candidateIds.length < 2) {
        throw new Error('Keine Stichwahl-Kandidaten vorhanden.');
    }
    runoff.status = 'open';
    if (window) {
        election.runoffWindow = { start: window.start || null, end: window.end || null };
        runoff.window = { ...election.runoffWindow };
    } else if (!election.runoffWindow.start) {
        const now = new Date();
        const end = new Date(now.getTime() + 3 * 24 * 60 * 60 * 1000);
        election.runoffWindow = { start: now.toISOString(), end: end.toISOString() };
        runoff.window = { ...election.runoffWindow };
    }
    election.status = ELECTION_STATUS.runoff_open;
    return upsertElectionAsync(election);
}

/**
 * @param {string} electionId
 */
export async function closeAndCountRunoffAsync(electionId) {
    if (!isRemoteMode()) return localCloseRunoff(electionId);
    const election = await getElectionAsync(electionId);
    if (!election) throw new Error('Wahl nicht gefunden.');
    const runoff = election.rounds.find((r) => r.kind === ROUND_RUNOFF);
    if (!runoff) throw new Error('Keine Stichwahl.');
    const evaled = await evaluateElectionAsync(election.id);
    const ballots = await ballotsForRoundAsync(election.id, runoff.id);

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
                ? 'Stichwahl unentschieden – Losentscheid setzen.'
                : 'Schulsprecher:in konnte nicht ermittelt werden.'
        );
    }

    election.offices.schulsprecherId = evaled.offices.schulsprecherId;
    election.offices.stv1Id = evaled.offices.stv1Id;
    election.offices.stv2Id = evaled.offices.stv2Id;
    election.offices.decidedAfter = 'runoff';
    election.status = ELECTION_STATUS.finalized;
    election.publishedResults = true;
    return upsertElectionAsync(election);
}

export async function seedDemoElectionAsync() {
    // Demo bleibt immer lokal (Screenshot-Auswertung ohne Backend)
    return localSeed();
}

export { createElection, ELECTION_STATUS, normalizeElection };

// re-export sync local helpers for results when remote ballots already cached... 
export { localEvaluate, normId, saveStore };
