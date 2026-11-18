/**
 * Datenmodell & Normalisierung für Schulwahlen (Schulsprecher / SGA Lehrer).
 */

import { ROUND_MAIN, ROUND_RUNOFF, normId } from './schulsprecherwahl-logic.js';
import {
    ELECTION_KIND_SCHULSPRECHER,
    ELECTION_KIND_SGA_LEHRER,
    emptyLotForKind,
    emptyOfficesForKind,
    electionIncludesSgaDeputies,
    getElectionKindDef,
    normalizeElectionKind,
    resolveElectionBallotSchema
} from './schulsprecherwahl-kinds.js';

export const STORAGE_KEY = 'ms365-schulsprecherwahl-v1';
export const ELECTION_STATUS = {
    draft: 'draft',
    open: 'open',
    closed: 'closed',
    runoff_open: 'runoff_open',
    runoff_closed: 'runoff_closed',
    finalized: 'finalized'
};

export {
    ELECTION_KIND_SCHULSPRECHER,
    ELECTION_KIND_SGA_LEHRER,
    normalizeElectionKind,
    getElectionKindDef,
    electionIncludesSgaDeputies,
    resolveElectionBallotSchema
};

/**
 * @returns {string}
 */
export function newId(prefix = 'id') {
    try {
        if (typeof crypto !== 'undefined' && crypto.randomUUID) {
            return prefix + '_' + crypto.randomUUID().replace(/-/g, '').slice(0, 12);
        }
    } catch {
        /* ignore */
    }
    return prefix + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

/**
 * @param {unknown} raw
 */
export function normalizeCandidate(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        id: normId(o.id) || newId('cand'),
        name: String(o.name || '').trim(),
        className: String(o.className || o.klasse || '').trim(),
        email: String(o.email || '').trim().toLowerCase(),
        entraId: String(o.entraId || o.oid || '').trim() || null
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeWindow(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        start: o.start ? String(o.start) : null,
        end: o.end ? String(o.end) : null
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeRound(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const kind = o.kind === ROUND_RUNOFF ? ROUND_RUNOFF : ROUND_MAIN;
    return {
        id: normId(o.id) || newId(kind === ROUND_RUNOFF ? 'runoff' : 'main'),
        kind,
        status: o.status === 'counted' ? 'counted' : o.status === 'open' ? 'open' : 'pending',
        candidateIds: Array.isArray(o.candidateIds) ? o.candidateIds.map(normId).filter(Boolean) : [],
        window: normalizeWindow(o.window),
        countedAt: o.countedAt ? String(o.countedAt) : null,
        stats: {
            cast: Number(o.stats && o.stats.cast) || 0,
            valid: Number(o.stats && o.stats.valid) || 0,
            invalid: Number(o.stats && o.stats.invalid) || 0,
            eligible: Number(o.stats && o.stats.eligible) || 0
        }
    };
}

/**
 * @param {unknown} raw
 * @param {{ electionKind: string, includeSgaDeputies?: boolean }} electionMeta
 */
function normalizeOffices(raw, electionMeta) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const base = emptyOfficesForKind(electionMeta);
    const decidedAfter =
        o.decidedAfter === 'runoff' ? 'runoff' : o.decidedAfter === 'main' ? 'main' : null;
    if (electionMeta.electionKind === ELECTION_KIND_SGA_LEHRER) {
        return {
            ...base,
            member1Id: normId(o.member1Id) || null,
            member2Id: normId(o.member2Id) || null,
            member3Id: normId(o.member3Id) || null,
            deputy1Id: normId(o.deputy1Id) || null,
            deputy2Id: normId(o.deputy2Id) || null,
            deputy3Id: normId(o.deputy3Id) || null,
            decidedAfter
        };
    }
    /** @type {Record<string, string|null>} */
    const out = {
        ...base,
        schulsprecherId: normId(o.schulsprecherId) || null,
        stv1Id: normId(o.stv1Id) || null,
        stv2Id: normId(o.stv2Id) || null,
        decidedAfter
    };
    if (electionMeta.includeSgaDeputies) {
        out.sgaDeputy1Id = normId(o.sgaDeputy1Id) || null;
        out.sgaDeputy2Id = normId(o.sgaDeputy2Id) || null;
        out.sgaDeputy3Id = normId(o.sgaDeputy3Id) || null;
    }
    return out;
}

/**
 * @param {unknown} raw
 * @param {{ electionKind: string, includeSgaDeputies?: boolean }} electionMeta
 */
function normalizeLot(raw, electionMeta) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const base = emptyLotForKind(electionMeta);
    if (electionMeta.electionKind === ELECTION_KIND_SGA_LEHRER) {
        return {
            ...base,
            member1Id: normId(o.member1Id) || null,
            member2Id: normId(o.member2Id) || null,
            member3Id: normId(o.member3Id) || null,
            deputy1Id: normId(o.deputy1Id) || null,
            deputy2Id: normId(o.deputy2Id) || null,
            deputy3Id: normId(o.deputy3Id) || null
        };
    }
    /** @type {Record<string, string|null>} */
    const out = {
        ...base,
        schulsprecherId: normId(o.schulsprecherId) || null,
        stv1Id: normId(o.stv1Id) || null,
        stv2Id: normId(o.stv2Id) || null
    };
    if (electionMeta.includeSgaDeputies) {
        out.sgaDeputy1Id = normId(o.sgaDeputy1Id) || null;
        out.sgaDeputy2Id = normId(o.sgaDeputy2Id) || null;
        out.sgaDeputy3Id = normId(o.sgaDeputy3Id) || null;
    }
    return out;
}

/**
 * @param {unknown} mode
 * @param {string} kind
 */
function normalizeEligibleMode(mode, kind) {
    const m = String(mode || '').trim().toLowerCase();
    if (m === 'manual') return 'manual';
    if (m === 'teachers') return 'teachers';
    if (m === 'students') return 'students';
    return getElectionKindDef(kind).eligibleModeDefault;
}

/**
 * @param {unknown} raw
 */
export function normalizeElection(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const electionKind = normalizeElectionKind(o.electionKind || o.kind);
    const kindDef = getElectionKindDef(electionKind);
    const includeSgaDeputies =
        electionKind === ELECTION_KIND_SCHULSPRECHER &&
        !!(o.includeSgaDeputies || o.includeSgaSchuelerStv);
    const electionMeta = { electionKind, includeSgaDeputies };
    const candidates = (Array.isArray(o.candidates) ? o.candidates : []).map(normalizeCandidate);
    let rounds = (Array.isArray(o.rounds) ? o.rounds : []).map(normalizeRound);
    if (!rounds.length) {
        rounds = [
            normalizeRound({
                id: 'main',
                kind: ROUND_MAIN,
                status: 'pending',
                candidateIds: candidates.map((c) => c.id)
            })
        ];
    }
    // SGA Lehrer: keine Stichwahl-Runden speichern/behalten
    if (electionKind === ELECTION_KIND_SGA_LEHRER) {
        rounds = rounds.filter((r) => r.kind !== ROUND_RUNOFF);
    }
    const status = Object.values(ELECTION_STATUS).includes(o.status) ? o.status : ELECTION_STATUS.draft;
    const defaultTitle = kindDef.defaultTitle;
    return {
        id: normId(o.id) || newId('wahl'),
        electionKind,
        includeSgaDeputies,
        title: String(o.title || defaultTitle).trim() || defaultTitle,
        scopeLabel: String(o.scopeLabel || 'Schulweit').trim() || 'Schulweit',
        schoolYear: String(o.schoolYear || '').trim(),
        status,
        window: normalizeWindow(o.window),
        runoffWindow: normalizeWindow(o.runoffWindow),
        candidates,
        eligibleCount: Math.max(0, Number(o.eligibleCount) || 0),
        eligibleMode: normalizeEligibleMode(o.eligibleMode, electionKind),
        rounds,
        offices: normalizeOffices(o.offices, electionMeta),
        lot: normalizeLot(o.lot, electionMeta),
        createdAt: o.createdAt ? String(o.createdAt) : new Date().toISOString(),
        updatedAt: o.updatedAt ? String(o.updatedAt) : new Date().toISOString(),
        publishedResults: !!o.publishedResults
    };
}

/**
 * @param {Partial<ReturnType<typeof normalizeElection>>} [partial]
 */
export function createElection(partial) {
    const now = new Date();
    const start = new Date(now.getTime());
    const end = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
    const electionKind = normalizeElectionKind(partial && (partial.electionKind || partial.kind));
    const kindDef = getElectionKindDef(electionKind);
    return normalizeElection({
        electionKind,
        title: kindDef.defaultTitle,
        scopeLabel: 'Schulweit',
        status: ELECTION_STATUS.draft,
        window: { start: start.toISOString(), end: end.toISOString() },
        eligibleMode: kindDef.eligibleModeDefault,
        ...partial,
        electionKind
    });
}

/**
 * @param {ReturnType<typeof normalizeWindow>} win
 * @param {Date} [now]
 */
export function isWithinWindow(win, now = new Date()) {
    if (!win) return true;
    const t = now.getTime();
    if (win.start) {
        const s = Date.parse(win.start);
        if (!Number.isNaN(s) && t < s) return false;
    }
    if (win.end) {
        const e = Date.parse(win.end);
        if (!Number.isNaN(e) && t > e) return false;
    }
    return true;
}

/**
 * Ob aktuell abgestimmt werden darf.
 * @param {ReturnType<typeof normalizeElection>} election
 * @param {Date} [now]
 */
export function votingPhase(election, now = new Date()) {
    const e = normalizeElection(election);
    if (e.status === ELECTION_STATUS.open) {
        const open = isWithinWindow(e.window, now) || (!e.window.start && !e.window.end);
        return {
            canVote: open,
            roundKind: ROUND_MAIN,
            round: e.rounds.find((r) => r.kind === ROUND_MAIN) || e.rounds[0],
            reason: open ? '' : 'Das Wahlfenster ist geschlossen.'
        };
    }
    if (e.status === ELECTION_STATUS.runoff_open) {
        if (e.electionKind === ELECTION_KIND_SGA_LEHRER) {
            return {
                canVote: false,
                roundKind: null,
                round: null,
                reason: 'Für die SGA-Lehrerwahl gibt es keine Stichwahl.'
            };
        }
        const round = e.rounds.find((r) => r.kind === ROUND_RUNOFF);
        const win = e.runoffWindow && (e.runoffWindow.start || e.runoffWindow.end) ? e.runoffWindow : e.window;
        const open = isWithinWindow(win, now) || (!win.start && !win.end);
        return {
            canVote: open && !!round,
            roundKind: ROUND_RUNOFF,
            round,
            reason: !round ? 'Keine Stichwahl konfiguriert.' : open ? '' : 'Das Stichwahl-Fenster ist geschlossen.'
        };
    }
    return {
        canVote: false,
        roundKind: null,
        round: null,
        reason:
            e.status === ELECTION_STATUS.draft
                ? 'Die Wahl ist noch nicht eröffnet.'
                : e.status === ELECTION_STATUS.finalized
                  ? 'Die Wahl ist abgeschlossen.'
                  : 'Derzeit keine Stimmabgabe möglich.'
    };
}

/**
 * @param {ReturnType<typeof normalizeElection>} election
 */
export function statusLabel(election) {
    const s = election && election.status;
    switch (s) {
        case ELECTION_STATUS.draft:
            return 'Entwurf';
        case ELECTION_STATUS.open:
            return 'Wahlgang offen';
        case ELECTION_STATUS.closed:
            return '1. Wahlgang geschlossen';
        case ELECTION_STATUS.runoff_open:
            return 'Stichwahl offen';
        case ELECTION_STATUS.runoff_closed:
            return 'Stichwahl geschlossen';
        case ELECTION_STATUS.finalized:
            return 'Abgeschlossen';
        default:
            return String(s || '–');
    }
}

/**
 * @param {unknown} raw
 */
export function normalizeStore(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        version: 1,
        elections: (Array.isArray(o.elections) ? o.elections : []).map(normalizeElection),
        /** @type {Array<{ electionId: string, roundId: string, voterKey: string, at: string }>} */
        voters: Array.isArray(o.voters)
            ? o.voters
                  .map((v) => ({
                      electionId: normId(v.electionId),
                      roundId: normId(v.roundId),
                      voterKey: String(v.voterKey || '').trim(),
                      at: String(v.at || '')
                  }))
                  .filter((v) => v.electionId && v.roundId && v.voterKey)
            : [],
        /** @type {Array<{ id: string, electionId: string, roundId: string, rankings?: string[], choice?: string, at: string }>} */
        ballots: Array.isArray(o.ballots)
            ? o.ballots
                  .map((b) => ({
                      id: normId(b.id) || newId('ballot'),
                      electionId: normId(b.electionId),
                      roundId: normId(b.roundId),
                      rankings: Array.isArray(b.rankings) ? b.rankings.map(normId) : undefined,
                      choice: b.choice != null ? normId(b.choice) : undefined,
                      at: String(b.at || '')
                  }))
                  .filter((b) => b.electionId && b.roundId)
            : []
    };
}
