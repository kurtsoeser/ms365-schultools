/**
 * Schulwahlen: Schulsprecher (BGBl. 388/1993) und SGA Lehrer (BGBl. 389/1993).
 * Schulsprecher: 6/5/4, absolute Mehrheit der Erstplatzierungen sonst Stichwahl.
 * Optional: SGA-Schüler-STV 3/2/1 am selben Stimmzettel.
 * SGA Lehrer: 6…1 Punkte, Top 3 Mitglieder + Platz 4–6 Stellvertretung.
 */

import {
    ELECTION_KIND_SCHULSPRECHER,
    ELECTION_KIND_SGA_LEHRER,
    getElectionKindDef,
    normalizeElectionKind,
    resolveElectionBallotSchema,
    seatsForElection
} from './schulsprecherwahl-kinds.js';

/** @typedef {{ id: string, name?: string, className?: string, email?: string }} Candidate */
/** @typedef {{ rankings?: string[], choice?: string }} BallotInput */

export const POINTS_SCHULSPRECHER = 6;
export const POINTS_STV1 = 5;
export const POINTS_STV2 = 4;
export const POINTS_SGA_LEHRER = Object.freeze([6, 5, 4, 3, 2, 1]);
export const POINTS_SGA_SCHUELER = Object.freeze([3, 2, 1]);

export const ROUND_MAIN = 'main';
export const ROUND_RUNOFF = 'runoff';

/**
 * @param {unknown} value
 * @returns {string}
 */
export function normId(value) {
    return String(value == null ? '' : value).trim();
}

/**
 * @param {unknown} kindOrOpts
 */
function resolveBallotSchema(kindOrOpts) {
    const schema = resolveElectionBallotSchema(kindOrOpts);
    if (kindOrOpts && typeof kindOrOpts === 'object' && !Array.isArray(kindOrOpts)) {
        if (Number(kindOrOpts.rankingCount) > 0) {
            schema.rankingCount = Number(kindOrOpts.rankingCount);
        }
        if (Array.isArray(kindOrOpts.points) && kindOrOpts.points.length) {
            schema.points = kindOrOpts.points.map(Number);
        }
    }
    return schema;
}

/**
 * Gültiger Hauptwahlgang: unterschiedliche Kandidaten-IDs gemäß Wahltyp (3 oder 6).
 * @param {BallotInput|null|undefined} ballot
 * @param {Iterable<string>} [allowedIds]
 * @param {unknown} [kindOrOpts] electionKind oder { kind, rankingCount, points }
 */
export function validateMainBallot(ballot, allowedIds, kindOrOpts) {
    const schema = resolveBallotSchema(kindOrOpts);
    const rankings = Array.isArray(ballot && ballot.rankings) ? ballot.rankings.map(normId).filter(Boolean) : [];
    if (rankings.length !== schema.rankingCount) {
        const pts = schema.points.join(' / ');
        return {
            ok: false,
            reason:
                schema.rankingCount === 3
                    ? 'Genau drei Reihungen (6 / 5 / 4 Punkte) nötig.'
                    : 'Genau ' + schema.rankingCount + ' Reihungen (' + pts + ' Punkte) nötig.'
        };
    }
    if (new Set(rankings).size !== schema.rankingCount) {
        return { ok: false, reason: 'Jeder Kandidat darf nur einmal gereiht werden.' };
    }
    if (allowedIds) {
        const allow = new Set([...allowedIds].map(normId).filter(Boolean));
        for (const id of rankings) {
            if (!allow.has(id)) {
                return { ok: false, reason: 'Unbekannte Kandidatur.' };
            }
        }
    }
    return { ok: true, rankings, kind: schema.kind };
}

/**
 * Stichwahl: eine Stimme für einen der beiden Kandidaten.
 * @param {BallotInput|null|undefined} ballot
 * @param {Iterable<string>} [allowedIds]
 */
export function validateRunoffBallot(ballot, allowedIds) {
    const choice = normId(ballot && ballot.choice);
    if (!choice) {
        return { ok: false, reason: 'Bitte eine Person auswählen.' };
    }
    if (allowedIds) {
        const allow = new Set([...allowedIds].map(normId).filter(Boolean));
        if (!allow.has(choice)) {
            return { ok: false, reason: 'Unbekannte Kandidatur.' };
        }
    }
    return { ok: true, choice };
}

/**
 * @param {string[]} rankings
 * @param {number[]|unknown} [pointsOrKind]
 * @returns {Map<string, number>}
 */
export function pointsFromRankings(rankings, pointsOrKind) {
    const map = new Map();
    const r = rankings.map(normId);
    const schema = resolveBallotSchema(pointsOrKind);
    const points =
        Array.isArray(pointsOrKind) && pointsOrKind.length
            ? pointsOrKind.map(Number)
            : schema.points;
    for (let i = 0; i < r.length; i++) {
        const id = r[i];
        const p = Number(points[i]);
        if (!id || !Number.isFinite(p)) continue;
        map.set(id, (map.get(id) || 0) + p);
    }
    return map;
}

/**
 * @param {Array<BallotInput|null|undefined>} ballots
 * @param {string[]} candidateIds
 * @param {unknown} [kindOrOpts]
 */
export function tallyMainRound(ballots, candidateIds, kindOrOpts) {
    const schema = resolveBallotSchema(kindOrOpts);
    const ids = [...new Set((candidateIds || []).map(normId).filter(Boolean))];
    /** @type {Map<string, { id: string, points: number, firstPlace: number, votes: number, sgaPoints: number }>} */
    const byId = new Map(
        ids.map((id) => [id, { id, points: 0, firstPlace: 0, votes: 0, sgaPoints: 0 }])
    );

    let valid = 0;
    let invalid = 0;
    const allow = new Set(ids);
    const coreCount = Number(schema.coreRankingCount) || schema.rankingCount;
    const corePoints = Array.isArray(schema.corePoints) ? schema.corePoints : schema.points;
    const sgaPoints = Array.isArray(schema.sgaPoints) ? schema.sgaPoints : [];

    for (const raw of ballots || []) {
        const check = validateMainBallot(raw, allow, schema);
        if (!check.ok) {
            invalid += 1;
            continue;
        }
        valid += 1;
        const rankings = check.rankings;
        // Schulsprecher-Ämter: nur 6/5/4 (erste drei); SGA-Lehrer: volle Punkte
        const coreRankings = rankings.slice(0, coreCount);
        const pts = pointsFromRankings(coreRankings, corePoints);
        for (const [id, n] of pts) {
            const row = byId.get(id);
            if (!row) continue;
            row.points += n;
            row.votes += 1;
        }
        const first = rankings[0];
        const firstRow = byId.get(first);
        if (firstRow) firstRow.firstPlace += 1;

        if (schema.includeSgaDeputies && sgaPoints.length) {
            const sgaRankings = rankings.slice(coreCount, coreCount + sgaPoints.length);
            const sgaPts = pointsFromRankings(sgaRankings, sgaPoints);
            for (const [id, n] of sgaPts) {
                const row = byId.get(id);
                if (!row) continue;
                row.sgaPoints += n;
            }
        }
    }

    const rows = ids.map((id) => byId.get(id)).filter(Boolean);
    if (schema.kind === ELECTION_KIND_SGA_LEHRER) {
        rows.sort((a, b) => {
            if (b.points !== a.points) return b.points - a.points;
            return a.id.localeCompare(b.id, 'de');
        });
    } else {
        rows.sort((a, b) => {
            if (b.firstPlace !== a.firstPlace) return b.firstPlace - a.firstPlace;
            if (b.points !== a.points) return b.points - a.points;
            return a.id.localeCompare(b.id, 'de');
        });
    }

    return {
        valid,
        invalid,
        cast: valid + invalid,
        rows,
        byId,
        kind: schema.kind,
        includeSgaDeputies: !!schema.includeSgaDeputies
    };
}

/**
 * SGA-Schüler-Stellvertreter: Top 3 nach 3/2/1-Punkten, ohne bereits gewählte SV-Ämter.
 * @param {{ rows?: Array<{ id: string, sgaPoints?: number, points?: number }>, byId?: Map<string, { id: string, sgaPoints?: number }> }} tally
 * @param {Iterable<string>} [excludeIds]
 */
export function resolveSgaSchuelerDeputies(tally, excludeIds) {
    const exclude = new Set([...((excludeIds && [...excludeIds]) || [])].map(normId).filter(Boolean));
    /** @type {Array<{ id: string, sgaPoints: number }>} */
    let rows = [];
    if (tally && tally.byId) {
        rows = [...tally.byId.values()].map((r) => ({
            id: r.id,
            sgaPoints: Number(r.sgaPoints) || 0
        }));
    } else {
        rows = (Array.isArray(tally && tally.rows) ? tally.rows : []).map((r) => ({
            id: r.id,
            sgaPoints: Number(r.sgaPoints != null ? r.sgaPoints : r.points) || 0
        }));
    }
    rows = rows
        .filter((r) => !exclude.has(normId(r.id)))
        .sort((a, b) => b.sgaPoints - a.sgaPoints || a.id.localeCompare(b.id, 'de'));

    const pick = [];
    let needsLot = false;
    let idx = 0;
    while (pick.length < 3 && idx < rows.length) {
        const pts = rows[idx].sgaPoints;
        const group = [];
        while (idx < rows.length && rows[idx].sgaPoints === pts) {
            group.push(rows[idx]);
            idx += 1;
        }
        const slotsLeft = 3 - pick.length;
        if (group.length > slotsLeft) needsLot = true;
        for (const r of group) {
            if (pick.length >= 3) break;
            pick.push(r.id);
        }
    }

    return {
        sgaDeputy1Id: pick[0] || null,
        sgaDeputy2Id: pick[1] || null,
        sgaDeputy3Id: pick[2] || null,
        tieNeedsLot: needsLot
    };
}

/**
 * @param {Array<BallotInput|null|undefined>} ballots
 * @param {string[]} candidateIds
 */
export function tallyRunoffRound(ballots, candidateIds) {
    const ids = [...new Set((candidateIds || []).map(normId).filter(Boolean))];
    const allow = new Set(ids);
    /** @type {Map<string, { id: string, votes: number }>} */
    const byId = new Map(ids.map((id) => [id, { id, votes: 0 }]));

    let valid = 0;
    let invalid = 0;

    for (const raw of ballots || []) {
        const check = validateRunoffBallot(raw, allow);
        if (!check.ok) {
            invalid += 1;
            continue;
        }
        valid += 1;
        const row = byId.get(check.choice);
        if (row) row.votes += 1;
    }

    const rows = ids.map((id) => byId.get(id)).filter(Boolean);
    rows.sort((a, b) => {
        if (b.votes !== a.votes) return b.votes - a.votes;
        return a.id.localeCompare(b.id, 'de');
    });

    return {
        valid,
        invalid,
        cast: valid + invalid,
        rows,
        byId
    };
}

/**
 * Absolute Mehrheit der gültigen Stimmzettel als Erstgereihte:in.
 * @param {{ valid: number, rows: Array<{ id: string, firstPlace: number, points: number }> }} tally
 */
export function resolveSchulsprecherFromMain(tally) {
    const valid = Number(tally && tally.valid) || 0;
    const rows = Array.isArray(tally && tally.rows) ? tally.rows : [];
    if (valid < 1 || !rows.length) {
        return {
            decided: false,
            schulsprecherId: null,
            needsRunoff: false,
            runoffCandidateIds: [],
            tieNeedsLot: false,
            majorityThreshold: 0
        };
    }

    const majorityThreshold = Math.floor(valid / 2) + 1;
    const top = rows[0];
    if (top.firstPlace >= majorityThreshold) {
        return {
            decided: true,
            schulsprecherId: top.id,
            needsRunoff: false,
            runoffCandidateIds: [],
            tieNeedsLot: false,
            majorityThreshold
        };
    }

    // Stichwahl: zwei häufigste Erstplatzierungen; bei >2 mit gleicher Häufigkeit Punkte, sonst Los.
    const firstPlaceMax = top.firstPlace;
    const tiedFirst = rows.filter((r) => r.firstPlace === firstPlaceMax);
    if (tiedFirst.length === 2) {
        return {
            decided: false,
            schulsprecherId: null,
            needsRunoff: true,
            runoffCandidateIds: tiedFirst.map((r) => r.id),
            tieNeedsLot: false,
            majorityThreshold
        };
    }
    if (tiedFirst.length === 1) {
        const secondPlace = rows[1] ? rows[1].firstPlace : -1;
        const tiedSecond = rows.filter((r) => r.firstPlace === secondPlace);
        if (tiedSecond.length === 1) {
            return {
                decided: false,
                schulsprecherId: null,
                needsRunoff: true,
                runoffCandidateIds: [tiedFirst[0].id, tiedSecond[0].id],
                tieNeedsLot: false,
                majorityThreshold
            };
        }
        // Mehrere Zweitplatzierte mit gleicher Erst-Anzahl → nach Punkten
        const byPoints = tiedSecond.slice().sort((a, b) => b.points - a.points || a.id.localeCompare(b.id, 'de'));
        const maxPts = byPoints[0].points;
        const ptsTied = byPoints.filter((r) => r.points === maxPts);
        if (ptsTied.length === 1) {
            return {
                decided: false,
                schulsprecherId: null,
                needsRunoff: true,
                runoffCandidateIds: [tiedFirst[0].id, ptsTied[0].id],
                tieNeedsLot: false,
                majorityThreshold
            };
        }
        return {
            decided: false,
            schulsprecherId: null,
            needsRunoff: true,
            runoffCandidateIds: [tiedFirst[0].id, ...ptsTied.map((r) => r.id)],
            tieNeedsLot: true,
            majorityThreshold
        };
    }

    // >2 mit gleicher höchster Erstplatzierung → die beiden mit den meisten Punkten, sonst Los
    const byPoints = tiedFirst.slice().sort((a, b) => b.points - a.points || a.id.localeCompare(b.id, 'de'));
    const firstPts = byPoints[0].points;
    const atFirst = byPoints.filter((r) => r.points === firstPts);
    if (atFirst.length === 2) {
        return {
            decided: false,
            schulsprecherId: null,
            needsRunoff: true,
            runoffCandidateIds: atFirst.map((r) => r.id),
            tieNeedsLot: false,
            majorityThreshold
        };
    }
    if (atFirst.length > 2) {
        return {
            decided: false,
            schulsprecherId: null,
            needsRunoff: true,
            runoffCandidateIds: atFirst.map((r) => r.id),
            tieNeedsLot: true,
            majorityThreshold
        };
    }
    // genau einer mit höchsten Punkten → Partner = nächstbeste Punktegruppe
    const rest = byPoints.filter((r) => r.points !== firstPts);
    if (!rest.length) {
        return {
            decided: false,
            schulsprecherId: null,
            needsRunoff: true,
            runoffCandidateIds: [atFirst[0].id],
            tieNeedsLot: true,
            majorityThreshold
        };
    }
    const secondPts = rest[0].points;
    const atSecond = rest.filter((r) => r.points === secondPts);
    if (atSecond.length === 1) {
        return {
            decided: false,
            schulsprecherId: null,
            needsRunoff: true,
            runoffCandidateIds: [atFirst[0].id, atSecond[0].id],
            tieNeedsLot: false,
            majorityThreshold
        };
    }
    return {
        decided: false,
        schulsprecherId: null,
        needsRunoff: true,
        runoffCandidateIds: [atFirst[0].id, ...atSecond.map((r) => r.id)],
        tieNeedsLot: true,
        majorityThreshold
    };
}

/**
 * 1. und 2. Stellvertreter: höchste / zweithöchste Punktezahl ohne Schulsprecher.
 * @param {{ rows: Array<{ id: string, points: number }> }} tally
 * @param {string|null|undefined} schulsprecherId
 */
export function resolveDeputiesFromMain(tally, schulsprecherId) {
    const exclude = normId(schulsprecherId);
    const rows = (Array.isArray(tally && tally.rows) ? tally.rows : [])
        .filter((r) => r && normId(r.id) !== exclude)
        .slice()
        .sort((a, b) => b.points - a.points || a.id.localeCompare(b.id, 'de'));

    const stv1 = rows[0] || null;
    const stv2 = rows[1] || null;
    const tieStv1 = stv1 && rows.filter((r) => r.points === stv1.points).length > 1;
    const remaining = stv1 ? rows.filter((r) => r.id !== stv1.id) : [];
    const tieStv2 = stv2 && remaining.filter((r) => r.points === stv2.points).length > 1;

    return {
        stv1Id: stv1 ? stv1.id : null,
        stv2Id: stv2 ? stv2.id : null,
        tieNeedsLotStv1: !!tieStv1,
        tieNeedsLotStv2: !!tieStv2
    };
}

/**
 * SGA Lehrer: Top 3 nach Punkten = Mitglieder, Platz 4–6 = Stellvertretung.
 * Los nur, wenn Gleichstand die Grenze Mitglieder↔Stellvertretung oder Stellvertretung↔aus
 * berührt (BGBl. 389/1993 § 5 Abs. 2). Reihenfolge innerhalb klarer Blöcke: Punkte, dann id.
 * @param {{ rows: Array<{ id: string, points: number }> }} tally
 */
export function resolveSgaLehrerFromMain(tally) {
    const rows = (Array.isArray(tally && tally.rows) ? tally.rows : [])
        .slice()
        .sort((a, b) => b.points - a.points || a.id.localeCompare(b.id, 'de'));

    /**
     * Nimmt bis zu `take` Sitze; bei Gleichstand über die Schnittkante → Los.
     * @param {Array<{ id: string, points: number }>} pool
     * @param {number} take
     */
    function takeSeats(pool, take) {
        /** @type {string[]} */
        const chosen = [];
        let needsLot = false;
        let idx = 0;
        while (chosen.length < take && idx < pool.length) {
            const pts = pool[idx].points;
            const group = [];
            while (idx < pool.length && pool[idx].points === pts) {
                group.push(pool[idx]);
                idx += 1;
            }
            const slotsLeft = take - chosen.length;
            if (group.length > slotsLeft) {
                needsLot = true;
            }
            for (const r of group) {
                if (chosen.length >= take) break;
                chosen.push(r.id);
            }
        }
        const remaining = pool.filter((r) => !chosen.includes(r.id));
        return { chosen, remaining, needsLot };
    }

    const members = takeSeats(rows, 3);
    const deputies = takeSeats(members.remaining, 3);
    const ordered = members.chosen.concat(deputies.chosen);

    return {
        member1Id: ordered[0] || null,
        member2Id: ordered[1] || null,
        member3Id: ordered[2] || null,
        deputy1Id: ordered[3] || null,
        deputy2Id: ordered[4] || null,
        deputy3Id: ordered[5] || null,
        orderedIds: ordered,
        tieNeedsLotMembers: members.needsLot,
        tieNeedsLotDeputies: deputies.needsLot,
        // Kompatibilität zu seat-weisen Flags in evaluateSgaLehrerElection
        tieNeedsLotMember1: members.needsLot,
        tieNeedsLotMember2: members.needsLot,
        tieNeedsLotMember3: members.needsLot,
        tieNeedsLotDeputy1: deputies.needsLot,
        tieNeedsLotDeputy2: deputies.needsLot,
        tieNeedsLotDeputy3: deputies.needsLot,
        anyTieNeedsLot: members.needsLot || deputies.needsLot
    };
}

/**
 * @param {{ rows: Array<{ id: string, votes: number }> }} tally
 */
export function resolveSchulsprecherFromRunoff(tally) {
    const rows = Array.isArray(tally && tally.rows) ? tally.rows : [];
    if (!rows.length) {
        return { schulsprecherId: null, tieNeedsLot: false };
    }
    const top = rows[0];
    const tied = rows.filter((r) => r.votes === top.votes);
    if (tied.length !== 1) {
        return { schulsprecherId: null, tieNeedsLot: true, tiedIds: tied.map((r) => r.id) };
    }
    return { schulsprecherId: top.id, tieNeedsLot: false, tiedIds: [] };
}

/**
 * @param {number} cast
 * @param {number} eligible
 */
export function turnoutPercent(cast, eligible) {
    const e = Number(eligible) || 0;
    const c = Number(cast) || 0;
    if (e <= 0) return null;
    return Math.round((c / e) * 1000) / 10;
}

/**
 * Prozentanteil bezogen auf gültige Stimmen.
 * @param {number} part
 * @param {number} valid
 */
export function sharePercent(part, valid) {
    const v = Number(valid) || 0;
    const p = Number(part) || 0;
    if (v <= 0) return 0;
    return Math.round((p / v) * 1000) / 10;
}

/**
 * @param {object} opts
 * @param {Array<BallotInput>} opts.mainBallots
 * @param {string[]} opts.candidateIds
 * @param {object} [opts.lot]
 */
function evaluateSgaLehrerElection(opts) {
    const candidateIds = (opts && opts.candidateIds) || [];
    const main = tallyMainRound((opts && opts.mainBallots) || [], candidateIds, ELECTION_KIND_SGA_LEHRER);
    const resolved = resolveSgaLehrerFromMain(main);
    const lot = (opts && opts.lot) || {};

    const pick = (key, fallback, needsLot) => {
        const fromLot = normId(lot[key]);
        if (needsLot && fromLot) return fromLot;
        return fallback;
    };

    const offices = {
        member1Id: pick('member1Id', resolved.member1Id, resolved.tieNeedsLotMember1),
        member2Id: pick('member2Id', resolved.member2Id, resolved.tieNeedsLotMember2),
        member3Id: pick('member3Id', resolved.member3Id, resolved.tieNeedsLotMember3),
        deputy1Id: pick('deputy1Id', resolved.deputy1Id, resolved.tieNeedsLotDeputy1),
        deputy2Id: pick('deputy2Id', resolved.deputy2Id, resolved.tieNeedsLotDeputy2),
        deputy3Id: pick('deputy3Id', resolved.deputy3Id, resolved.tieNeedsLotDeputy3),
        decidedAfter: 'main',
        tieNeedsLotMember1: resolved.tieNeedsLotMember1 && !normId(lot.member1Id),
        tieNeedsLotMember2: resolved.tieNeedsLotMember2 && !normId(lot.member2Id),
        tieNeedsLotMember3: resolved.tieNeedsLotMember3 && !normId(lot.member3Id),
        tieNeedsLotDeputy1: resolved.tieNeedsLotDeputy1 && !normId(lot.deputy1Id),
        tieNeedsLotDeputy2: resolved.tieNeedsLotDeputy2 && !normId(lot.deputy2Id),
        tieNeedsLotDeputy3: resolved.tieNeedsLotDeputy3 && !normId(lot.deputy3Id)
    };

    return {
        kind: ELECTION_KIND_SGA_LEHRER,
        main,
        first: null,
        sga: resolved,
        runoffTally: null,
        runoffResolve: null,
        needsRunoff: false,
        runoffCandidateIds: [],
        majorityThreshold: 0,
        offices
    };
}

/**
 * Vollständige Auswertung 1. Wahlgang (+ optional Stichwahl-Ergebnis für Amt).
 * @param {object} opts
 * @param {Array<BallotInput>} opts.mainBallots
 * @param {string[]} opts.candidateIds
 * @param {string} [opts.electionKind]
 * @param {Array<BallotInput>} [opts.runoffBallots]
 * @param {string[]} [opts.runoffCandidateIds]
 * @param {string|null} [opts.lotSchulsprecherId] manuelle Losentscheidung
 * @param {string|null} [opts.lotStv1Id]
 * @param {string|null} [opts.lotStv2Id]
 * @param {object} [opts.lot] Los-Overrides (SGA Lehrer)
 */
export function evaluateElection(opts) {
    const kind = normalizeElectionKind(opts && (opts.electionKind || opts.kind));
    if (kind === ELECTION_KIND_SGA_LEHRER) {
        return evaluateSgaLehrerElection(opts || {});
    }

    const includeSgaDeputies = !!(opts && (opts.includeSgaDeputies || opts.includeSgaSchuelerStv));
    const ballotOpts = {
        electionKind: ELECTION_KIND_SCHULSPRECHER,
        includeSgaDeputies
    };
    const candidateIds = (opts && opts.candidateIds) || [];
    const main = tallyMainRound((opts && opts.mainBallots) || [], candidateIds, ballotOpts);
    const first = resolveSchulsprecherFromMain(main);

    let schulsprecherId = first.schulsprecherId;
    let decidedAfter = /** @type {'main'|'runoff'|null} */ (first.decided ? 'main' : null);
    let runoffTally = null;
    let runoffResolve = null;
    let needsRunoff = first.needsRunoff;
    let runoffCandidateIds = first.runoffCandidateIds.slice();
    let tieNeedsLotSchulsprecher = first.tieNeedsLot;

    if (first.tieNeedsLot && opts && opts.lotSchulsprecherId) {
        const pick = normId(opts.lotSchulsprecherId);
        if (runoffCandidateIds.includes(pick) || candidateIds.map(normId).includes(pick)) {
            // Los entscheidet, wer in die Stichwahl kommt – wenn genau 2 übrig: runoff ids setzen
            if (runoffCandidateIds.length > 2) {
                const others = runoffCandidateIds.filter((id) => id !== pick);
                // Mit Los einen Partner wählen, wenn Admin lotSchulsprecherId als einen der beiden setzt
                // und lotRunoffPartnerId fehlt: behalte pick + höchste Punkte unter Rest
                runoffCandidateIds = [pick, others[0]].filter(Boolean);
                tieNeedsLotSchulsprecher = runoffCandidateIds.length !== 2;
            }
        }
    }

    if (!first.decided && Array.isArray(opts && opts.runoffBallots) && opts.runoffBallots.length) {
        const runoffIds =
            Array.isArray(opts.runoffCandidateIds) && opts.runoffCandidateIds.length
                ? opts.runoffCandidateIds.map(normId)
                : runoffCandidateIds;
        runoffTally = tallyRunoffRound(opts.runoffBallots, runoffIds);
        runoffResolve = resolveSchulsprecherFromRunoff(runoffTally);
        if (runoffResolve.tieNeedsLot && opts && opts.lotSchulsprecherId) {
            schulsprecherId = normId(opts.lotSchulsprecherId);
            decidedAfter = 'runoff';
            needsRunoff = false;
            tieNeedsLotSchulsprecher = false;
        } else if (runoffResolve.schulsprecherId) {
            schulsprecherId = runoffResolve.schulsprecherId;
            decidedAfter = 'runoff';
            needsRunoff = false;
            tieNeedsLotSchulsprecher = false;
        } else {
            tieNeedsLotSchulsprecher = !!runoffResolve.tieNeedsLot;
        }
    }

    const deputies = resolveDeputiesFromMain(main, schulsprecherId);
    let stv1Id = deputies.stv1Id;
    let stv2Id = deputies.stv2Id;
    if (deputies.tieNeedsLotStv1 && opts && opts.lotStv1Id) stv1Id = normId(opts.lotStv1Id);
    if (deputies.tieNeedsLotStv2 && opts && opts.lotStv2Id) stv2Id = normId(opts.lotStv2Id);

    /** @type {ReturnType<typeof resolveSgaSchuelerDeputies>|null} */
    let sgaSchueler = null;
    /** @type {Record<string, string|null|boolean>} */
    const offices = {
        schulsprecherId,
        stv1Id,
        stv2Id,
        decidedAfter,
        tieNeedsLotSchulsprecher,
        tieNeedsLotStv1: deputies.tieNeedsLotStv1 && !opts?.lotStv1Id,
        tieNeedsLotStv2: deputies.tieNeedsLotStv2 && !opts?.lotStv2Id
    };

    if (includeSgaDeputies) {
        const lot = (opts && opts.lot) || {};
        sgaSchueler = resolveSgaSchuelerDeputies(main, [schulsprecherId, stv1Id, stv2Id]);
        let sgaDeputy1Id = sgaSchueler.sgaDeputy1Id;
        let sgaDeputy2Id = sgaSchueler.sgaDeputy2Id;
        let sgaDeputy3Id = sgaSchueler.sgaDeputy3Id;
        if (sgaSchueler.tieNeedsLot && normId(lot.sgaDeputy1Id)) sgaDeputy1Id = normId(lot.sgaDeputy1Id);
        if (sgaSchueler.tieNeedsLot && normId(lot.sgaDeputy2Id)) sgaDeputy2Id = normId(lot.sgaDeputy2Id);
        if (sgaSchueler.tieNeedsLot && normId(lot.sgaDeputy3Id)) sgaDeputy3Id = normId(lot.sgaDeputy3Id);
        offices.sgaDeputy1Id = sgaDeputy1Id;
        offices.sgaDeputy2Id = sgaDeputy2Id;
        offices.sgaDeputy3Id = sgaDeputy3Id;
        offices.tieNeedsLotSgaDeputies = sgaSchueler.tieNeedsLot && !normId(lot.sgaDeputy1Id);
    }

    return {
        kind: ELECTION_KIND_SCHULSPRECHER,
        includeSgaDeputies,
        main,
        first,
        sga: null,
        sgaSchueler,
        runoffTally,
        runoffResolve,
        needsRunoff: needsRunoff && !schulsprecherId,
        runoffCandidateIds,
        majorityThreshold: first.majorityThreshold,
        offices
    };
}

/**
 * UI-Zeilen für Ergebnis-Tabelle (Hauptwahlgang).
 * @param {ReturnType<typeof tallyMainRound>} tally
 * @param {Candidate[]} candidates
 * @param {object} [offices]
 * @param {unknown} [kindOrOpts]
 */
export function mainResultRows(tally, candidates, offices, kindOrOpts) {
    const schema = resolveBallotSchema(kindOrOpts || (tally && tally.kind));
    const seats = seatsForElection(kindOrOpts || schema);
    const byCand = new Map((candidates || []).map((c) => [normId(c.id), c]));
    const valid = tally.valid;
    const maxPoints = Math.max(Number(schema.corePoints && schema.corePoints[0]) || POINTS_SCHULSPRECHER, 1);

    /** @type {Map<string, string>} */
    const badgeById = new Map();
    for (const seat of seats) {
        const cid = normId(offices && offices[seat.key]);
        if (cid) badgeById.set(cid, seat.badge);
    }

    return (tally.rows || []).map((row) => {
        const c = byCand.get(row.id) || { id: row.id, name: row.id };
        const sgaPts = Number(row.sgaPoints) || 0;
        return {
            id: row.id,
            name: c.name || row.id,
            className: c.className || '',
            email: c.email || '',
            points: row.points,
            sgaPoints: sgaPts,
            firstPlace: row.firstPlace,
            firstPlaceShare: sharePercent(row.firstPlace, valid),
            pointsShare: sharePercent(row.points, valid * maxPoints),
            badge: badgeById.get(row.id) || '',
            barValue: row.points,
            barMax: Math.max(...tally.rows.map((r) => r.points), 1)
        };
    });
}

/**
 * UI-Zeilen Stichwahl.
 * @param {ReturnType<typeof tallyRunoffRound>} tally
 * @param {Candidate[]} candidates
 * @param {string|null|undefined} winnerId
 */
export function runoffResultRows(tally, candidates, winnerId) {
    const byCand = new Map((candidates || []).map((c) => [normId(c.id), c]));
    const win = normId(winnerId);
    const valid = tally.valid;
    const maxVotes = Math.max(...(tally.rows || []).map((r) => r.votes), 1);

    return (tally.rows || []).map((row) => {
        const c = byCand.get(row.id) || { id: row.id, name: row.id };
        return {
            id: row.id,
            name: c.name || row.id,
            className: c.className || '',
            email: c.email || '',
            votes: row.votes,
            share: sharePercent(row.votes, valid),
            badge: row.id === win ? 'Gewählt' : '',
            barValue: row.votes,
            barMax: maxVotes
        };
    });
}
