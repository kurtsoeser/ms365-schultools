/**
 * Schulsprecherwahl nach österreichischer VO (BGBl. Nr. 388/1993):
 * Wahlpunkte 6 / 5 / 4 für Schulsprecher / 1. STV / 2. STV.
 * Schulsprecher braucht absolute Mehrheit der Erstplatzierungen, sonst Stichwahl.
 * Stellvertreter aus Punktezahl des 1. Wahlgangs (ohne gewählten Schulsprecher).
 */

/** @typedef {{ id: string, name?: string, className?: string, email?: string }} Candidate */
/** @typedef {{ rankings?: string[], choice?: string }} BallotInput */

export const POINTS_SCHULSPRECHER = 6;
export const POINTS_STV1 = 5;
export const POINTS_STV2 = 4;

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
 * Gültiger Hauptwahlgang: genau 3 unterschiedliche Kandidaten-IDs in Reihenfolge 6/5/4.
 * @param {BallotInput|null|undefined} ballot
 * @param {Iterable<string>} [allowedIds]
 */
export function validateMainBallot(ballot, allowedIds) {
    const rankings = Array.isArray(ballot && ballot.rankings) ? ballot.rankings.map(normId).filter(Boolean) : [];
    if (rankings.length !== 3) {
        return { ok: false, reason: 'Genau drei Reihungen (6 / 5 / 4 Punkte) nötig.' };
    }
    if (new Set(rankings).size !== 3) {
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
    return { ok: true, rankings };
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
 * @returns {Map<string, number>}
 */
export function pointsFromRankings(rankings) {
    const map = new Map();
    const r = rankings.map(normId);
    if (r[0]) map.set(r[0], (map.get(r[0]) || 0) + POINTS_SCHULSPRECHER);
    if (r[1]) map.set(r[1], (map.get(r[1]) || 0) + POINTS_STV1);
    if (r[2]) map.set(r[2], (map.get(r[2]) || 0) + POINTS_STV2);
    return map;
}

/**
 * @param {Array<BallotInput|null|undefined>} ballots
 * @param {string[]} candidateIds
 */
export function tallyMainRound(ballots, candidateIds) {
    const ids = [...new Set((candidateIds || []).map(normId).filter(Boolean))];
    /** @type {Map<string, { id: string, points: number, firstPlace: number, votes: number }>} */
    const byId = new Map(
        ids.map((id) => [id, { id, points: 0, firstPlace: 0, votes: 0 }])
    );

    let valid = 0;
    let invalid = 0;
    const allow = new Set(ids);

    for (const raw of ballots || []) {
        const check = validateMainBallot(raw, allow);
        if (!check.ok) {
            invalid += 1;
            continue;
        }
        valid += 1;
        const rankings = check.rankings;
        const pts = pointsFromRankings(rankings);
        for (const [id, n] of pts) {
            const row = byId.get(id);
            if (!row) continue;
            row.points += n;
            row.votes += 1;
        }
        const first = rankings[0];
        const firstRow = byId.get(first);
        if (firstRow) firstRow.firstPlace += 1;
    }

    const rows = ids.map((id) => byId.get(id)).filter(Boolean);
    rows.sort((a, b) => {
        if (b.firstPlace !== a.firstPlace) return b.firstPlace - a.firstPlace;
        if (b.points !== a.points) return b.points - a.points;
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
 * Vollständige Auswertung 1. Wahlgang (+ optional Stichwahl-Ergebnis für Amt).
 * @param {object} opts
 * @param {Array<BallotInput>} opts.mainBallots
 * @param {string[]} opts.candidateIds
 * @param {Array<BallotInput>} [opts.runoffBallots]
 * @param {string[]} [opts.runoffCandidateIds]
 * @param {string|null} [opts.lotSchulsprecherId] manuelle Losentscheidung
 * @param {string|null} [opts.lotStv1Id]
 * @param {string|null} [opts.lotStv2Id]
 */
export function evaluateElection(opts) {
    const candidateIds = (opts && opts.candidateIds) || [];
    const main = tallyMainRound((opts && opts.mainBallots) || [], candidateIds);
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

    return {
        main,
        first,
        runoffTally,
        runoffResolve,
        needsRunoff: needsRunoff && !schulsprecherId,
        runoffCandidateIds,
        majorityThreshold: first.majorityThreshold,
        offices: {
            schulsprecherId,
            stv1Id,
            stv2Id,
            decidedAfter,
            tieNeedsLotSchulsprecher,
            tieNeedsLotStv1: deputies.tieNeedsLotStv1 && !opts?.lotStv1Id,
            tieNeedsLotStv2: deputies.tieNeedsLotStv2 && !opts?.lotStv2Id
        }
    };
}

/**
 * UI-Zeilen für Ergebnis-Tabelle (Hauptwahlgang).
 * @param {ReturnType<typeof tallyMainRound>} tally
 * @param {Candidate[]} candidates
 * @param {{ schulsprecherId?: string|null, stv1Id?: string|null, stv2Id?: string|null }} [offices]
 */
export function mainResultRows(tally, candidates, offices) {
    const byCand = new Map((candidates || []).map((c) => [normId(c.id), c]));
    const ss = normId(offices && offices.schulsprecherId);
    const s1 = normId(offices && offices.stv1Id);
    const s2 = normId(offices && offices.stv2Id);
    const valid = tally.valid;

    return (tally.rows || []).map((row) => {
        const c = byCand.get(row.id) || { id: row.id, name: row.id };
        let badge = '';
        if (row.id === ss) badge = 'Gewählt · Schulsprecher:in';
        else if (row.id === s1) badge = 'Gewählt · 1. Stellvertretung';
        else if (row.id === s2) badge = 'Gewählt · 2. Stellvertretung';
        return {
            id: row.id,
            name: c.name || row.id,
            className: c.className || '',
            email: c.email || '',
            points: row.points,
            firstPlace: row.firstPlace,
            firstPlaceShare: sharePercent(row.firstPlace, valid),
            pointsShare: sharePercent(row.points, valid * POINTS_SCHULSPRECHER),
            badge,
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
