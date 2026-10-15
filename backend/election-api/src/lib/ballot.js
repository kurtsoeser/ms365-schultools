'use strict';

/**
 * Leichte Server-Validierung (Spiegel der Frontend-Logik).
 */

function normId(value) {
    return String(value == null ? '' : value).trim();
}

/**
 * @param {unknown} ballot
 * @param {Iterable<string>} allowedIds
 */
function validateMainBallot(ballot, allowedIds) {
    const rankings = Array.isArray(ballot && ballot.rankings)
        ? ballot.rankings.map(normId).filter(Boolean)
        : [];
    if (rankings.length !== 3) {
        return { ok: false, reason: 'Genau drei Reihungen (6 / 5 / 4 Punkte) nötig.' };
    }
    if (new Set(rankings).size !== 3) {
        return { ok: false, reason: 'Jeder Kandidat darf nur einmal gereiht werden.' };
    }
    const allow = new Set([...allowedIds].map(normId).filter(Boolean));
    for (const id of rankings) {
        if (!allow.has(id)) return { ok: false, reason: 'Unbekannte Kandidatur.' };
    }
    return { ok: true, rankings };
}

/**
 * @param {unknown} ballot
 * @param {Iterable<string>} allowedIds
 */
function validateRunoffBallot(ballot, allowedIds) {
    const choice = normId(ballot && ballot.choice);
    if (!choice) return { ok: false, reason: 'Bitte eine Person auswählen.' };
    const allow = new Set([...allowedIds].map(normId).filter(Boolean));
    if (!allow.has(choice)) return { ok: false, reason: 'Unbekannte Kandidatur.' };
    return { ok: true, choice };
}

/**
 * @param {object} election
 * @param {Date} [now]
 */
function votingPhase(election, now = new Date()) {
    const status = String((election && election.status) || '');
    const within = (win) => {
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
    };

    const rounds = Array.isArray(election && election.rounds) ? election.rounds : [];
    if (status === 'open') {
        const round = rounds.find((r) => r.kind === 'main') || rounds[0];
        const open = within(election.window);
        return {
            canVote: open && !!round,
            roundKind: 'main',
            round,
            reason: open ? '' : 'Das Wahlfenster ist geschlossen.'
        };
    }
    if (status === 'runoff_open') {
        const round = rounds.find((r) => r.kind === 'runoff');
        const win =
            election.runoffWindow && (election.runoffWindow.start || election.runoffWindow.end)
                ? election.runoffWindow
                : election.window;
        const open = within(win);
        return {
            canVote: open && !!round,
            roundKind: 'runoff',
            round,
            reason: !round
                ? 'Keine Stichwahl konfiguriert.'
                : open
                  ? ''
                  : 'Das Stichwahl-Fenster ist geschlossen.'
        };
    }
    return {
        canVote: false,
        roundKind: null,
        round: null,
        reason: 'Derzeit keine Stimmabgabe möglich.'
    };
}

/**
 * Öffentliche Sicht (ohne Admin-Felder).
 * @param {object} election
 */
function publicElectionView(election) {
    if (!election) return null;
    return {
        id: election.id,
        title: election.title,
        scopeLabel: election.scopeLabel,
        schoolYear: election.schoolYear,
        status: election.status,
        window: election.window,
        runoffWindow: election.runoffWindow,
        candidates: election.candidates,
        eligibleCount: election.eligibleCount,
        rounds: (election.rounds || []).map((r) => ({
            id: r.id,
            kind: r.kind,
            status: r.status,
            candidateIds: r.candidateIds,
            window: r.window,
            countedAt: r.countedAt,
            stats: r.stats
        })),
        offices: election.offices,
        publishedResults: !!election.publishedResults,
        updatedAt: election.updatedAt
    };
}

module.exports = {
    normId,
    validateMainBallot,
    validateRunoffBallot,
    votingPhase,
    publicElectionView
};
