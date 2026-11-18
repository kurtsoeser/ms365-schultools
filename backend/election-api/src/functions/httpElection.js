'use strict';

const { app } = require('@azure/functions');
const {
    jsonResponse,
    corsPreflightResponse,
    bearerTokenFromRequest,
    readJsonBody,
    errorJsonResponse
} = require('../lib/http-utils');
const { validateCallerToken } = require('../lib/validate-token');
const {
    listElections,
    getElection,
    upsertElection,
    deleteElection,
    hasVoted,
    castAnonymousVote,
    listBallots,
    countVoters,
    isElectionAdmin
} = require('../lib/tables');
const {
    validateMainBallot,
    validateRunoffBallot,
    votingPhase,
    publicElectionView,
    normId
} = require('../lib/ballot');

function stripAdmin(election) {
    if (!election) return null;
    const copy = Object.assign({}, election);
    delete copy._adminOids;
    return copy;
}

async function requireAuth(request) {
    return validateCallerToken(bearerTokenFromRequest(request));
}

app.http('httpElectionCors', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'election/{*rest}',
    handler: async () => corsPreflightResponse()
});

/** GET /api/election – Liste für Mandant */
app.http('httpElectionList', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'election',
    handler: async (request, context) => {
        try {
            const claims = await requireAuth(request);
            const list = await listElections(claims.tid);
            return jsonResponse(200, {
                elections: list.map((e) => ({
                    ...publicElectionView(e),
                    isAdmin: isElectionAdmin(e, claims.oid)
                }))
            });
        } catch (e) {
            return errorJsonResponse(context, 'election list', e);
        }
    }
});

/** GET /api/election/{id} */
app.http('httpElectionGet', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'election/{electionId}',
    handler: async (request, context) => {
        try {
            const electionId = normId(request.params.electionId);
            // Falls die statische Route election/health nicht greift
            if (electionId === 'health') {
                const { ensureTables } = require('../lib/tables');
                const { getConfig } = require('../lib/config');
                getConfig();
                await ensureTables();
                return jsonResponse(200, { ok: true, storage: 'azure-tables' });
            }
            const claims = await requireAuth(request);
            const election = await getElection(claims.tid, electionId);
            if (!election) {
                return jsonResponse(404, { error: 'Wahl nicht gefunden.', code: 'not_found' });
            }
            const admin = isElectionAdmin(election, claims.oid);
            const view = admin ? stripAdmin(election) : publicElectionView(election);
            return jsonResponse(200, {
                election: view,
                isAdmin: admin
            });
        } catch (e) {
            return errorJsonResponse(context, 'election get', e);
        }
    }
});

/** PUT /api/election/{id} – Admin upsert */
app.http('httpElectionPut', {
    methods: ['PUT'],
    authLevel: 'anonymous',
    route: 'election/{electionId}',
    handler: async (request, context) => {
        try {
            const claims = await requireAuth(request);
            const electionId = normId(request.params.electionId);
            const body = await readJsonBody(request);
            const existing = await getElection(claims.tid, electionId);

            let adminOids = existing && Array.isArray(existing._adminOids) ? existing._adminOids.slice() : [];
            if (!existing) {
                adminOids = [claims.oid];
            } else if (!isElectionAdmin(existing, claims.oid)) {
                const err = new Error('Keine Berechtigung, diese Wahl zu bearbeiten.');
                err.status = 403;
                err.code = 'forbidden';
                throw err;
            }

            if (Array.isArray(body.adminOids) && isElectionAdmin(existing || { _adminOids: adminOids }, claims.oid)) {
                const next = body.adminOids.map((x) => String(x).toLowerCase()).filter(Boolean);
                if (next.length) adminOids = next;
                if (!adminOids.includes(claims.oid)) adminOids.push(claims.oid);
            }

            const election = Object.assign({}, body.election || body, {
                id: electionId,
                createdAt: (existing && existing.createdAt) || new Date().toISOString()
            });
            delete election._adminOids;
            delete election.adminOids;

            const saved = await upsertElection(election, claims.tid, adminOids);
            return jsonResponse(200, { election: stripAdmin(Object.assign(saved, { _adminOids: adminOids })), isAdmin: true });
        } catch (e) {
            return errorJsonResponse(context, 'election put', e);
        }
    }
});

/** DELETE /api/election/{id} */
app.http('httpElectionDelete', {
    methods: ['DELETE'],
    authLevel: 'anonymous',
    route: 'election/{electionId}',
    handler: async (request, context) => {
        try {
            const claims = await requireAuth(request);
            const electionId = normId(request.params.electionId);
            const existing = await getElection(claims.tid, electionId);
            if (!existing) {
                return jsonResponse(404, { error: 'Wahl nicht gefunden.', code: 'not_found' });
            }
            if (!isElectionAdmin(existing, claims.oid)) {
                const err = new Error('Keine Berechtigung.');
                err.status = 403;
                err.code = 'forbidden';
                throw err;
            }
            await deleteElection(claims.tid, electionId);
            return jsonResponse(200, { ok: true });
        } catch (e) {
            return errorJsonResponse(context, 'election delete', e);
        }
    }
});

/** GET /api/election/{id}/me – hat gewählt? */
app.http('httpElectionMe', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'election/{electionId}/me',
    handler: async (request, context) => {
        try {
            const claims = await requireAuth(request);
            const electionId = normId(request.params.electionId);
            const election = await getElection(claims.tid, electionId);
            if (!election) {
                return jsonResponse(404, { error: 'Wahl nicht gefunden.', code: 'not_found' });
            }
            const phase = votingPhase(election);
            let voted = false;
            if (phase.round) {
                voted = await hasVoted(claims.tid, electionId, phase.round.id, claims.oid);
            }
            return jsonResponse(200, {
                canVote: phase.canVote && !voted,
                voted,
                reason: voted ? 'Bereits abgestimmt.' : phase.reason,
                roundKind: phase.roundKind,
                roundId: phase.round ? phase.round.id : null,
                election: publicElectionView(election)
            });
        } catch (e) {
            return errorJsonResponse(context, 'election me', e);
        }
    }
});

/** POST /api/election/{id}/vote */
app.http('httpElectionVote', {
    methods: ['POST'],
    authLevel: 'anonymous',
    route: 'election/{electionId}/vote',
    handler: async (request, context) => {
        try {
            const claims = await requireAuth(request);
            const electionId = normId(request.params.electionId);
            const election = await getElection(claims.tid, electionId);
            if (!election) {
                return jsonResponse(404, { error: 'Wahl nicht gefunden.', code: 'not_found' });
            }
            const phase = votingPhase(election);
            if (!phase.canVote || !phase.round) {
                const err = new Error(phase.reason || 'Keine Stimmabgabe möglich.');
                err.status = 403;
                err.code = 'not_open';
                throw err;
            }

            const body = await readJsonBody(request);
            const allow = new Set(
                (phase.round.candidateIds && phase.round.candidateIds.length
                    ? phase.round.candidateIds
                    : (election.candidates || []).map((c) => c.id)
                ).map(normId)
            );

            /** @type {{ rankings?: string[], choice?: string }} */
            let ballotPayload;
            if (phase.roundKind === 'main') {
                const check = validateMainBallot(body, allow, election);
                if (!check.ok) {
                    const err = new Error(check.reason);
                    err.status = 400;
                    err.code = 'invalid_ballot';
                    throw err;
                }
                ballotPayload = { rankings: check.rankings };
            } else {
                const check = validateRunoffBallot(body, allow);
                if (!check.ok) {
                    const err = new Error(check.reason);
                    err.status = 400;
                    err.code = 'invalid_ballot';
                    throw err;
                }
                ballotPayload = { choice: check.choice };
            }

            const result = await castAnonymousVote({
                tid: claims.tid,
                electionId,
                roundId: phase.round.id,
                oid: claims.oid,
                ballotPayload
            });
            return jsonResponse(200, result);
        } catch (e) {
            return errorJsonResponse(context, 'election vote', e);
        }
    }
});

/** GET /api/election/{id}/ballots?roundId= – nur Admin */
app.http('httpElectionBallots', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'election/{electionId}/ballots',
    handler: async (request, context) => {
        try {
            const claims = await requireAuth(request);
            const electionId = normId(request.params.electionId);
            const election = await getElection(claims.tid, electionId);
            if (!election) {
                return jsonResponse(404, { error: 'Wahl nicht gefunden.', code: 'not_found' });
            }
            if (!isElectionAdmin(election, claims.oid)) {
                const err = new Error('Keine Berechtigung.');
                err.status = 403;
                err.code = 'forbidden';
                throw err;
            }
            const url = new URL(request.url);
            const roundId = normId(url.searchParams.get('roundId'));
            if (!roundId) {
                return jsonResponse(400, { error: 'roundId fehlt.', code: 'missing_round' });
            }
            const ballots = await listBallots(claims.tid, electionId, roundId);
            const voters = await countVoters(claims.tid, electionId, roundId);
            return jsonResponse(200, { ballots, voterCount: voters, roundId });
        } catch (e) {
            return errorJsonResponse(context, 'election ballots', e);
        }
    }
});
