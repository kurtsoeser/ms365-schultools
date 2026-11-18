import { describe, expect, it } from 'vitest';
import {
    validateMainBallot,
    validateRunoffBallot,
    votingPhase,
    publicElectionView
} from '../src/lib/ballot.js';

describe('election-api ballot', () => {
    it('validiert Hauptwahlgang', () => {
        expect(validateMainBallot({ rankings: ['a', 'b', 'c'] }, ['a', 'b', 'c']).ok).toBe(true);
        expect(validateMainBallot({ rankings: ['a', 'a', 'b'] }, ['a', 'b']).ok).toBe(false);
    });

    it('validiert SGA-Lehrer-Hauptwahlgang (6 Reihungen)', () => {
        const ids = ['a', 'b', 'c', 'd', 'e', 'f', 'g'];
        const election = { electionKind: 'sga_lehrer' };
        expect(
            validateMainBallot({ rankings: ['a', 'b', 'c', 'd', 'e', 'f'] }, ids, election).ok
        ).toBe(true);
        expect(validateMainBallot({ rankings: ['a', 'b', 'c'] }, ids, election).ok).toBe(false);
    });

    it('validiert Schulsprecher mit SGA-Schüler-STV (6 Reihungen)', () => {
        const ids = ['a', 'b', 'c', 'd', 'e', 'f', 'g'];
        const election = { electionKind: 'schulsprecher', includeSgaDeputies: true };
        expect(
            validateMainBallot({ rankings: ['a', 'b', 'c', 'd', 'e', 'f'] }, ids, election).ok
        ).toBe(true);
        expect(validateMainBallot({ rankings: ['a', 'b', 'c'] }, ids, election).ok).toBe(false);
    });

    it('validiert Stichwahl', () => {
        expect(validateRunoffBallot({ choice: 'a' }, ['a', 'b']).ok).toBe(true);
        expect(validateRunoffBallot({ choice: 'x' }, ['a', 'b']).ok).toBe(false);
    });

    it('prüft Wahlfenster', () => {
        const election = {
            status: 'open',
            window: { start: '2020-01-01T00:00:00.000Z', end: '2099-01-01T00:00:00.000Z' },
            rounds: [{ id: 'main', kind: 'main', candidateIds: ['a'] }]
        };
        expect(votingPhase(election).canVote).toBe(true);
        expect(votingPhase({ ...election, status: 'draft' }).canVote).toBe(false);
    });

    it('blockiert Stichwahl bei SGA Lehrer', () => {
        const election = {
            electionKind: 'sga_lehrer',
            status: 'runoff_open',
            window: { start: '2020-01-01T00:00:00.000Z', end: '2099-01-01T00:00:00.000Z' },
            rounds: [{ id: 'runoff', kind: 'runoff', candidateIds: ['a', 'b'] }]
        };
        expect(votingPhase(election).canVote).toBe(false);
    });

    it('blendet Admin-Felder in public view nicht ein', () => {
        const view = publicElectionView({
            id: 'w1',
            title: 'Test',
            status: 'open',
            candidates: [],
            rounds: [],
            _adminOids: ['secret']
        });
        expect(view.id).toBe('w1');
        expect(view.electionKind).toBe('schulsprecher');
        expect(view._adminOids).toBeUndefined();
    });
});
