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
        expect(view._adminOids).toBeUndefined();
    });
});
