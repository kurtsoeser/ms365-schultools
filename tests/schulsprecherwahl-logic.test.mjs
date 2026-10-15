import { describe, expect, it } from 'vitest';
import {
    validateMainBallot,
    validateRunoffBallot,
    tallyMainRound,
    tallyRunoffRound,
    resolveSchulsprecherFromMain,
    resolveDeputiesFromMain,
    evaluateElection,
    POINTS_SCHULSPRECHER,
    POINTS_STV1,
    POINTS_STV2,
    sharePercent,
    turnoutPercent
} from '../src/tools/schulsprecherwahl/schulsprecherwahl-logic.js';

describe('schulsprecherwahl-logic ballots', () => {
    it('akzeptiert gültige 6/5/4-Reihung', () => {
        const r = validateMainBallot({ rankings: ['a', 'b', 'c'] }, ['a', 'b', 'c', 'd']);
        expect(r.ok).toBe(true);
        expect(r.rankings).toEqual(['a', 'b', 'c']);
    });

    it('lehnt doppelte Kandidaten ab', () => {
        const r = validateMainBallot({ rankings: ['a', 'a', 'b'] }, ['a', 'b']);
        expect(r.ok).toBe(false);
    });

    it('validiert Stichwahl-Stimme', () => {
        expect(validateRunoffBallot({ choice: 'a' }, ['a', 'b']).ok).toBe(true);
        expect(validateRunoffBallot({ choice: 'c' }, ['a', 'b']).ok).toBe(false);
    });
});

describe('schulsprecherwahl-logic auszählung', () => {
    const ids = ['emin', 'matthias', 'laya', 'sara', 'jonas'];

    it('vergibt Wahlpunkte 6/5/4', () => {
        const tally = tallyMainRound(
            [
                { rankings: ['emin', 'matthias', 'laya'] },
                { rankings: ['emin', 'laya', 'matthias'] },
                { rankings: ['matthias', 'emin', 'laya'] }
            ],
            ids
        );
        expect(tally.valid).toBe(3);
        expect(tally.byId.get('emin').points).toBe(POINTS_SCHULSPRECHER * 2 + POINTS_STV1);
        expect(tally.byId.get('emin').firstPlace).toBe(2);
        expect(tally.byId.get('matthias').firstPlace).toBe(1);
    });

    it('wählt Schulsprecher bei absoluter Mehrheit der Erstplatzierungen', () => {
        const ballots = [
            { rankings: ['emin', 'matthias', 'laya'] },
            { rankings: ['emin', 'laya', 'sara'] },
            { rankings: ['emin', 'sara', 'jonas'] },
            { rankings: ['matthias', 'emin', 'laya'] }
        ];
        const evaled = evaluateElection({ mainBallots: ballots, candidateIds: ids });
        expect(evaled.offices.schulsprecherId).toBe('emin');
        expect(evaled.offices.decidedAfter).toBe('main');
        expect(evaled.needsRunoff).toBe(false);
        expect(evaled.offices.stv1Id).toBeTruthy();
        expect(evaled.offices.stv2Id).toBeTruthy();
        expect(evaled.offices.stv1Id).not.toBe('emin');
        expect(evaled.offices.stv2Id).not.toBe('emin');
    });

    it('ordnet Stichwahl an ohne absolute Mehrheit', () => {
        const ballots = [
            { rankings: ['emin', 'laya', 'sara'] },
            { rankings: ['emin', 'laya', 'jonas'] },
            { rankings: ['matthias', 'laya', 'sara'] },
            { rankings: ['matthias', 'sara', 'jonas'] },
            { rankings: ['laya', 'sara', 'jonas'] }
        ];
        // 5 gültig → Mehrheit 3; emin/matthias je 2 Erst, laya 1 → Stichwahl emin vs matthias
        const main = tallyMainRound(ballots, ids);
        const first = resolveSchulsprecherFromMain(main);
        expect(first.decided).toBe(false);
        expect(first.needsRunoff).toBe(true);
        expect(first.runoffCandidateIds).toHaveLength(2);
        expect(first.runoffCandidateIds).toContain('emin');
        expect(first.runoffCandidateIds).toContain('matthias');
    });

    it('bestimmt Stellvertreter aus Punkten ohne Schulsprecher', () => {
        const tally = tallyMainRound(
            [
                { rankings: ['emin', 'matthias', 'laya'] },
                { rankings: ['emin', 'laya', 'matthias'] },
                { rankings: ['emin', 'matthias', 'sara'] }
            ],
            ids
        );
        const dep = resolveDeputiesFromMain(tally, 'emin');
        expect(dep.stv1Id).toBe('matthias');
        expect(dep.stv2Id).toBe('laya');
    });

    it('entscheidet Stichwahl per einfacher Mehrheit', () => {
        const mainBallots = [
            { rankings: ['emin', 'matthias', 'laya'] },
            { rankings: ['matthias', 'emin', 'laya'] }
        ];
        const runoffBallots = [
            { choice: 'emin' },
            { choice: 'emin' },
            { choice: 'matthias' }
        ];
        const evaled = evaluateElection({
            mainBallots,
            candidateIds: ids,
            runoffBallots,
            runoffCandidateIds: ['emin', 'matthias']
        });
        expect(evaled.offices.schulsprecherId).toBe('emin');
        expect(evaled.offices.decidedAfter).toBe('runoff');
        // STV bleiben aus 1. Wahlgang
        expect(evaled.offices.stv1Id).toBeTruthy();
        expect(evaled.offices.stv1Id).not.toBe('emin');
    });

    it('zählt Stichwahl-Stimmen', () => {
        const t = tallyRunoffRound(
            [{ choice: 'a' }, { choice: 'b' }, { choice: 'a' }, { choice: 'x' }],
            ['a', 'b']
        );
        expect(t.valid).toBe(3);
        expect(t.invalid).toBe(1);
        expect(t.byId.get('a').votes).toBe(2);
    });

    it('berechnet Beteiligung und anteile', () => {
        expect(turnoutPercent(587, 673)).toBe(87.2);
        expect(sharePercent(373, 587)).toBe(63.5);
    });
});
