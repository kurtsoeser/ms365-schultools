import { describe, expect, it } from 'vitest';
import {
    validateMainBallot,
    validateRunoffBallot,
    tallyMainRound,
    tallyRunoffRound,
    resolveSchulsprecherFromMain,
    resolveDeputiesFromMain,
    resolveSgaLehrerFromMain,
    resolveSgaSchuelerDeputies,
    evaluateElection,
    POINTS_SCHULSPRECHER,
    POINTS_STV1,
    POINTS_STV2,
    sharePercent,
    turnoutPercent
} from '../src/tools/schulsprecherwahl/schulsprecherwahl-logic.js';
import {
    ELECTION_KIND_SGA_LEHRER,
    getElectionKindDef,
    normalizeElectionKind,
    resolveElectionBallotSchema
} from '../src/tools/schulsprecherwahl/schulsprecherwahl-kinds.js';
import { createElection, normalizeElection } from '../src/tools/schulsprecherwahl/schulsprecherwahl-model.js';

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

describe('schulwahlen kinds + sga lehrer', () => {
    const sgaIds = ['a', 'b', 'c', 'd', 'e', 'f', 'g'];

    it('normalisiert electionKind', () => {
        expect(normalizeElectionKind(undefined)).toBe('schulsprecher');
        expect(normalizeElectionKind('sga_lehrer')).toBe(ELECTION_KIND_SGA_LEHRER);
        expect(getElectionKindDef('sga_lehrer').rankingCount).toBe(6);
    });

    it('legt SGA-Lehrer-Wahl mit passenden Defaults an', () => {
        const e = createElection({ electionKind: 'sga_lehrer' });
        expect(e.electionKind).toBe('sga_lehrer');
        expect(e.eligibleMode).toBe('teachers');
        expect(e.title).toContain('SGA');
        expect(e.offices.member1Id).toBeNull();
        expect(e.offices.schulsprecherId).toBeUndefined();
    });

    it('migriert Legacy-Wahlen ohne electionKind', () => {
        const e = normalizeElection({ title: 'Alt', candidates: [] });
        expect(e.electionKind).toBe('schulsprecher');
        expect(e.offices.schulsprecherId).toBeNull();
    });

    it('validiert 6er-Reihung für SGA Lehrer', () => {
        const ok = validateMainBallot(
            { rankings: ['a', 'b', 'c', 'd', 'e', 'f'] },
            sgaIds,
            ELECTION_KIND_SGA_LEHRER
        );
        expect(ok.ok).toBe(true);
        const bad = validateMainBallot({ rankings: ['a', 'b', 'c'] }, sgaIds, ELECTION_KIND_SGA_LEHRER);
        expect(bad.ok).toBe(false);
    });

    it('ordnet Top 3 Mitglieder und Platz 4–6 Stellvertretung zu', () => {
        const ballots = [
            { rankings: ['a', 'b', 'c', 'd', 'e', 'f'] },
            { rankings: ['a', 'b', 'c', 'd', 'e', 'f'] }
        ];
        const evaled = evaluateElection({
            electionKind: ELECTION_KIND_SGA_LEHRER,
            mainBallots: ballots,
            candidateIds: sgaIds
        });
        expect(evaled.needsRunoff).toBe(false);
        expect(evaled.offices.member1Id).toBe('a');
        expect(evaled.offices.member2Id).toBe('b');
        expect(evaled.offices.member3Id).toBe('c');
        expect(evaled.offices.deputy1Id).toBe('d');
        expect(evaled.offices.deputy2Id).toBe('e');
        expect(evaled.offices.deputy3Id).toBe('f');
    });

    it('braucht kein Los, wenn Gleichstand nur innerhalb der drei Mitglieder liegt', () => {
        // a klar vorn; b und c gleich → beide Mitglieder, d klar darunter
        const tally = tallyMainRound(
            [
                { rankings: ['a', 'b', 'c', 'd', 'e', 'f'] },
                { rankings: ['a', 'c', 'b', 'd', 'e', 'f'] }
            ],
            sgaIds,
            ELECTION_KIND_SGA_LEHRER
        );
        expect(tally.byId.get('b').points).toBe(tally.byId.get('c').points);
        const resolved = resolveSgaLehrerFromMain(tally);
        expect(resolved.member1Id).toBe('a');
        expect([resolved.member2Id, resolved.member3Id].sort()).toEqual(['b', 'c']);
        expect(resolved.anyTieNeedsLot).toBe(false);
    });

    it('markiert Losbedarf, wenn mehr als drei Kandidat:innen für die Mitgliedersitze infrage kommen', () => {
        // a,b,c,d alle gleich viele Punkte → >3 für Mitglieder
        const ballots = [
            { rankings: ['a', 'b', 'c', 'd', 'e', 'f'] },
            { rankings: ['b', 'c', 'd', 'a', 'e', 'f'] },
            { rankings: ['c', 'd', 'a', 'b', 'e', 'f'] },
            { rankings: ['d', 'a', 'b', 'c', 'e', 'f'] }
        ];
        const tally = tallyMainRound(ballots, sgaIds, ELECTION_KIND_SGA_LEHRER);
        expect(tally.byId.get('a').points).toBe(tally.byId.get('d').points);
        const resolved = resolveSgaLehrerFromMain(tally);
        expect(resolved.tieNeedsLotMembers).toBe(true);
        expect(resolved.anyTieNeedsLot).toBe(true);
    });

    it('aktiviert 6er-Schema bei includeSgaDeputies', () => {
        const schema = resolveElectionBallotSchema({
            electionKind: 'schulsprecher',
            includeSgaDeputies: true
        });
        expect(schema.rankingCount).toBe(6);
        expect(schema.points).toEqual([6, 5, 4, 3, 2, 1]);
        expect(schema.includeSgaDeputies).toBe(true);
        expect(schema.seats).toHaveLength(6);
    });

    it('wertet Schulsprecher + SGA-Schüler-STV vom selben Stimmzettel aus', () => {
        const ids = ['a', 'b', 'c', 'd', 'e', 'f', 'g'];
        // a klar Schulsprecher (alle Erst); STV aus 6/5/4; SGA-STV aus 3/2/1 → d,e,f
        const ballots = [
            { rankings: ['a', 'b', 'c', 'd', 'e', 'f'] },
            { rankings: ['a', 'b', 'c', 'd', 'e', 'f'] },
            { rankings: ['a', 'c', 'b', 'd', 'f', 'e'] }
        ];
        const evaled = evaluateElection({
            electionKind: 'schulsprecher',
            includeSgaDeputies: true,
            mainBallots: ballots,
            candidateIds: ids
        });
        expect(evaled.offices.schulsprecherId).toBe('a');
        expect(evaled.offices.stv1Id).toBeTruthy();
        expect(evaled.offices.stv1Id).not.toBe('a');
        expect(evaled.offices.sgaDeputy1Id).toBe('d');
        expect(evaled.offices.sgaDeputy2Id).toBeTruthy();
        expect(evaled.offices.sgaDeputy3Id).toBeTruthy();
        // SGA-STV dürfen nicht die bereits gewählten SV-Ämter sein
        expect([
            evaled.offices.sgaDeputy1Id,
            evaled.offices.sgaDeputy2Id,
            evaled.offices.sgaDeputy3Id
        ]).not.toContain(evaled.offices.schulsprecherId);
        expect(validateMainBallot({ rankings: ['a', 'b', 'c'] }, ids, {
            electionKind: 'schulsprecher',
            includeSgaDeputies: true
        }).ok).toBe(false);
    });

    it('schließt gewählte SV-Ämter von SGA-Schüler-STV aus', () => {
        const tally = {
            byId: new Map([
                ['a', { id: 'a', sgaPoints: 9 }],
                ['b', { id: 'b', sgaPoints: 6 }],
                ['c', { id: 'c', sgaPoints: 3 }],
                ['d', { id: 'd', sgaPoints: 2 }]
            ])
        };
        const resolved = resolveSgaSchuelerDeputies(tally, ['a']);
        expect(resolved.sgaDeputy1Id).toBe('b');
        expect(resolved.sgaDeputy2Id).toBe('c');
        expect(resolved.sgaDeputy3Id).toBe('d');
    });

    it('speichert includeSgaDeputies im Model', () => {
        const e = createElection({ electionKind: 'schulsprecher', includeSgaDeputies: true });
        expect(e.includeSgaDeputies).toBe(true);
        expect(e.offices.sgaDeputy1Id).toBeNull();
    });
});
