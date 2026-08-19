import { describe, expect, it } from 'vitest';
import {
    buildHygieneTargets,
    findClassTeamForClass,
    hygieneStatusForTarget,
    summarizeHygieneScan
} from '../src/shared/membership-hygiene.js';

describe('membership-hygiene', () => {
    it('buildHygieneTargets sammelt SLG, Verwaltung und Klassen', () => {
        const container = {
            setup: {
                matched: {
                    schuelerGroupId: 'g-s',
                    lehrerGroupId: 'g-l',
                    verwaltungGroupId: 'g-v'
                }
            },
            core: {
                classTeams: [{ classCode: '1A', graphGroupId: 'g-1a' }]
            }
        };
        const settings = {
            students: [{ klasse: '1A', email: 'a@s.at' }, { klasse: '2B', email: 'b@s.at' }],
            teachers: [{ code: 'KV', email: 't@s.at' }],
            admin: [{ role: 'Sek', email: 'sek@s.at' }],
            classes: [{ code: '1A', name: '1A', year: '2030' }]
        };
        const targets = buildHygieneTargets(container, settings);
        const ids = targets.map((t) => t.id);
        expect(ids).toContain('slg-schueler');
        expect(ids).toContain('slg-lehrer');
        expect(ids).toContain('verwaltung');
        expect(ids).toContain('klasse-1A');
        const klasse = targets.find((t) => t.id === 'klasse-1A');
        expect(klasse.listCount).toBe(1);
        expect(klasse.groupId).toBe('g-1a');
    });

    it('hygieneStatusForTarget erkennt Abweichungen', () => {
        expect(
            hygieneStatusForTarget({ groupId: 'g1', listCount: 10 }, 10)
        ).toBe('ok');
        expect(
            hygieneStatusForTarget({ groupId: 'g1', listCount: 10 }, 8)
        ).toBe('mismatch');
        expect(hygieneStatusForTarget({ groupId: '', listCount: 5 }, null)).toBe('unmatched');
    });

    it('findClassTeamForClass unterscheidet gleiche Kürzel nach Abschlussjahr', () => {
        const teams = [
            {
                classCode: '1HMA',
                abschlussJahr: '2029',
                graphGroupId: '',
                stableMailNickname: 'jg2029hma'
            },
            {
                classCode: '1HMA',
                abschlussJahr: '2031',
                graphGroupId: 'g-1hma-2031',
                stableMailNickname: 'jg2031hma'
            }
        ];
        const cls = { code: '1HMA', name: 'Klasse 1HMA', year: '2031' };
        const team = findClassTeamForClass(cls, teams);
        expect(team && team.graphGroupId).toBe('g-1hma-2031');
    });

    it('summarizeHygieneScan zählt Status korrekt', () => {
        const targets = [
            { id: 'a', groupId: 'g1', listCount: 2 },
            { id: 'b', groupId: 'g2', listCount: 3 },
            { id: 'c', groupId: null, listCount: 1 }
        ];
        const summary = summarizeHygieneScan(targets, { g1: 2, g2: 5 });
        expect(summary.counts.ok).toBe(1);
        expect(summary.counts.mismatch).toBe(1);
        expect(summary.counts.unmatched).toBe(1);
    });
});
