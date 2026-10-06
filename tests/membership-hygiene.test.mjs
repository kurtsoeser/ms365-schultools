import { describe, expect, it } from 'vitest';
import {
    buildHygieneTargets,
    findClassTeamForClass,
    resolveClassGraphGroupId,
    countLinkedClassTeamsForClasses,
    hygieneStatusForTarget,
    aggregateHygieneStatuses,
    hygieneStatusDashboardTone,
    hygieneStatusDashboardHint,
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

    it('countLinkedClassTeamsForClasses zählt nur Stammdaten-Klassen, nicht verwaiste Teams', () => {
        const classes = [
            { code: '1A', year: '2032' },
            { code: '1B', year: '2032' },
            { code: '2A', year: '2031' },
            { code: '3A', year: '2030' },
            { code: '4A', year: '2029' }
        ];
        const classTeams = [
            { classCode: '1A', abschlussJahr: '2032', graphGroupId: 'g-1a' },
            { classCode: '1B', abschlussJahr: '2032', graphGroupId: 'g-1b' },
            { classCode: '2A', abschlussJahr: '2031', graphGroupId: 'g-2a' },
            { classCode: '3A', abschlussJahr: '2030', graphGroupId: 'g-3a' },
            { classCode: '4A', abschlussJahr: '2029', graphGroupId: 'g-4a' },
            { classCode: '5A', abschlussJahr: '2028', graphGroupId: 'g-old-5a' },
            { classCode: '0A', abschlussJahr: '2030', graphGroupId: 'g-orphan' },
            { classCode: 'X', graphGroupId: 'g-x' }
        ];
        const r = countLinkedClassTeamsForClasses(classes, classTeams);
        expect(r.total).toBe(5);
        expect(r.linked).toBe(5);
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

    it('resolveClassGraphGroupId nutzt classGroupMatchByKey wenn classTeams leer', () => {
        const cls = { code: '2B', year: '2031' };
        expect(resolveClassGraphGroupId(cls, [], {})).toBe('');
        expect(
            resolveClassGraphGroupId(cls, [], {
                '2B': { groupId: 'g-2b', notFound: false }
            })
        ).toBe('g-2b');
        expect(
            resolveClassGraphGroupId(cls, [{ classCode: '2B', abschlussJahr: '2031', graphGroupId: 'g-team' }], {
                '2B': { groupId: 'g-map' }
            })
        ).toBe('g-team');
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

    it('aggregateHygieneStatuses und Dashboard-Töne', () => {
        expect(aggregateHygieneStatuses(['ok', 'ok'])).toBe('ok');
        expect(aggregateHygieneStatuses(['ok', 'mismatch'])).toBe('mismatch');
        expect(aggregateHygieneStatuses(['unmatched', 'unmatched'])).toBe('unmatched');
        expect(aggregateHygieneStatuses(['ok', 'unmatched'])).toBe('mismatch');
        expect(hygieneStatusDashboardTone('ok')).toBe('ok');
        expect(hygieneStatusDashboardTone('mismatch')).toBe('warn');
        expect(hygieneStatusDashboardTone('unknown')).toBe('pending');
        expect(hygieneStatusDashboardHint('ok')).toBe('Konsistent');
    });
});
