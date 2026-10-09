import { describe, expect, it } from 'vitest';
import {
    hygieneTargetMetricLine,
    hygieneTargetNumericHint,
    registerSnapshotLine,
    resolveToolTaskRowFacts
} from '../src/shared/dashboard-task-row-facts-logic.js';

describe('dashboard-task-row-facts-logic', () => {
    it('registerSnapshotLine fasst Register-Zahlen zusammen', () => {
        const line = registerSnapshotLine({
            teachers: [{}, {}],
            students: [{}, {}, {}],
            classes: [{}, {}],
            subjects: [{}]
        });
        expect(line).toContain('2 Lehrkräfte');
        expect(line).toContain('3 Schüler:innen');
        expect(line).toContain('2 Klassen');
    });

    it('hygieneTargetMetricLine trennt Zahlen und Status', () => {
        const hygieneApi = {
            buildHygieneTargets: function () {
                return [{ id: 'slg-schueler', listCount: 26, groupId: 'g1' }];
            },
            loadHygieneScanCache: function () {
                return { rows: [{ id: 'slg-schueler', groupCount: 26 }] };
            },
            hygieneStatusDashboardTone: function () {
                return 'ok';
            }
        };
        expect(hygieneTargetMetricLine('slg-schueler', 'ok', null, {}, hygieneApi)).toBe(
            '26 in Stammdaten · 26 in M365'
        );
        expect(hygieneTargetNumericHint('slg-schueler', 'ok', null, {}, hygieneApi)).toContain(
            'Konsistent'
        );
    });

    it('Personen-Kachel nutzt Register-Zahlen', () => {
        const facts = resolveToolTaskRowFacts({
            toolId: 'personen-verwaltung',
            href: 'tools/personen-verwaltung.html',
            settings: { teachers: [{}], students: [{}, {}] }
        });
        expect(facts && facts.desc).toMatch(/1 Lehrkräfte/);
        expect(facts && facts.chip).toBe('');
    });
});
