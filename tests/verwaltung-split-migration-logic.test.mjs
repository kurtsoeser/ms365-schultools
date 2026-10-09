import { describe, expect, it } from 'vitest';
import {
    assessVerwaltungSplitMigration,
    buildVerwaltungSplitExecutionPlan
} from '../src/tools/verwaltung/verwaltung-split-migration-logic.js';

describe('verwaltung-split-migration-logic', () => {
    it('zeigt Banner wenn Verwaltungsgruppe ohne Schulleitung-Gruppe', () => {
        const a = assessVerwaltungSplitMigration({
            verwaltungGroupId: '11111111-1111-1111-1111-111111111111',
            schulleitungGroupId: null,
            schulleitungEmails: ['dir@schule.at'],
            verwaltungEmails: ['sek@schule.at']
        });
        expect(a.phase).toBe('need_schulleitung_group');
        expect(a.showBanner).toBe(true);
    });

    it('Plan: Schulleitung aus Verwaltungsgruppe entfernen und Gruppen syncen', () => {
        const plan = buildVerwaltungSplitExecutionPlan(
            ['dir@schule.at'],
            ['sek@schule.at'],
            ['dir@schule.at', 'sek@schule.at', 'alt@schule.at'],
            []
        );
        expect(plan.schulleitung.join).toEqual(['dir@schule.at']);
        expect(plan.verwaltung.join).toEqual([]);
        expect(plan.verwaltung.leave).toContain('dir@schule.at');
        expect(plan.verwaltung.leave).toContain('alt@schule.at');
        expect(plan.schulleitungStillInVerwaltungGroup).toEqual(['dir@schule.at']);
    });

    it('kein Banner nach completedAt', () => {
        const a = assessVerwaltungSplitMigration({
            verwaltungGroupId: 'x',
            completedAt: '2026-01-01T00:00:00.000Z'
        });
        expect(a.showBanner).toBe(false);
        expect(a.phase).toBe('done');
    });
});
