import { describe, it, expect } from 'vitest';
import {
    toIsoDateOnly,
    inclusiveDayCount,
    isMultiDay,
    approvalPath,
    validateFreistellung,
    computeDashboardKpis,
    filterFreistellungen
} from '../src/tools/freistellung-planer/freistellung-planer-logic.js';
import { itemVisibleForRole } from '../src/tools/freistellung-planer/freistellung-planer-state.js';

describe('freistellung-planer-logic', () => {
    it('toIsoDateOnly parses dates', () => {
        expect(toIsoDateOnly('2026-10-05T12:00:00Z')).toBe('2026-10-05');
    });

    it('inclusiveDayCount counts both ends', () => {
        expect(inclusiveDayCount('2026-10-01', '2026-10-01')).toBe(1);
        expect(inclusiveDayCount('2026-10-01', '2026-10-03')).toBe(3);
    });

    it('isMultiDay from 2 days', () => {
        expect(isMultiDay('2026-10-01', '2026-10-01')).toBe(false);
        expect(isMultiDay('2026-10-01', '2026-10-02')).toBe(true);
    });

    it('approvalPath: 1 Tag nur KV, mehrtägig KV+Direktion', () => {
        const one = approvalPath('2026-10-01', '2026-10-01');
        expect(one.multiDay).toBe(false);
        expect(one.steps).toEqual(['Klassenvorstand']);

        const multi = approvalPath('2026-10-01', '2026-10-03');
        expect(multi.multiDay).toBe(true);
        expect(multi.steps).toEqual(['Klassenvorstand', 'Direktion']);
    });

    it('validateFreistellung requires name, klasse, kv', () => {
        const r = validateFreistellung({
            draft: {
                beginn: '2026-10-10',
                ende: '2026-10-10',
                kategorie: 'Sonstiges'
            },
            today: '2026-10-01'
        });
        expect(r.ok).toBe(false);
        expect(r.errors.length).toBeGreaterThan(0);
    });

    it('validateFreistellung accepts complete draft', () => {
        const r = validateFreistellung({
            draft: {
                schuelerName: 'Max',
                klasse: '3AHW',
                beginn: '2026-10-10',
                ende: '2026-10-12',
                kategorie: 'Ärztlicher Termin',
                kvEmail: 'kv@schule.at',
                beschreibung: 'Kontrolle'
            },
            today: '2026-10-01'
        });
        expect(r.ok).toBe(true);
        expect(r.path.multiDay).toBe(true);
    });

    it('computeDashboardKpis counts status', () => {
        const kpi = computeDashboardKpis([
            { status: 'Ausstehend', beginn: '2026-10-05', ende: '2026-10-05' },
            { status: 'Genehmigt', beginn: '2026-10-08', ende: '2026-10-10' },
            { status: 'Abgelehnt', beginn: '2026-11-01', ende: '2026-11-01' }
        ], '2026-10-01');
        expect(kpi.ausstehend).toBe(1);
        expect(kpi.genehmigt).toBe(1);
        expect(kpi.abgelehnt).toBe(1);
        expect(kpi.mehrtage).toBe(1);
    });

    it('filterFreistellungen supports multiDay and mine', () => {
        const items = [
            {
                status: 'Ausstehend',
                klasse: '3AHW',
                beginn: '2026-10-01',
                ende: '2026-10-01',
                authorEmail: 'a@schule.at',
                kvEmail: 'kv@schule.at'
            },
            {
                status: 'Ausstehend',
                klasse: '4AHW',
                beginn: '2026-10-01',
                ende: '2026-10-03',
                authorEmail: 'b@schule.at',
                kvEmail: 'kv@schule.at'
            }
        ];
        expect(filterFreistellungen(items, { multiDay: '1' }).length).toBe(1);
        expect(
            filterFreistellungen(items, {}, { onlyMine: true, accountEmail: 'a@schule.at' }).length
        ).toBe(1);
    });

    it('itemVisibleForRole schränkt Schüler und KV ein', () => {
        const row = {
            authorEmail: 'a@schule.at',
            kvEmail: 'kv@schule.at',
            status: 'Ausstehend'
        };
        const schuelerState = { role: 'schueler', accountEmail: 'a@schule.at' };
        const otherSchueler = { role: 'schueler', accountEmail: 'x@schule.at' };
        const kvState = { role: 'kv', accountEmail: 'kv@schule.at' };
        expect(itemVisibleForRole(schuelerState, row)).toBe(true);
        expect(itemVisibleForRole(otherSchueler, row)).toBe(false);
        expect(itemVisibleForRole(kvState, row)).toBe(true);
        expect(itemVisibleForRole({ role: 'direktion', accountEmail: 'dir@schule.at' }, row)).toBe(true);
    });
});
