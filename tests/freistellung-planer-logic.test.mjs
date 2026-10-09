import { describe, it, expect } from 'vitest';
import {
    toIsoDateOnly,
    toIsoDateTimeLocal,
    toSharePointDateTime,
    hasClockTime,
    toDateTimeMs,
    inclusiveDayCount,
    isMultiDay,
    approvalPath,
    validateFreistellung,
    computeDashboardKpis,
    filterFreistellungen,
    freistellungMatchesKvScope,
    deriveKvClassCodesFromFreistellungItems,
    normalizeFreistellungClassCode,
    personNamesLooselyMatch,
    itemCoversDay,
    monthGridDates
} from '../src/tools/freistellung-planer/freistellung-planer-logic.js';
import { itemVisibleForRole } from '../src/tools/freistellung-planer/freistellung-planer-state.js';

describe('freistellung-planer-logic', () => {
    it('toIsoDateOnly parses dates', () => {
        expect(toIsoDateOnly('2026-10-05T12:00:00Z')).toBe('2026-10-05');
    });

    it('toIsoDateTimeLocal und Stunden-Vergleich', () => {
        expect(toIsoDateTimeLocal('2026-10-06T08:30')).toBe('2026-10-06T08:30');
        expect(toIsoDateTimeLocal('2026-10-06')).toBe('2026-10-06T00:00');
        expect(toSharePointDateTime('2026-10-06T08:30')).toBe('2026-10-06T08:30:00');
        expect(hasClockTime('2026-10-06T08:00')).toBe(true);
        expect(hasClockTime('2026-10-06T00:00')).toBe(false);
        expect(toDateTimeMs('2026-10-06T09:00')).toBeGreaterThan(toDateTimeMs('2026-10-06T08:00'));
        expect(isMultiDay('2026-10-06T08:00', '2026-10-06T12:00')).toBe(false);
        expect(approvalPath('2026-10-06T08:00', '2026-10-06T10:00').steps).toEqual(['Klassenvorstand']);
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

    it('personNamesLooselyMatch: Vorname/Nachname unabhängig von Reihenfolge', () => {
        expect(personNamesLooselyMatch('Brian May', 'brian may')).toBe(true);
        expect(personNamesLooselyMatch('May, Brian', 'Brian May')).toBe(true);
    });

    it('deriveKvClassCodesFromFreistellungItems: 1A wenn Brian KV auf einem Antrag', () => {
        const codes = deriveKvClassCodesFromFreistellungItems(
            [
                { klasse: '1A', kvName: 'Brian May', kvEmail: '' },
                { klasse: '1A', kvName: '', kvEmail: '', schuelerName: 'Andere' },
                { klasse: '3A', kvName: '', kvEmail: '' }
            ],
            'brian.may@kurtrocks.com',
            'Brian May'
        );
        expect(codes.has('1A')).toBe(true);
        expect(normalizeFreistellungClassCode('DEMO Klasse 1A')).toBe('1A');
        const visible = filterFreistellungen(
            [
                { klasse: '1A', kvName: '', kvEmail: '' },
                { klasse: '3A', kvName: '', kvEmail: '' }
            ],
            {},
            {
                onlyKv: true,
                accountEmail: 'brian.may@kurtrocks.com',
                accountName: 'Brian May',
                kvClassCodes: codes
            }
        );
        expect(visible.length).toBe(1);
        expect(visible[0].klasse).toBe('1A');
    });

    it('freistellungMatchesKvScope: KV-Name aus Personenfeld', () => {
        expect(
            freistellungMatchesKvScope(
                { kvName: 'Brian May', klasse: '1A' },
                { accountEmail: 'brian@schule.at', accountName: 'Brian May', onlyKv: true }
            )
        ).toBe(true);
    });

    it('filterFreistellungen: KV sieht Klasse auch ohne kvEmail im Antrag', () => {
        const items = [
            { klasse: '1A', kvEmail: '', authorEmail: 's@schule.at' },
            { klasse: '2B', kvEmail: 'other@schule.at', authorEmail: 'x@schule.at' }
        ];
        const codes = new Set(['1A']);
        expect(
            filterFreistellungen(items, {}, { onlyKv: true, accountEmail: 'kv@schule.at', kvClassCodes: codes })
                .length
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

    it('itemCoversDay und monthGridDates für Kalender', () => {
        expect(
            itemCoversDay({ beginn: '2026-03-01T08:00', ende: '2026-03-03T16:00' }, '2026-03-02')
        ).toBe(true);
        expect(itemCoversDay({ beginn: '2026-03-01', ende: '2026-03-01' }, '2026-03-02')).toBe(false);
        expect(monthGridDates(2026, 3).length).toBe(42);
    });
});
