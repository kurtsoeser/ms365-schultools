import { describe, it, expect } from 'vitest';
import {
    validateAntrag,
    filterItems,
    itemCoversDay,
    formatDeDateTime,
    formatDisplayTextWithDeDates
} from '../src/tools/lehrer-freistellung-planer/lfr-logic.js';

describe('lfr-logic', () => {
    it('validateAntrag requires titel and dates', () => {
        const bad = validateAntrag({ titel: 'x', lehrerEmail: 'a@b.at' });
        expect(bad.ok).toBe(false);
        const ok = validateAntrag({
            titel: 'Fortbildung',
            beginn: '2026-03-01T08:00',
            ende: '2026-03-01T16:00',
            kategorie: 'Fortbildung',
            lehrerEmail: 'lehrer@schule.at'
        });
        expect(ok.ok).toBe(true);
    });

    it('filterItems scopes to account email', () => {
        const items = [
            { lehrerEmail: 'a@schule.at', status: 'Ausstehend' },
            { lehrerEmail: 'b@schule.at', status: 'Ausstehend' }
        ];
        expect(filterItems(items, {}, { scopeAll: false, accountEmail: 'a@schule.at' }).length).toBe(1);
        expect(filterItems(items, {}, { scopeAll: true, accountEmail: 'a@schule.at' }).length).toBe(2);
    });

    it('filterItems: ohne LehrerEmail keine fremden Altdaten in „Meine Anträge“', () => {
        const items = [
            { lehrerName: 'Kurt Söser', lehrerEmail: '', status: 'Genehmigt', titel: 'Fortbildung' },
            {
                lehrerName: 'Lehrer01 DEMO',
                lehrerEmail: 'lehrer01.demo@ms365.schule',
                status: 'Ausstehend',
                titel: 'Ski'
            }
        ];
        const mine = filterItems(items, {}, {
            scopeAll: false,
            accountEmail: 'lehrer01.demo@ms365.schule'
        });
        expect(mine).toHaveLength(1);
        expect(mine[0].titel).toBe('Ski');
        expect(
            filterItems(items, {}, { scopeAll: false, accountEmail: '' })
        ).toHaveLength(0);
    });

    it('filterItems: SharePoint-eigene Zeile und gleicher Anzeigename zählen', () => {
        const items = [
            { titel: 'Ski', lehrerEmail: '', lehrerName: 'Lehrer01 DEMO', callerVisible: true },
            { titel: 'Fremd', lehrerEmail: 'kurt@schule.at', lehrerName: 'Kurt Söser' }
        ];
        const mine = filterItems(items, {}, {
            scopeAll: false,
            accountEmail: 'lehrer01.demo@ms365.schule',
            accountName: 'Lehrer01 DEMO'
        });
        expect(mine.map((x) => x.titel)).toEqual(['Ski']);
    });

    it('filterItems: Author/createdBy zählt als eigen wenn LehrerEmail leer', () => {
        const items = [
            {
                titel: 'Ski',
                lehrerEmail: '',
                authorEmail: 'lehrer01.demo@ms365.schule',
                status: 'Ausstehend'
            },
            {
                titel: 'Fremd',
                lehrerEmail: '',
                authorEmail: 'kurt@schule.at',
                status: 'Genehmigt'
            }
        ];
        const mine = filterItems(items, {}, {
            scopeAll: false,
            accountEmail: 'lehrer01.demo@ms365.schule'
        });
        expect(mine.map((x) => x.titel)).toEqual(['Ski']);
    });

    it('itemCoversDay spans range', () => {
        expect(
            itemCoversDay({ beginn: '2026-03-01T08:00', ende: '2026-03-03T16:00' }, '2026-03-02')
        ).toBe(true);
    });

    it('formatDeDateTime zeigt lokale DE-Zeit ohne Z', () => {
        const s = formatDeDateTime('2026-11-20T08:30:00Z');
        expect(s).toMatch(/20\.11\.2026/);
        expect(s).toMatch(/Uhr/);
        expect(s).not.toMatch(/T|Z/);
    });

    it('formatDisplayTextWithDeDates ersetzt ISO in PA-Freitext', () => {
        const raw =
            'Genehmiger: Direktor Kommentar: ok Datum: 2026-11-19T17:55:21Z ---';
        const pretty = formatDisplayTextWithDeDates(raw);
        expect(pretty).toContain('19.11.2026');
        expect(pretty).not.toMatch(/T17:55:21Z/);
        expect(
            formatDisplayTextWithDeDates('GENEHMIGT: Kurt - 2026-11-20T08:30:00Z')
        ).not.toMatch(/T08:30:00Z/);
    });
});
