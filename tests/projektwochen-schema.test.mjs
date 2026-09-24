import { describe, expect, it } from 'vitest';
import {
    LIST_TITLES,
    AKTIONEN_COLUMNS,
    ANGEBOTE_COLUMNS,
    REQUIRED_COLUMNS,
    newEntityId,
    nextProjectWeekRange,
    toGraphColumnBody
} from '../src/tools/projektwochen/projektwochen-schema.js';

describe('projektwochen-schema', () => {
    it('definiert PW-Aktionen und PW-Angebote', () => {
        expect(LIST_TITLES.aktionen).toBe('PW-Aktionen');
        expect(LIST_TITLES.angebote).toBe('PW-Angebote');
        expect(AKTIONEN_COLUMNS.some((c) => c.name === 'BuchungAbDefault')).toBe(true);
        expect(ANGEBOTE_COLUMNS.some((c) => c.name === 'BuchungAb')).toBe(true);
        expect(ANGEBOTE_COLUMNS.some((c) => c.name === 'HinweisEltern')).toBe(true);
        expect(REQUIRED_COLUMNS['PW-Aktionen'].length).toBe(AKTIONEN_COLUMNS.length);
        expect(REQUIRED_COLUMNS['PW-Angebote'].length).toBe(ANGEBOTE_COLUMNS.length);
    });

    it('toGraphColumnBody begrenzt Single-Line maxLength und setzt Mehrzeiler ohne maxLength', () => {
        const single = toGraphColumnBody({
            name: 'X',
            displayName: 'X',
            text: { allowMultipleLines: false, maxLength: 500 }
        });
        expect(single.text.maxLength).toBe(255);
        const multi = toGraphColumnBody({
            name: 'Y',
            displayName: 'Y',
            text: { allowMultipleLines: true, textType: 'plain' }
        });
        expect(multi.text.allowMultipleLines).toBe(true);
        expect(multi.text.maxLength).toBeUndefined();
        const num = toGraphColumnBody({
            name: 'Z',
            displayName: 'Z',
            number: {}
        });
        expect(num.number.decimalPlaces).toBe('automatic');
    });

    it('newEntityId erzeugt Präfix-IDs', () => {
        const id = newEntityId('ang');
        expect(id.startsWith('ang-')).toBe(true);
        expect(id.length).toBeGreaterThan(8);
    });

    it('nextProjectWeekRange liefert Mo–Fr und BuchungAb', () => {
        // Mittwoch 2026-09-23 → nächster Mo ist 2026-09-28
        const range = nextProjectWeekRange(new Date(2026, 8, 23));
        expect(range.startIso).toBe('2026-09-28');
        expect(range.endIso).toBe('2026-10-02');
        expect(range.buchungAbIso).toMatch(/^2026-09-14T/);
    });
});
