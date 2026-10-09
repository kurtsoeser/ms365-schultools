import { describe, it, expect } from 'vitest';
import { validateAntrag, filterItems, itemCoversDay } from '../src/tools/lehrer-freistellung-planer/lfr-logic.js';

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

    it('itemCoversDay spans range', () => {
        expect(
            itemCoversDay({ beginn: '2026-03-01T08:00', ende: '2026-03-03T16:00' }, '2026-03-02')
        ).toBe(true);
    });
});
