import { describe, it, expect } from 'vitest';
import { loadScript } from './kursteams-vm.mjs';

describe('kursteam-filter-logic', () => {
    it('applyRowFilters: Fach-Ausschluss, Klasse erforderlich, Duplikate', () => {
        const ctx = loadScript('src/tools/kursteams/kursteam-filter-logic.js');
        const { applyRowFilters } = ctx.ms365KursteamFilterLogic;

        const raw = [
            { klasse: '1A', fach: 'D', lehrer: 'X', gruppe: '' },
            { klasse: '', fach: 'M', lehrer: 'Y', gruppe: '' },
            { klasse: '1A', fach: 'ORD', lehrer: 'Z', gruppe: '' },
            { klasse: '1A', fach: 'D', lehrer: 'X', gruppe: '' }
        ];

        const r1 = applyRowFilters(raw, ['ORD'], false);
        expect(r1.filtered.length).toBe(2);
        expect(r1.removedByFilter).toBe(2);
        expect(r1.removedByDuplicate).toBe(0);

        const r2 = applyRowFilters(raw, [], true);
        expect(r2.filtered.length).toBe(2);
        expect(r2.removedByDuplicate).toBe(1);
    });
});
