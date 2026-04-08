import { describe, it, expect } from 'vitest';
import { loadScript } from './kursteams-vm.mjs';

describe('kursteam-manual-view-logic', () => {
    it('applyManualFiltersAndSort: Filter und Sortierung', () => {
        const ctx = loadScript('src/tools/kursteams/kursteam-manual-view-logic.js');
        const { applyManualFiltersAndSort } = ctx.ms365KursteamManualViewLogic;

        const filteredData = [
            { id: 1, klasse: '2B', fach: 'D', lehrer: 'AB', gruppe: '' },
            { id: 2, klasse: '1A', fach: 'M', lehrer: 'CD', gruppe: '' }
        ];
        const manualSort = { key: 'klasse', dir: 1 };
        const filters = { klasse: '1', fach: '', lehrer: '' };
        const view = applyManualFiltersAndSort(filteredData, manualSort, filters);
        expect(view.length).toBe(1);
        expect(view[0].row.klasse).toBe('1A');
    });
});
