import { describe, it, expect } from 'vitest';
import {
    filterUnterrichtsteamRows,
    uniqueFieldValues,
    updateRowByKey,
    rowStableKey
} from '../src/tools/unterrichtsteams-katalog/unterrichtsteams-katalog-logic.js';

describe('unterrichtsteams-katalog-logic', () => {
    const sample = [
        { klasse: '1A', fach: 'BW', lehrerCode: 'FRECH', gruppenmail: 'a', graphGroupId: 'g1' },
        { klasse: '1AK', fach: 'BESPK', lehrerCode: 'RINNE', gruppenmail: 'b', graphGroupId: '' }
    ];

    it('filtert nach Klasse und M365', () => {
        const f = filterUnterrichtsteamRows(sample, { klasse: '1A', linkedOnly: true });
        expect(f).toHaveLength(1);
        expect(f[0].fach).toBe('BW');
    });

    it('liefert Filteroptionen', () => {
        expect(uniqueFieldValues(sample, 'fach')).toEqual(['BESPK', 'BW']);
    });

    it('aktualisiert Zeile per Key', () => {
        const key = rowStableKey(sample[0]);
        const { rows, ok } = updateRowByKey(sample, key, { teamName: 'SJ26-27 | 1A | BW-FRECH' });
        expect(ok).toBe(true);
        expect(rows[0].teamName).toBe('SJ26-27 | 1A | BW-FRECH');
    });
});
