import { describe, it, expect } from 'vitest';
import {
    filterUnterrichtsteamRows,
    uniqueFieldValues,
    updateRowByKey,
    rowStableKey,
    hydrateUnterrichtsbelegungFromKursteamState,
    findUnterrichtsbelegungInOtherYears,
    applyAbgleichMatchesToRows
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

    it('hydratiert Belegung aus Kursteam teamsData', () => {
        const state = {
            teamsGenerated: true,
            teamsData: [
                {
                    teamName: 'SJ26 | 1A | D',
                    gruppenmail: 'sj26-1a-d',
                    besitzer: 't@school.at',
                    klasse: '1A',
                    fach: 'D',
                    isValid: true
                }
            ],
            yearPrefix: 'SJ26'
        };
        const { ok, imported, snapshot } = hydrateUnterrichtsbelegungFromKursteamState(null, state);
        expect(ok).toBe(true);
        expect(imported).toBe(1);
        expect(snapshot.rows[0].klasse).toBe('1A');
    });

    it('applyAbgleichMatchesToRows setzt graphGroupId per Mail oder Slot', () => {
        const rows = [
            { klasse: '1A', fach: 'D', lehrerCode: 'MAY', gruppenmail: 'demo-1a-d', graphGroupId: '' },
            { klasse: '2B', fach: 'M', lehrerCode: 'FOO', gruppenmail: '', graphGroupId: '' }
        ];
        const matched = [
            {
                matchBy: 'gruppenmail',
                klasse: '1A',
                fach: 'D',
                lehrerCode: 'MAY',
                gruppenmail: 'demo-1a-d',
                graphGroupId: 'g-1a-d'
            },
            {
                matchBy: 'unterricht',
                klasse: '2B',
                fach: 'M',
                lehrerCode: 'FOO',
                gruppenmail: 'demo-2b-m',
                graphGroupId: 'g-2b-m'
            }
        ];
        const { rows: next, linked } = applyAbgleichMatchesToRows(rows, matched);
        expect(linked).toBe(2);
        expect(next[0].graphGroupId).toBe('g-1a-d');
        expect(next[1].graphGroupId).toBe('g-2b-m');
        expect(next[1].gruppenmail).toBe('demo-2b-m');
    });

    it('findet Belegung in anderem Schuljahr', () => {
        const container = {
            years: {
                current: '2025/26',
                byLabel: {
                    '2025/26': { unterrichtsbelegung: null },
                    '2024/25': { unterrichtsbelegung: { rows: [{ klasse: '1A' }, { klasse: '2B' }] } }
                }
            }
        };
        expect(findUnterrichtsbelegungInOtherYears(container, '2025/26')).toEqual({
            year: '2024/25',
            count: 2
        });
    });
});
