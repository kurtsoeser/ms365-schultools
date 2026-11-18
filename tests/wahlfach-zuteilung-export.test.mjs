import { describe, expect, it } from 'vitest';
import { aoaToCsv, buildErgebnisExportAoa } from '../src/tools/wahlfach-zuteilung/wahlfach-zuteilung-export.js';

describe('wahlfach-zuteilung-export', () => {
    it('baut CSV-Zeilen für Access-Import', () => {
        const state = {
            runde: { schuljahr: '2026/27', titel: 'Wahlfächer' },
            angebote: [
                {
                    id: 'a1',
                    bezeichnung: 'Chor',
                    kapazitaetMin: 8,
                    kapazitaetMax: 20,
                    aktiv: true
                }
            ],
            anmeldungen: [
                {
                    id: 'n1',
                    schuelerName: 'Max',
                    klasse: '3AK',
                    wunsch1Id: 'a1',
                    wunsch2Id: '',
                    wunsch3Id: '',
                    status: 'zugeordnet'
                }
            ],
            zuteilungen: [
                {
                    anmeldungId: 'n1',
                    angebotId: 'a1',
                    zugeteiltAm: '2026-11-06T10:00:00Z',
                    zugeteiltVon: 'SLG'
                }
            ]
        };
        const aoa = buildErgebnisExportAoa(state);
        expect(aoa[0][0]).toBe('Schuljahr');
        expect(aoa.length).toBeGreaterThan(1);
        expect(aoa[1][2]).toBe('Chor');
        expect(aoa[1][7]).toBe('Max');
        const csv = aoaToCsv(aoa);
        expect(csv).toContain('Chor');
        expect(csv).toContain(';');
    });
});
