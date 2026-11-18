import { describe, expect, it } from 'vitest';
import {
    buildEmpfaengerLinesForKlassen,
    parseKlasseCodesField
} from '../src/tools/einsammler/einsammler-stammdaten.js';

describe('einsammler-stammdaten', () => {
    it('parst Klassen-Codes', () => {
        expect(parseKlasseCodesField('3A; 3B,4C')).toEqual(['3A', '3B', '4C']);
    });

    it('baut Empfängerzeilen aus Schülern', () => {
        const students = [
            {
                name: 'Anna',
                klasse: '3A',
                parents: [{ email: 'eltern@example.com', name: 'E' }]
            },
            { name: 'Ben', klasse: '3A', parents: [] },
            { name: 'Cara', klasse: '2B', parents: [{ email: 'x@y.at', name: 'X' }] }
        ];
        const lines = buildEmpfaengerLinesForKlassen(['3A'], students);
        expect(lines).toContain('Anna;3A;eltern@example.com');
        expect(lines).toContain('Ben;3A;');
        expect(lines.some((l) => l.includes('Cara'))).toBe(false);
    });

    it('legt je Eltern-Mail eine Empfängerzeile an', () => {
        const students = [
            {
                name: 'Anna',
                klasse: '3A',
                parents: [
                    { email: 'a@example.com', name: 'A' },
                    { email: 'b@example.com', name: 'B' }
                ]
            }
        ];
        const lines = buildEmpfaengerLinesForKlassen(['3A'], students);
        expect(lines).toEqual(['Anna;3A;a@example.com', 'Anna;3A;b@example.com']);
    });
});
