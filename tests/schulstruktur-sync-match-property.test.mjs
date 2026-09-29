import { describe, it, expect } from 'vitest';
import {
    keyAppearsAsToken,
    suggestTenantGroupForUnitFromList
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-match.js';

/** Einfacher deterministischer PRNG (xorshift). */
function makeRng(seed) {
    let s = seed >>> 0 || 1;
    return function next() {
        s ^= s << 13;
        s ^= s >>> 17;
        s ^= s << 5;
        return (s >>> 0) / 0xffffffff;
    };
}

function randomClassCode(rng) {
    const n = 1 + Math.floor(rng() * 12);
    const letter = String.fromCharCode(65 + Math.floor(rng() * 26));
    return String(n) + letter;
}

describe('Match Property-Tests (Klassenkürzel)', () => {
    it('keyAppearsAsToken: N niemals Token in (N*10+x)A', () => {
        for (let n = 1; n <= 9; n++) {
            const needle = String(n) + 'a';
            const hay = 'klasse ' + String(n) + '1a';
            expect(keyAppearsAsToken(hay, needle)).toBe(false);
        }
    });

    it('Zahlenfalle: 1–9 + Buchstabe matcht nicht 1x + Buchstabe', () => {
        const letters = 'ABCDEFGHJKLMNPRSTUVWXYZ';
        for (let n = 1; n <= 9; n++) {
            for (let li = 0; li < letters.length; li++) {
                const L = letters[li];
                const short = String(n) + L;
                const longNum = String(n) + '1' + L; // z. B. 1A → 11A
                const groups = [{ id: 'long', bezeichnung: 'Klasse ' + longNum, alias: longNum.toLowerCase() }];
                expect(suggestTenantGroupForUnitFromList({ bezeichnung: short }, groups)).toBe('');
            }
        }
    });

    it('Token-Match findet Klasse NA in „Klasse NA“ eindeutig', () => {
        const rng = makeRng(7);
        for (let i = 0; i < 40; i++) {
            const code = randomClassCode(rng);
            const groups = [
                { id: 'g', bezeichnung: 'Klasse ' + code, alias: code.toLowerCase() },
                { id: 'x', bezeichnung: 'Andere', alias: '' }
            ];
            expect(suggestTenantGroupForUnitFromList({ bezeichnung: code }, groups)).toBe('g');
        }
    });
});
