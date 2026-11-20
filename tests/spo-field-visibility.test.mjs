import { describe, expect, it } from 'vitest';
import { isAlwaysHideShowFormula } from '../src/shared/spo-field-visibility.js';
import {
    buildAlwaysHideFormula,
    buildTitleHideFormula,
    buildSeminarShowFormula
} from '../src/tools/srdp-anmeldung/srdp-anmeldung-formatting.js';

describe('isAlwaysHideShowFormula', () => {
    it('erkennt dauerhaftes Ausblenden', () => {
        expect(isAlwaysHideShowFormula('=false')).toBe(true);
        expect(isAlwaysHideShowFormula(' = false ')).toBe(true);
        expect(isAlwaysHideShowFormula(buildTitleHideFormula())).toBe(true);
        expect(isAlwaysHideShowFormula(buildAlwaysHideFormula())).toBe(true);
    });

    it('lässt echte Bedingungsformeln durch', () => {
        expect(isAlwaysHideShowFormula(buildSeminarShowFormula())).toBe(false);
        expect(isAlwaysHideShowFormula("=if([$Variante]=='Variante 1','true','false')")).toBe(false);
        expect(isAlwaysHideShowFormula('')).toBe(false);
    });
});
