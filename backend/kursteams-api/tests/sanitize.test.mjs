import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const { sanitizeEducationClassCode } = require('../src/lib/kursteam-create.js');

describe('sanitizeEducationClassCode', () => {
    it('entfernt Sonderzeichen und kürzt auf 50 Zeichen', () => {
        expect(
            sanitizeEducationClassCode({
                gruppenmail: 'jg2030-1hma',
                teamName: 'JG 20/30 1 HMA'
            })
        ).toBe('jg20301hma');
    });

    it('fallback auf teamName', () => {
        expect(sanitizeEducationClassCode({ teamName: 'Klasse 1A' })).toBe('Klasse1A');
    });
});
