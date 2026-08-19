import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const { normalizeDomain, domainFromEmail } = require('../src/lib/exchange-smtp.js');

describe('exchange-smtp helpers', () => {
    it('normalisiert Domain ohne @', () => {
        expect(normalizeDomain('@modeebensee.at')).toBe('modeebensee.at');
    });

    it('liest Domain aus E-Mail', () => {
        expect(domainFromEmail('kurt.soeser@modeebensee.at')).toBe('modeebensee.at');
    });
});
