import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const { createRateLimiter } = require('../backend/einsammler-api/src/lib/rate-limit.js');

describe('einsammler-api rate-limit', () => {
    it('blockiert nach max Treffern', () => {
        const lim = createRateLimiter({ windowMs: 60_000, max: 3 });
        expect(lim.check('ip-1').ok).toBe(true);
        expect(lim.check('ip-1').ok).toBe(true);
        expect(lim.check('ip-1').ok).toBe(true);
        expect(lim.check('ip-1').ok).toBe(false);
        expect(lim.check('ip-2').ok).toBe(true);
    });
});
