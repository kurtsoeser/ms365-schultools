import { createHmac } from 'node:crypto';
import { createRequire } from 'node:module';
import { afterEach, describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const TOKEN_PATH = '../backend/einsammler-api/src/lib/token.js';

function loadTokenMod() {
    const resolved = require.resolve(TOKEN_PATH);
    delete require.cache[resolved];
    return require(TOKEN_PATH);
}

describe('einsammler-api token (R2/a5)', () => {
    afterEach(() => {
        delete process.env.EINSAMMLER_LINK_SECRET;
        delete require.cache[require.resolve(TOKEN_PATH)];
    });

    it('ohne Secret: API abgeschaltet, Token abgelehnt', () => {
        delete process.env.EINSAMMLER_LINK_SECRET;
        const tokenMod = loadTokenMod();
        expect(tokenMod.publicApiEnabled()).toBe(false);
        const body = Buffer.from(
            JSON.stringify({ v: 1, eid: 'e', aid: 'a', titel: 'T', exp: Date.now() + 99999 }),
            'utf8'
        ).toString('base64url');
        expect(tokenMod.decodePublicToken(body).ok).toBe(false);
        expect(String(tokenMod.decodePublicToken(body).error || '')).toMatch(/nicht freigeschaltet|Secret/i);
    });

    it('mit Secret: unsigniert abgelehnt, signiert ok', () => {
        process.env.EINSAMMLER_LINK_SECRET = 'api-test-secret';
        const tokenMod = loadTokenMod();
        expect(tokenMod.publicApiEnabled()).toBe(true);
        const payload = { v: 1, eid: 'e', aid: 'a', titel: 'T', exp: Date.now() + 99999 };
        const body = Buffer.from(JSON.stringify(payload), 'utf8').toString('base64url');
        expect(tokenMod.decodePublicToken(body).ok).toBe(false);
        const sig = createHmac('sha256', 'api-test-secret').update(body, 'utf8').digest('base64url');
        const ok = tokenMod.decodePublicToken(body + '.' + sig);
        expect(ok.ok).toBe(true);
        expect(ok.payload.aid).toBe('a');
    });
});
