import { createHmac } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import {
    buildFragmentPublicUrl,
    hmacSha256Base64Url,
    readParentLinkToken,
    signTokenBody,
    verifySignedToken
} from '../src/shared/parent-link-token.js';

describe('parent-link-token', () => {
    it('HMAC stimmt mit node:crypto überein', () => {
        const secret = 's3cret';
        const body = 'eyJ2IjoxfQ';
        const expected = createHmac('sha256', secret).update(body, 'utf8').digest('base64url');
        expect(hmacSha256Base64Url(secret, body)).toBe(expected);
    });

    it('signiert und verifiziert; lehnt Fälschung ab', () => {
        const secret = 'school-link-secret';
        const token = signTokenBody({ v: 1, eid: 'e1', aid: 'a1', titel: 'T' }, secret);
        expect(verifySignedToken(token, secret).ok).toBe(true);
        expect(verifySignedToken(token.split('.')[0], secret).ok).toBe(false);
        expect(verifySignedToken(token.split('.')[0] + '.nope', secret).ok).toBe(false);
    });

    it('baut Fragment-URL und liest Token', () => {
        const url = buildFragmentPublicUrl('einsammler-antwort.html', 'tok.sig');
        expect(url).toBe('einsammler-antwort.html#p=tok.sig');
        expect(readParentLinkToken(url)).toBe('tok.sig');
    });
});
