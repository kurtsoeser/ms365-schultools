import { describe, expect, it } from 'vitest';
import { toBase64Url } from '../src/shared/parent-link-token.js';
import {
    buildPublicPayload,
    buildPublicUrl,
    buildResponseSlip,
    decodePublicToken,
    decodeResponseSlip,
    encodePublicToken,
    encodeResponseSlip,
    readParentLinkToken
} from '../src/tools/einsammler/einsammler-token.js';

const SECRET = 'unit-test-link-secret';

describe('einsammler-token', () => {
    it('rundtrip public payload (signiert)', () => {
        const payload = buildPublicPayload({
            einsammler: {
                id: 'ein-1',
                titel: 'Foto',
                textInhalt: 'Bitte zustimmen',
                modus: 'einwilligung',
                typ: 'foto',
                faelligAm: '2030-01-01',
                status: 'aktiv'
            },
            antwort: { id: 'ant-1', schuelerName: 'Anna', klasseCode: '3A' },
            schoolName: 'Testschule',
            expDays: 10
        });
        const token = encodePublicToken(payload, SECRET);
        expect(token).toContain('.');
        const decoded = decodePublicToken(token, { secret: SECRET, requireSignature: true });
        expect(decoded.ok).toBe(true);
        expect(decoded.payload.aid).toBe('ant-1');
        expect(decoded.payload.titel).toBe('Foto');
        expect(buildPublicUrl('einsammler-antwort.html', token)).toContain('#p=');
        expect(buildPublicUrl('einsammler-antwort.html', token)).not.toContain('?p=');
    });

    it('lehnt selbst gebauten unsignierten Token ab', () => {
        const payload = buildPublicPayload({
            einsammler: { id: 'ein-1', titel: 'X', status: 'aktiv' },
            antwort: { id: 'ant-1', schuelerName: 'A' },
            expDays: 10
        });
        const forged = toBase64Url(payload);
        const decoded = decodePublicToken(forged, { secret: SECRET, requireSignature: true });
        expect(decoded.ok).toBe(false);
        expect(String(decoded.error || '')).toMatch(/Signatur/i);
    });

    it('lehnt abgelaufene Links ab', () => {
        const payload = buildPublicPayload({
            einsammler: { id: 'ein-1', titel: 'X', status: 'aktiv' },
            antwort: { id: 'ant-1', schuelerName: 'A' },
            expDays: 1
        });
        payload.exp = Date.now() - 1000;
        const decoded = decodePublicToken(encodePublicToken(payload, SECRET), {
            secret: SECRET,
            requireSignature: true
        });
        expect(decoded.ok).toBe(false);
    });

    it('liest Token aus Fragment und Legacy-Query', () => {
        expect(readParentLinkToken({ hash: '#p=abc', search: '' })).toBe('abc');
        expect(readParentLinkToken({ hash: '', search: '?p=legacy' })).toBe('legacy');
        expect(readParentLinkToken('https://x.test/a.html#p=frag')).toBe('frag');
    });

    it('rundtrip response slip', () => {
        const slip = buildResponseSlip({
            eid: 'ein-1',
            aid: 'ant-1',
            status: 'zugestimmt',
            kommentar: 'ok',
            schuelerName: 'Anna'
        });
        const code = encodeResponseSlip(slip);
        const decoded = decodeResponseSlip(code);
        expect(decoded.ok).toBe(true);
        expect(decoded.slip.status).toBe('zugestimmt');
    });
});
