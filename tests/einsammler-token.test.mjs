import { describe, expect, it } from 'vitest';
import {
    buildPublicPayload,
    buildPublicUrl,
    buildResponseSlip,
    decodePublicToken,
    decodeResponseSlip,
    encodePublicToken,
    encodeResponseSlip
} from '../src/tools/einsammler/einsammler-token.js';

describe('einsammler-token', () => {
    it('rundtrip public payload', () => {
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
        const token = encodePublicToken(payload);
        const decoded = decodePublicToken(token);
        expect(decoded.ok).toBe(true);
        expect(decoded.payload.aid).toBe('ant-1');
        expect(decoded.payload.titel).toBe('Foto');
        expect(buildPublicUrl('einsammler-antwort.html', token)).toContain('p=');
    });

    it('lehnt abgelaufene Links ab', () => {
        const payload = buildPublicPayload({
            einsammler: { id: 'ein-1', titel: 'X', status: 'aktiv' },
            antwort: { id: 'ant-1', schuelerName: 'A' },
            expDays: 1
        });
        payload.exp = Date.now() - 1000;
        const decoded = decodePublicToken(encodePublicToken(payload));
        expect(decoded.ok).toBe(false);
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
