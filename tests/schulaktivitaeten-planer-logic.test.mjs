import { describe, it, expect } from 'vitest';
import {
    validateAktivitaet,
    rangesOverlap,
    daysBetween,
    toIsoDateOnly
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-logic.js';

describe('schulaktivitaeten-planer-logic', () => {
    it('toIsoDateOnly parses dates', () => {
        expect(toIsoDateOnly('2026-10-05T12:00:00Z')).toBe('2026-10-05');
    });

    it('rangesOverlap detects inclusive overlap', () => {
        expect(rangesOverlap('2026-10-01', '2026-10-03', '2026-10-03', '2026-10-05')).toBe(true);
        expect(rangesOverlap('2026-10-01', '2026-10-02', '2026-10-03', '2026-10-05')).toBe(false);
    });

    it('daysBetween counts calendar days', () => {
        expect(daysBetween('2026-10-01', '2026-10-08')).toBe(7);
    });

    it('rejects short notice', () => {
        const r = validateAktivitaet({
            draft: {
                titel: 'Museum',
                typ: 'Exkursion',
                klasseCode: '3AK',
                startdatum: '2026-10-02',
                enddatum: '2026-10-02',
                status: 'beantragt'
            },
            existing: [],
            rules: { minVorlaufTage: 7, maxGleichzeitigProKlasse: 1 },
            today: '2026-10-01'
        });
        expect(r.ok).toBe(false);
        expect(r.errors.some((e) => /Vorlauf/i.test(e))).toBe(true);
    });

    it('accepts valid draft with enough lead time', () => {
        const r = validateAktivitaet({
            draft: {
                titel: 'Museum',
                typ: 'Exkursion',
                klasseCode: '3AK',
                ort: 'Wien',
                startdatum: '2026-10-15',
                enddatum: '2026-10-15',
                status: 'beantragt'
            },
            existing: [],
            rules: { minVorlaufTage: 7, maxGleichzeitigProKlasse: 1 },
            today: '2026-10-01'
        });
        expect(r.ok).toBe(true);
    });
});
