import { describe, it, expect } from 'vitest';
import { formatReadableDateRange } from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-detail-ui.js';

describe('formatReadableDateRange', () => {
    it('formats single day with weekday', () => {
        expect(formatReadableDateRange('2026-11-26')).toBe('Do, 26. November 2026');
    });

    it('formats multi-day in same month compactly', () => {
        expect(formatReadableDateRange('2026-11-26', '2026-11-28')).toBe('26.–28. November 2026');
    });

    it('treats equal end as single day', () => {
        expect(formatReadableDateRange('2026-11-26', '2026-11-26')).toBe('Do, 26. November 2026');
    });
});
