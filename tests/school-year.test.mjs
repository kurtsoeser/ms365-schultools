import { describe, it, expect } from 'vitest';
import {
    schoolYearStartYear,
    currentSchoolYearLabel,
    nextSchoolYearLabel,
    parseSchoolYearStartYear,
    isSchoolYearLabel
} from '../src/shared/utils/school-year.js';

describe('school-year (Sep–Aug)', () => {
    it('März 2026 → Startjahr 2025', () => {
        expect(schoolYearStartYear(new Date('2026-03-15T12:00:00'))).toBe(2025);
        expect(currentSchoolYearLabel(new Date('2026-03-15T12:00:00'))).toBe('2025/26');
    });

    it('1. September 2026 → Startjahr 2026', () => {
        expect(schoolYearStartYear(new Date('2026-09-01T08:00:00'))).toBe(2026);
        expect(currentSchoolYearLabel(new Date('2026-09-01T08:00:00'))).toBe('2026/27');
    });

    it('31. August → noch Vorjahr', () => {
        expect(currentSchoolYearLabel(new Date('2026-08-31T23:00:00'))).toBe('2025/26');
    });

    it('nextSchoolYearLabel und parse', () => {
        expect(nextSchoolYearLabel('2025/26')).toBe('2026/27');
        expect(parseSchoolYearStartYear('2025/2026')).toBe(2025);
        expect(isSchoolYearLabel('2025/26')).toBe(true);
        expect(isSchoolYearLabel('foo')).toBe(false);
    });
});
