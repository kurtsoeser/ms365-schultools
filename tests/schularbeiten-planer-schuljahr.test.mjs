import { describe, it, expect } from 'vitest';
import {
    normalizeSchuljahr,
    currentSchoolYearFromDate,
    inferSchuljahrFromIsoDate,
    effectiveSchuljahrForItem,
    matchesSchuljahrFilter,
    pickActiveRegelwerk
} from '../src/tools/schularbeiten-planer/schularbeiten-planer-schuljahr.js';
import { LIST_TITLES, titlesForListKey } from '../src/tools/schularbeiten-planer/schularbeiten-planer-schema.js';

describe('schularbeiten-planer-schuljahr', () => {
    it('normalisiert Schuljahr-Strings', () => {
        expect(normalizeSchuljahr('2026/27')).toBe('2026/27');
        expect(normalizeSchuljahr('2026-27')).toBe('2026/27');
        expect(normalizeSchuljahr('2026.2027')).toBe('2026/27');
        expect(normalizeSchuljahr('')).toBe('');
    });

    it('leitet Schuljahr aus Datum ab (Start September)', () => {
        expect(inferSchuljahrFromIsoDate('2026-10-02')).toBe('2026/27');
        expect(inferSchuljahrFromIsoDate('2027-03-15')).toBe('2026/27');
        expect(inferSchuljahrFromIsoDate('2027-08-20')).toBe('2026/27');
        expect(inferSchuljahrFromIsoDate('2027-09-01')).toBe('2027/28');
    });

    it('filtert Legacy-Zeilen ohne Schuljahr-Spalte per Datum', () => {
        const item = { schuljahr: '', datum: '2026-11-10' };
        expect(effectiveSchuljahrForItem(item)).toBe('2026/27');
        expect(matchesSchuljahrFilter(item, '2026/27')).toBe(true);
        expect(matchesSchuljahrFilter(item, '2025/26')).toBe(false);
    });

    it('wählt aktives Regelwerk passend zum Schuljahr', () => {
        const rules = [
            { regelwerkId: 'a', aktiv: true, schuljahr: '2025/26' },
            { regelwerkId: 'b', aktiv: true, schuljahr: '2026/27' }
        ];
        expect(pickActiveRegelwerk(rules, '2026/27').regelwerkId).toBe('b');
    });

    it('aktuelles Schuljahr ab September', () => {
        expect(currentSchoolYearFromDate(new Date('2026-10-02'))).toBe('2026/27');
        expect(currentSchoolYearFromDate(new Date('2026-08-15'))).toBe('2025/26');
    });
});

describe('schularbeiten-planer-list-names', () => {
    it('verwendet SAP-Präfix und kennt Legacy-Namen', () => {
        expect(LIST_TITLES.schularbeiten).toBe('SAP-Schularbeiten');
        expect(titlesForListKey('schularbeiten')).toContain('Schularbeiten');
        expect(titlesForListKey('schularbeiten')[0]).toBe('SAP-Schularbeiten');
    });
});
