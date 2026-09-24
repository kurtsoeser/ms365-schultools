import { describe, it, expect } from 'vitest';
import {
    DEFAULT_RULES,
    toIsoDateOnly,
    daysBetween,
    isoWeekKey,
    addDays,
    dateInWindow,
    validateSchularbeit,
    computeDashboardKpis,
    buildWeeklyDistribution
} from '../src/tools/schularbeiten-planer/schularbeiten-planer-logic.js';
import { newEntityId, DEFAULT_REGELWERK_FIELDS } from '../src/tools/schularbeiten-planer/schularbeiten-planer-schema.js';

describe('schularbeiten-planer-logic helpers', () => {
    it('toIsoDateOnly aus Date und String', () => {
        expect(toIsoDateOnly('2026-10-14')).toBe('2026-10-14');
        expect(toIsoDateOnly('2026-10-14T12:00:00Z')).toBe('2026-10-14');
        expect(toIsoDateOnly(null)).toBe(null);
    });

    it('daysBetween und addDays', () => {
        expect(daysBetween('2026-10-01', '2026-10-08')).toBe(7);
        expect(addDays('2026-10-01', -1)).toBe('2026-09-30');
    });

    it('isoWeekKey stabil', () => {
        expect(isoWeekKey('2026-10-14')).toMatch(/^2026-W\d{2}$/);
        expect(isoWeekKey('2026-10-14')).toBe(isoWeekKey('2026-10-16'));
    });

    it('dateInWindow inklusiv', () => {
        const win = { typ: 'gesperrt', startdatum: '2026-10-26', enddatum: '2026-10-31', titel: 'Herbst' };
        expect(dateInWindow('2026-10-26', win)).toBe(true);
        expect(dateInWindow('2026-10-31', win)).toBe(true);
        expect(dateInWindow('2026-11-01', win)).toBe(false);
    });
});

describe('validateSchularbeit', () => {
    const baseDraft = {
        schularbeitId: 'sa-new',
        fachCode: 'D',
        klasseCode: '3AK',
        lehrerCode: 'BAU',
        datum: '2026-11-10',
        dauerMinuten: 100,
        semester: 'WS',
        status: 'beantragt'
    };

    const windows = [
        { titel: 'Herbstferien', typ: 'gesperrt', startdatum: '2026-10-26', enddatum: '2026-10-31' }
    ];

    it('akzeptiert gültigen Antrag mit genug Vorlauf', () => {
        const r = validateSchularbeit({
            draft: baseDraft,
            existing: [],
            rules: DEFAULT_RULES,
            windows,
            today: '2026-10-01'
        });
        expect(r.errors).toEqual([]);
        expect(r.canSubmit).toBe(true);
    });

    it('blockiert bei zu kurzer Ankündigungsfrist', () => {
        const r = validateSchularbeit({
            draft: baseDraft,
            existing: [],
            rules: DEFAULT_RULES,
            windows: [],
            today: '2026-11-05'
        });
        expect(r.canSubmit).toBe(false);
        expect(r.errors.some((e) => /Ankündigungsfrist/i.test(e))).toBe(true);
    });

    it('blockiert zweiten Termin am selben Tag', () => {
        const r = validateSchularbeit({
            draft: baseDraft,
            existing: [
                {
                    schularbeitId: 'sa-1',
                    fachCode: 'M',
                    klasseCode: '3AK',
                    datum: '2026-11-10',
                    status: 'fixiert',
                    dauerMinuten: 50
                }
            ],
            rules: DEFAULT_RULES,
            windows: [],
            today: '2026-10-01'
        });
        expect(r.canSubmit).toBe(false);
        expect(r.errors.some((e) => /pro Tag/i.test(e))).toBe(true);
    });

    it('blockiert bei mehr als maxProWoche', () => {
        const r = validateSchularbeit({
            draft: { ...baseDraft, datum: '2026-11-13' },
            existing: [
                {
                    schularbeitId: 'sa-1',
                    klasseCode: '3AK',
                    fachCode: 'M',
                    datum: '2026-11-10',
                    status: 'fixiert',
                    dauerMinuten: 50
                },
                {
                    schularbeitId: 'sa-2',
                    klasseCode: '3AK',
                    fachCode: 'E',
                    datum: '2026-11-12',
                    status: 'beantragt',
                    dauerMinuten: 50
                }
            ],
            rules: DEFAULT_RULES,
            windows: [],
            today: '2026-10-01'
        });
        expect(r.canSubmit).toBe(false);
        expect(r.errors.some((e) => /pro Kalenderwoche/i.test(e))).toBe(true);
    });

    it('blockiert Sperrzeit', () => {
        const r = validateSchularbeit({
            draft: { ...baseDraft, datum: '2026-10-28' },
            existing: [],
            rules: DEFAULT_RULES,
            windows,
            today: '2026-10-01'
        });
        expect(r.canSubmit).toBe(false);
        expect(r.errors.some((e) => /Sperrzeit/i.test(e))).toBe(true);
    });

    it('warnt nach Sperrzeit-Tag', () => {
        const r = validateSchularbeit({
            draft: { ...baseDraft, datum: '2026-11-01' },
            existing: [],
            rules: DEFAULT_RULES,
            windows,
            today: '2026-10-01'
        });
        expect(r.canSubmit).toBe(true);
        expect(r.warnings.some((w) => /schulfrei/i.test(w))).toBe(true);
    });

    it('ignoriert abgelehnte bei Kontingent-Zählung', () => {
        const r = validateSchularbeit({
            draft: baseDraft,
            existing: [
                {
                    schularbeitId: 'sa-x',
                    klasseCode: '3AK',
                    fachCode: 'M',
                    datum: '2026-11-10',
                    status: 'abgelehnt',
                    dauerMinuten: 50
                }
            ],
            rules: DEFAULT_RULES,
            windows: [],
            today: '2026-10-01'
        });
        expect(r.canSubmit).toBe(true);
    });
});

describe('computeDashboardKpis + schema', () => {
    it('zählt offene und diese Woche', () => {
        const k = computeDashboardKpis(
            [
                { status: 'beantragt', datum: '2026-09-23' },
                { status: 'fixiert', datum: '2026-09-24' },
                { status: 'fixiert', datum: '2026-12-01' }
            ],
            '2026-09-23'
        );
        expect(k.offen).toBe(1);
        expect(k.dieseWoche).toBeGreaterThanOrEqual(1);
        expect(k.fixiertNaechste2Wochen).toBeGreaterThanOrEqual(1);
    });

    it('zählt Konflikte bei Doppelbelegung am Tag', () => {
        const k = computeDashboardKpis(
            [
                {
                    status: 'fixiert',
                    klasseCode: '3AK',
                    fachCode: 'D',
                    datum: '2026-11-10',
                    dauerMinuten: 100
                },
                {
                    status: 'beantragt',
                    klasseCode: '3AK',
                    fachCode: 'M',
                    datum: '2026-11-10',
                    dauerMinuten: 50
                }
            ],
            '2026-10-01',
            DEFAULT_RULES
        );
        expect(k.konflikte).toBeGreaterThanOrEqual(1);
    });

    it('newEntityId und Default-Regelwerk', () => {
        expect(newEntityId('sa')).toMatch(/^sa-[a-z0-9]+$/i);
        expect(DEFAULT_REGELWERK_FIELDS.MaxProTag).toBe(1);
        expect(DEFAULT_REGELWERK_FIELDS.Aktiv).toBe(true);
    });
});

describe('buildWeeklyDistribution', () => {
    it('stapelt Fächer pro Woche', () => {
        const today = '2026-09-23';
        const week = isoWeekKey(today);
        const dist = buildWeeklyDistribution(
            [
                { status: 'fixiert', fachCode: 'D', datum: today, klasseCode: '1AK' },
                { status: 'beantragt', fachCode: 'M', datum: today, klasseCode: '1AK' },
                { status: 'abgelehnt', fachCode: 'E', datum: today, klasseCode: '1AK' }
            ],
            { today, weekCount: 4 }
        );
        expect(dist.faecher).toContain('D');
        expect(dist.faecher).toContain('M');
        expect(dist.faecher).not.toContain('E');
        const w = dist.weeks.find((x) => x.key === week);
        expect(w).toBeTruthy();
        expect(w.total).toBe(2);
        expect(w.byFach.D).toBe(1);
        expect(w.byFach.M).toBe(1);
    });
});
