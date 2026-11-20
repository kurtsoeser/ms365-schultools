import { describe, it, expect } from 'vitest';
import {
    validateAktivitaet,
    rangesOverlap,
    daysBetween,
    toIsoDateOnly,
    klasseCodesOf,
    normalizeTeilnahme,
    appendAenderungsLog,
    computeDashboardKpis,
    initialStatusForSubmit
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-logic.js';
import {
    nextStatusAfterSave,
    canEditItem,
    canKvStellungnahme,
    sortByStartdatum
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-state.js';

describe('schulaktivitaeten-planer-logic', () => {
    it('toIsoDateOnly parses dates', () => {
        expect(toIsoDateOnly('2026-10-05')).toBe('2026-10-05');
        expect(toIsoDateOnly('2026-10-05T12:00:00Z')).toBe('2026-10-05');
    });

    it('toIsoDateOnly keeps calendar day for SharePoint dateOnly CET midnight', () => {
        // 30.11.2026 dateOnly → SPO oft 29.11. 23:00 UTC (Mitternacht Wien)
        expect(toIsoDateOnly('2026-11-29T23:00:00Z')).toBe('2026-11-30');
        expect(toIsoDateOnly('2026-12-03T23:00:00Z')).toBe('2026-12-04');
        expect(toIsoDateOnly('2026-11-30T00:00:00+01:00')).toBe('2026-11-30');
    });

    it('toIsoDateOnly keeps multi-day Wintersportwoche range', () => {
        const start = toIsoDateOnly('2026-11-29T23:00:00.0000000Z');
        const end = toIsoDateOnly('2026-12-03T23:00:00.0000000Z');
        expect(start).toBe('2026-11-30');
        expect(end).toBe('2026-12-04');
        expect(daysBetween(start, end)).toBe(4);
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

    it('supports multi-class and group participation', () => {
        const n = normalizeTeilnahme({
            teilnahmeModus: 'gruppe',
            klasseCodes: '3AK; 3BK',
            gruppenText: 'kath. Religion'
        });
        expect(n.klasseCode).toBe('3AK');
        expect(klasseCodesOf(n)).toEqual(['3AK', '3BK']);
        const r = validateAktivitaet({
            draft: {
                titel: 'Dom',
                typ: 'Exkursion',
                ...n,
                startdatum: '2026-11-18',
                enddatum: '2026-11-18',
                status: 'beantragt'
            },
            existing: [],
            rules: { minVorlaufTage: 7, maxGleichzeitigProKlasse: 1 },
            today: '2026-11-01'
        });
        expect(r.ok).toBe(true);
    });

    it('requires group name in gruppe mode', () => {
        const r = validateAktivitaet({
            draft: {
                titel: 'Dom',
                typ: 'Exkursion',
                teilnahmeModus: 'gruppe',
                klasseCodes: '3AK',
                startdatum: '2026-11-18',
                status: 'beantragt'
            },
            existing: [],
            rules: { minVorlaufTage: 0, maxGleichzeitigProKlasse: 1 },
            today: '2026-11-01'
        });
        expect(r.ok).toBe(false);
        expect(r.errors.some((e) => /Unterrichtsgruppe/i.test(e))).toBe(true);
    });

    it('tracks nachforderung and untis KPIs', () => {
        const kpi = computeDashboardKpis([
            { status: 'beantragt', startdatum: '2026-12-01' },
            { status: 'nachforderung', startdatum: '2026-12-02' },
            { status: 'genehmigt', startdatum: '2026-12-03', untisErledigt: false },
            { status: 'genehmigt', startdatum: '2026-12-04', untisErledigt: true }
        ]);
        expect(kpi.offen).toBe(1);
        expect(kpi.nachforderung).toBe(1);
        expect(kpi.untisOffen).toBe(1);
    });

    it('appendAenderungsLog keeps history', () => {
        const log = appendAenderungsLog('alt', 'neu', '2026-11-06T10:00:00Z');
        expect(log).toContain('alt');
        expect(log).toContain('neu');
    });
});

describe('schulaktivitaeten-planer-state workflow', () => {
    it('allows teacher edit on nachforderung and resubmits to beantragt', () => {
        const item = {
            status: 'nachforderung',
            lehrerEmail: 'a@schule.at',
            beantragtVon: 'a@schule.at'
        };
        expect(canEditItem(item, { role: 'lehrer', accountEmail: 'a@schule.at' })).toBe(true);
        expect(nextStatusAfterSave(item, { role: 'lehrer' })).toBe('beantragt');
    });

    it('resubmits nachforderung to kv_stellungnahme when path is kv_dann_direktion', () => {
        expect(
            nextStatusAfterSave(
                { status: 'nachforderung' },
                { role: 'lehrer' },
                { genehmigungPfad: 'kv_dann_direktion' }
            )
        ).toBe('kv_stellungnahme');
        expect(initialStatusForSubmit({ genehmigungPfad: 'kv_dann_direktion' })).toBe('kv_stellungnahme');
        expect(initialStatusForSubmit({ genehmigungPfad: 'direktion' })).toBe('beantragt');
    });

    it('keeps genehmigt when admin edits', () => {
        expect(nextStatusAfterSave({ status: 'genehmigt' }, { role: 'admin' })).toBe('genehmigt');
    });

    it('allows KV stellungnahme on waiting items', () => {
        const item = { status: 'kv_stellungnahme', klasseCodes: '3AK' };
        expect(
            canKvStellungnahme(item, {
                role: 'kv',
                accountEmail: 'kv@schule.at',
                stammdaten: { classes: [{ code: '3AK', headEmail: 'kv@schule.at' }] }
            })
        ).toBe(true);
        expect(
            canKvStellungnahme(item, {
                role: 'kv',
                demoRoleOverride: true,
                accountEmail: '',
                stammdaten: { classes: [] }
            })
        ).toBe(true);
        expect(
            canKvStellungnahme(item, { role: 'lehrer', accountEmail: '', stammdaten: { classes: [] } })
        ).toBe(false);
    });

    it('sorts Untis list by startdatum', () => {
        const sorted = sortByStartdatum([
            { titel: 'B', startdatum: '2026-12-01' },
            { titel: 'A', startdatum: '2026-11-01' }
        ]);
        expect(sorted.map((x) => x.titel)).toEqual(['A', 'B']);
    });

    it('treats short notice as warning when vorlaufModus=warnung', () => {
        const r = validateAktivitaet({
            draft: {
                titel: 'Kurzfristig',
                typ: 'Klassenabwesenheit',
                klasseCode: '3AK',
                startdatum: '2026-10-02',
                enddatum: '2026-10-02',
                status: 'beantragt'
            },
            existing: [],
            rules: { minVorlaufTage: 7, maxGleichzeitigProKlasse: 1, vorlaufModus: 'warnung' },
            today: '2026-10-01'
        });
        expect(r.ok).toBe(true);
        expect(r.warnings.some((w) => /Vorlauf/i.test(w))).toBe(true);
    });
});
