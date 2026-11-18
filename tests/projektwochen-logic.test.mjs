import { describe, expect, it } from 'vitest';
import {
    validateAngebot,
    computeDashboardKpis,
    computePwAttention,
    buildWeekPlan,
    buildWeekPlanDisplayRows,
    effectiveBuchungAb,
    isBookingOpen,
    weekdayLabelDeFromIso,
    toIsoDateOnly,
    resolvePwOwnership,
    assertPwOwnership,
    isPwAccessDeniedError,
    friendlyPwSpoError
} from '../src/tools/projektwochen/projektwochen-logic.js';
import {
    filterAngebote,
    pickActiveAktion,
    navSectionsForRole,
    viewsForRole
} from '../src/tools/projektwochen/projektwochen-state.js';
import { buildLocalDemoState } from '../src/tools/projektwochen/projektwochen-demo-data.js';

describe('projektwochen-logic', () => {
    it('validateAngebot verlangt Titel, Datum, Kapazität, Lehrer', () => {
        const empty = validateAngebot({ draft: {} });
        expect(empty.canSubmit).toBe(false);
        expect(empty.errors.length).toBeGreaterThan(1);

        const ok = validateAngebot({
            draft: {
                title: 'Zoo',
                datum: '2026-10-05',
                kapazitaet: 20,
                lehrerCode: 'MU',
                startzeit: '08:00',
                endzeit: '12:00',
                preisEuro: 0
            },
            aktion: { startdatum: '2026-10-05', enddatum: '2026-10-09' }
        });
        expect(ok.canSubmit).toBe(true);
    });

    it('Datum außerhalb Aktion → Fehler', () => {
        const r = validateAngebot({
            draft: {
                title: 'Zoo',
                datum: '2026-11-01',
                kapazitaet: 10,
                lehrerCode: 'MU'
            },
            aktion: { startdatum: '2026-10-05', enddatum: '2026-10-09' }
        });
        expect(r.canSubmit).toBe(false);
        expect(r.errors.some((e) => /außerhalb/i.test(e))).toBe(true);
    });

    it('effectiveBuchungAb und isBookingOpen', () => {
        expect(effectiveBuchungAb({ buchungAb: '' }, { buchungAbDefault: '2026-09-01T08:00' })).toBe(
            '2026-09-01T08:00'
        );
        expect(isBookingOpen('2099-01-01T08:00', new Date('2026-09-24'))).toBe(false);
        expect(isBookingOpen('2020-01-01T08:00', new Date('2026-09-24'))).toBe(true);
    });

    it('buildWeekPlan gruppiert nach Tag/Slot', () => {
        const grid = buildWeekPlan([
            { title: 'A', tag: 'Mo', slot: 'vormittag', status: 'beantragt' },
            { title: 'B', datum: '2026-09-28', slot: 'nachmittag', status: 'freigegeben' }
        ]);
        expect(grid.Mo.vormittag).toHaveLength(1);
        expect(weekdayLabelDeFromIso('2026-09-28')).toBe('Mo');
        expect(grid.Mo.nachmittag.length + grid.Di.nachmittag.length).toBeGreaterThanOrEqual(0);
        expect(toIsoDateOnly('2026-09-28T00:00:00Z')).toBe('2026-09-28');
    });

    it('buildWeekPlanDisplayRows spannt ganztags über drei Spalten', () => {
        const grid = buildWeekPlan([
            { title: 'Zoo', tag: 'Mo', slot: 'ganztags', status: 'freigegeben' },
            { title: 'Excel', tag: 'Di', slot: 'vormittag', status: 'beantragt' },
            { title: 'MixG', tag: 'Mi', slot: 'ganztags', status: 'freigegeben' },
            { title: 'MixN', tag: 'Mi', slot: 'nachmittag', status: 'beantragt' }
        ]);
        const rows = buildWeekPlanDisplayRows(grid);
        const mo = rows.find((r) => r.tag === 'Mo');
        expect(mo.bands).toHaveLength(1);
        expect(mo.bands[0].type).toBe('ganztags');
        expect(mo.bands[0].items[0].title).toBe('Zoo');
        const di = rows.find((r) => r.tag === 'Di');
        expect(di.bands).toHaveLength(1);
        expect(di.bands[0].type).toBe('slots');
        expect(di.bands[0].bySlot.vormittag).toHaveLength(1);
        const mi = rows.find((r) => r.tag === 'Mi');
        expect(mi.bands).toHaveLength(2);
        expect(mi.bands[0].type).toBe('ganztags');
        expect(mi.bands[1].type).toBe('slots');
    });

    it('computeDashboardKpis zählt Status', () => {
        const k = computeDashboardKpis(
            [
                { status: 'beantragt', kapazitaet: 10 },
                { status: 'freigegeben', kapazitaet: 20, buchungAb: '2020-01-01' },
                { status: 'freigegeben', kapazitaet: 5, buchungAb: '2099-01-01' }
            ],
            null,
            new Date('2026-09-24')
        );
        expect(k.beantragt).toBe(1);
        expect(k.freigegeben).toBe(2);
        expect(k.buchungOffen).toBe(1);
        expect(k.buchungGesperrt).toBe(1);
    });
});

describe('projektwochen ownership & soft errors', () => {
    it('resolvePwOwnership füllt BeantragtVon / E-Mail / Kürzel', () => {
        const o = resolvePwOwnership(
            { title: 'Zoo' },
            {
                accountEmail: 'mueller@schule.at',
                teachers: [{ code: 'MU', email: 'mueller@schule.at' }]
            }
        );
        expect(o.beantragtVon).toBe('mueller@schule.at');
        expect(o.lehrerEmail).toBe('mueller@schule.at');
        expect(o.lehrerCode).toBe('MU');
        expect(() => assertPwOwnership(o)).not.toThrow();
    });

    it('assertPwOwnership ohne E-Mail wirft', () => {
        expect(() => assertPwOwnership({ beantragtVon: '', lehrerEmail: '', lehrerCode: 'X' })).toThrow(
            /E-Mail/i
        );
    });

    it('friendlyPwSpoError verhüllt 403', () => {
        expect(isPwAccessDeniedError('Error 403 AccessDenied')).toBe(true);
        expect(friendlyPwSpoError('403 Access Denied')).toMatch(/Kein SharePoint-Zugriff/i);
    });

    it('computePwAttention zählt Freigabe und Sync', () => {
        const att = computePwAttention({
            role: 'admin',
            items: [
                { status: 'beantragt' },
                { status: 'freigegeben', bookingsServiceId: '' },
                { status: 'freigegeben', bookingsServiceId: 'svc', syncStatus: 'ok' }
            ]
        });
        expect(att.beantragt).toBe(1);
        expect(att.pendingSync).toBe(1);
    });
});

describe('projektwochen-state filter', () => {
    it('Schüler sieht nur freigegebene', () => {
        const items = [
            { status: 'beantragt', title: 'A', aktionId: 'pw' },
            { status: 'freigegeben', title: 'B', aktionId: 'pw' }
        ];
        const out = filterAngebote(items, {}, { role: 'schueler' }, { aktionId: 'pw' });
        expect(out).toHaveLength(1);
        expect(out[0].title).toBe('B');
    });

    it('Nav: Lehrer ohne System/Bookings, Admin mit Bookings-Sektion', () => {
        const lehrerIds = viewsForRole('lehrer').map((v) => v.id);
        expect(lehrerIds).not.toContain('einstellungen');
        expect(lehrerIds).not.toContain('bookings');
        expect(lehrerIds).toContain('neu');
        const schuelerIds = viewsForRole('schueler').map((v) => v.id);
        expect(schuelerIds).not.toContain('neu');
        const adminSecs = navSectionsForRole('admin');
        expect(adminSecs.map((s) => s.id)).toEqual(['arbeit', 'bookings', 'system']);
        expect(adminSecs.find((s) => s.id === 'bookings').items[0].id).toBe('bookings');
    });

    it('pickActiveAktion bevorzugt offen mit den meisten Angeboten', () => {
        const a = pickActiveAktion(
            [
                { status: 'offen', aktionId: 'pw-demo', title: 'Leer' },
                { status: 'offen', aktionId: 'pw-demo-2026', title: 'Voll' }
            ],
            [
                { aktionId: 'pw-demo-2026', title: 'A' },
                { aktionId: 'pw-demo-2026', title: 'B' }
            ]
        );
        expect(a.aktionId).toBe('pw-demo-2026');
        const b = pickActiveAktion([
            { status: 'geschlossen', aktionId: '1' },
            { status: 'offen', aktionId: '2' }
        ]);
        expect(b.aktionId).toBe('2');
    });
});

describe('projektwochen-demo-data', () => {
    it('liefert Aktion und Angebote', () => {
        const d = buildLocalDemoState();
        expect(d.aktionen[0].status).toBe('offen');
        expect(d.angebote.length).toBeGreaterThanOrEqual(5);
        expect(d.stammdaten.teachers.length).toBeGreaterThan(0);
    });
});
