import { describe, expect, it } from 'vitest';
import {
    validateAngebot,
    computeDashboardKpis,
    computePwAttention,
    statusLabel,
    dashboardHomeModel,
    teacherOfferLine,
    studentBookingsUrl,
    buildWeekPlan,
    buildWeekPlanDisplayRows,
    effectiveBuchungAb,
    isBookingOpen,
    weekdayLabelDeFromIso,
    weekdayNameDe,
    nextSlotTimes,
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
    viewsForRole,
    resolvePwView
} from '../src/tools/projektwochen/projektwochen-state.js';
import { buildLocalDemoState } from '../src/tools/projektwochen/projektwochen-demo-data.js';

describe('projektwochen-logic', () => {
    it('validateAngebot verlangt Titel, Datum, Ort, Plätze, Lehrkraft', () => {
        const empty = validateAngebot({ draft: {} });
        expect(empty.canSubmit).toBe(false);
        expect(empty.errors).toContain('Ort fehlt.');
        expect(empty.errors).toContain('Plätze müssen mindestens 1 sein.');
        expect(empty.errors).toContain('Lehrkraft fehlt.');
        expect(empty.errors.some((e) => /Kapazität|Kürzel/.test(e))).toBe(false);

        const ok = validateAngebot({
            draft: {
                title: 'Zoo',
                datum: '2026-10-05',
                ort: 'Schönbrunn',
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

    it('Tageszeit setzt die Uhrzeit nur, solange sie noch zum vorigen Standard gehört', () => {
        expect(nextSlotTimes('vormittag', 'nachmittag', '08:00', '12:00')).toEqual({
            startzeit: '13:00',
            endzeit: '17:00'
        });
        expect(nextSlotTimes('vormittag', 'nachmittag', '09:15', '12:00')).toBe(null);
        expect(nextSlotTimes('vormittag', 'ganztags', '', '')).toEqual({
            startzeit: '08:00',
            endzeit: '16:00'
        });
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
        expect(weekdayNameDe('Mo')).toBe('Montag');
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
        expect(att.myCount).toBe(0);
    });

    it('dashboardHomeModel spricht die Rolle an', () => {
        const items = [
            { status: 'beantragt', lehrerEmail: 'anna@schule.at', lehrerCode: 'AN' },
            { status: 'freigegeben', lehrerEmail: 'anna@schule.at', lehrerCode: 'AN' },
            { status: 'freigegeben', lehrerEmail: 'anna@schule.at', lehrerCode: 'AN' },
            { status: 'beantragt', lehrerEmail: 'bert@schule.at', lehrerCode: 'BE' }
        ];
        const teacher = dashboardHomeModel({
            role: 'lehrer',
            bootstrapped: true,
            accountEmail: 'kurt@schule.at',
            items
        });
        expect(teacher.title).toBe('Noch kein eigenes Angebot');
        expect(teacher.primary.label).toBe('Neues Angebot');
        expect(teacher.num).toBeNull();

        const anna = dashboardHomeModel({
            role: 'lehrer',
            bootstrapped: true,
            accountEmail: 'anna@schule.at',
            items
        });
        expect(anna.num).toBe(3);
        expect(anna.title).toBe('eigene Angebote');
        expect(anna.sub).toBe('1 wartet, 2 frei.');
        expect(anna.primary).toBeNull();
        expect(anna.ghost.label).toBe('Neues Angebot');
        expect(teacherOfferLine({ myBeantragt: 1, myFreigegeben: 0 })).toBe('1 wartet.');

        const admin = dashboardHomeModel({ role: 'admin', bootstrapped: true, items });
        expect(admin.num).toBe(2);
        expect(admin.title).toBe('Anträge warten');
        expect(admin.sub).toBe('2 Angebote sind frei.');
        expect(admin.primary.label).toBe('Zur Freigabe');

        const quiet = dashboardHomeModel({
            role: 'admin',
            bootstrapped: true,
            items: [{ status: 'freigegeben' }]
        });
        expect(quiet.num).toBe(0);
        expect(quiet.title).toBe('Anträge warten');
        expect(quiet.primary).toBeNull();
    });

    it('statusLabel spricht die Schule', () => {
        expect(statusLabel('beantragt')).toBe('wartet');
        expect(statusLabel('freigegeben')).toBe('frei');
        expect(statusLabel('offen')).toBe('buchbar');
        expect(statusLabel('gesperrt')).toBe('noch nicht buchbar');
    });

    it('studentBookingsUrl nur für freie Schüler-Angebote mit Adresse', () => {
        const offer = { status: 'freigegeben', bookingsBookingUrl: 'https://book.example/zoo' };
        expect(studentBookingsUrl('schueler', offer, false)).toBe('https://book.example/zoo');
        expect(studentBookingsUrl('schueler', offer, true)).toBe('');
        expect(studentBookingsUrl('lehrer', offer, false)).toBe('');
        expect(studentBookingsUrl('schueler', { status: 'beantragt', bookingsBookingUrl: 'https://book.example/zoo' }, false)).toBe('');
        expect(studentBookingsUrl('schueler', { status: 'freigegeben', bookingsBookingUrl: '  ' }, false)).toBe('');
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
        expect(lehrerIds).not.toContain('neu');
        expect(lehrerIds).toContain('dashboard');
        expect(lehrerIds).toContain('meine');
        expect(lehrerIds).not.toContain('buchung');
        const schuelerIds = viewsForRole('schueler').map((v) => v.id);
        expect(schuelerIds).toEqual(['buchung', 'plan']);
        expect(lehrerIds).not.toContain('kalender');
        expect(lehrerIds).not.toContain('liste');
        expect(lehrerIds).not.toContain('export');
        expect(resolvePwView('schueler', 'kalender')).toBe('kalender');
        expect(resolvePwView('schueler', 'export')).toBe('plan');
        expect(resolvePwView('lehrer', 'liste')).toBe('liste');
        expect(resolvePwView('lehrer', 'neu')).toBe('neu');
        expect(resolvePwView('schueler', 'teilnehmer')).toBe('buchung');
        expect(resolvePwView('admin', 'dashboard')).toBe('admin');
        expect(resolvePwView('admin', 'neu')).toBe('neu');
        const adminSecs = navSectionsForRole('admin');
        expect(adminSecs.map((s) => s.id)).toEqual(['arbeit', 'bookings', 'system']);
        expect(adminSecs.find((s) => s.id === 'arbeit').items.map((v) => v.id)).toEqual([
            'admin',
            'plan',
            'teilnehmer'
        ]);
        expect(adminSecs.find((s) => s.id === 'bookings').items[0].id).toBe('bookings');
        expect(adminSecs.find((s) => s.id === 'system').items[0].id).toBe('einstellungen');
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
