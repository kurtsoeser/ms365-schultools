import { describe, expect, it } from 'vitest';
import {
    emptyPeriodLabel,
    listForDay,
    layoutTimedEventsForDay,
    startOfWeekMonday,
    timedEventPlacement,
    weekDays,
    monthGridDays
} from '../src/tools/termine-hub/termine-hub-logic.js';
import {
    buildGraphEventBody,
    calendarBasePath,
    mapGraphEventToTermin,
    toGraphDateTimeParam
} from '../src/tools/termine-hub/termine-hub-outlook.js';
import {
    audienceKeyFromPersonas,
    filterEventsByVisibleSources,
    normalizeSettings,
    resolveEnabledSources,
    SOURCE_IDS,
    sourceAllowedForAudience,
    sourceWritableForAudience
} from '../src/tools/termine-hub/termine-hub-settings.js';

describe('Kalender und Termine Outlook helpers', () => {
    it('startOfWeekMonday beginnt am Montag', () => {
        // 2026-11-11 = Mittwoch
        const mon = startOfWeekMonday(new Date(2026, 10, 11));
        expect(mon.getDay()).toBe(1);
        expect(mon.getDate()).toBe(9);
    });

    it('weekDays liefert 7 Tage', () => {
        const days = weekDays(new Date(2026, 10, 11));
        expect(days).toHaveLength(7);
        expect(days[0].getDay()).toBe(1);
        expect(days[6].getDay()).toBe(0);
    });

    it('emptyPeriodLabel erklärt leeres Raster', () => {
        expect(emptyPeriodLabel('week')).toBe('Keine Termine in dieser Woche.');
        expect(emptyPeriodLabel('month')).toBe('Keine Termine in diesem Monat.');
        expect(emptyPeriodLabel('agenda')).toBe('Keine Termine in diesem Zeitraum.');
    });

    it('monthGridDays liefert 42 Zellen', () => {
        expect(monthGridDays(2026, 10)).toHaveLength(42);
    });

    it('listForDay filtert korrekt', () => {
        const termine = [
            { id: '1', title: 'A', start: '2026-11-10T10:00:00.000Z', end: '2026-11-10T11:00:00.000Z' },
            { id: '2', title: 'B', start: '2026-11-11T10:00:00.000Z', end: '2026-11-11T11:00:00.000Z' }
        ];
        expect(listForDay(termine, '2026-11-10').map((t) => t.id)).toEqual(['1']);
    });

    it('calendarBasePath für me und user', () => {
        expect(calendarBasePath({ kind: 'me' })).toBe('/me/calendar');
        expect(calendarBasePath({ kind: 'user', upn: 'schulkalender@schule.at' })).toBe(
            '/users/schulkalender%40schule.at/calendar'
        );
        expect(
            calendarBasePath({
                kind: 'user',
                upn: 'ich@schule.at',
                signedInEmail: 'ich@schule.at'
            })
        ).toBe('/me/calendar');
    });

    it('mapGraphEventToTermin mappt all-day Ende exklusiv', () => {
        const t = mapGraphEventToTermin(
            {
                id: 'ev1',
                subject: 'Schulfest',
                isAllDay: true,
                start: { dateTime: '2026-06-12' },
                end: { dateTime: '2026-06-13' },
                location: { displayName: 'Hof' }
            },
            { id: 'schul', label: 'Schulkalender' }
        );
        expect(t.title).toBe('Schulfest');
        expect(t.allDay).toBe(true);
        expect(t.sourceId).toBe('schul');
        expect(t.ort).toBe('Hof');
        expect(String(t.start)).toContain('2026-06-12');
    });

    it('buildGraphEventBody setzt all-day Ende exklusiv', () => {
        const body = buildGraphEventBody({
            title: 'Konferenz',
            allDay: true,
            start: '2026-11-10',
            end: '2026-11-10',
            ort: 'Aula'
        });
        expect(body.isAllDay).toBe(true);
        expect(body.start.dateTime).toBe('2026-11-10');
        expect(body.end.dateTime).toBe('2026-11-11');
        expect(body.location.displayName).toBe('Aula');
    });

    it('toGraphDateTimeParam wirft bei ungültig', () => {
        expect(() => toGraphDateTimeParam('kein-datum')).toThrow(/Ungültig/);
    });

    it('resolveEnabledSources respektiert UPNs', () => {
        const src = resolveEnabledSources(
            normalizeSettings({
                schulkalenderUpn: 'schul@x.at',
                lehrerkalenderUpn: '',
                schuelerkalenderUpn: '',
                enabledSources: [
                    SOURCE_IDS.ME,
                    SOURCE_IDS.SCHUL,
                    SOURCE_IDS.LEHRER,
                    SOURCE_IDS.SCHUELER
                ]
            })
        );
        expect(src.map((s) => s.id)).toEqual([SOURCE_IDS.ME, SOURCE_IDS.SCHUL]);
    });

    it('resolveEnabledSources filtert nach Rolle', () => {
        const settings = normalizeSettings({
            schulkalenderUpn: 'schul@x.at',
            lehrerkalenderUpn: 'lehrer@x.at',
            schuelerkalenderUpn: 'schueler@x.at',
            enabledSources: [
                SOURCE_IDS.ME,
                SOURCE_IDS.SCHUL,
                SOURCE_IDS.LEHRER,
                SOURCE_IDS.SCHUELER
            ]
        });
        expect(
            resolveEnabledSources(settings, { audience: 'lehrer' }).map((s) => s.id)
        ).toEqual([SOURCE_IDS.ME, SOURCE_IDS.SCHUL, SOURCE_IDS.LEHRER]);
        expect(
            resolveEnabledSources(settings, { audience: 'schueler' }).map((s) => s.id)
        ).toEqual([SOURCE_IDS.ME, SOURCE_IDS.SCHUL, SOURCE_IDS.SCHUELER]);
        expect(
            resolveEnabledSources(settings, { audience: 'it' }).map((s) => s.id)
        ).toEqual([
            SOURCE_IDS.ME,
            SOURCE_IDS.SCHUL,
            SOURCE_IDS.LEHRER,
            SOURCE_IDS.SCHUELER
        ]);
    });

    it('audience helpers map Personas', () => {
        expect(audienceKeyFromPersonas({ isIt: true })).toBe('it');
        expect(audienceKeyFromPersonas({ isLehrer: true })).toBe('lehrer');
        expect(audienceKeyFromPersonas({ isSchueler: true })).toBe('schueler');
        expect(audienceKeyFromPersonas({ loggedIn: false })).toBe('');
        expect(audienceKeyFromPersonas({})).toBe('');
        expect(audienceKeyFromPersonas({ loggedIn: true })).toBe('self');
        expect(
            audienceKeyFromPersonas({
                loggedIn: true,
                isSchueler: true,
                unknownUser: true
            })
        ).toBe('self');
        expect(sourceAllowedForAudience(SOURCE_IDS.LEHRER, 'schueler')).toBe(false);
        expect(sourceAllowedForAudience(SOURCE_IDS.SCHUELER, 'lehrer')).toBe(false);
        expect(sourceAllowedForAudience(SOURCE_IDS.SCHUL, 'schueler')).toBe(true);
        expect(sourceAllowedForAudience(SOURCE_IDS.SCHUL, '')).toBe(false);
        expect(sourceAllowedForAudience(SOURCE_IDS.ME, 'self')).toBe(true);
        expect(sourceAllowedForAudience(SOURCE_IDS.SCHUL, 'self')).toBe(false);
        expect(sourceWritableForAudience(SOURCE_IDS.SCHUL, 'schueler')).toBe(false);
        expect(sourceWritableForAudience(SOURCE_IDS.ME, 'schueler')).toBe(true);
        expect(sourceWritableForAudience(SOURCE_IDS.ME, 'self')).toBe(true);
        expect(sourceWritableForAudience(SOURCE_IDS.SCHUL, 'self')).toBe(false);
        expect(sourceWritableForAudience(SOURCE_IDS.SCHUL, 'lehrer')).toBe(true);
        expect(resolveEnabledSources(
            normalizeSettings({
                schulkalenderUpn: 'schul@x.at',
                enabledSources: [SOURCE_IDS.ME, SOURCE_IDS.SCHUL]
            }),
            { audience: '' }
        )).toEqual([]);
    });

    it('self-Audience: eigener Kalender auch ohne Portal-Rolle und ohne enabled me', () => {
        const settings = normalizeSettings({
            schulkalenderUpn: 'schul@x.at',
            lehrerkalenderUpn: 'lehrer@x.at',
            schuelerkalenderUpn: 'schueler@x.at',
            enabledSources: [SOURCE_IDS.SCHUL, SOURCE_IDS.LEHRER, SOURCE_IDS.SCHUELER]
        });
        const src = resolveEnabledSources(settings, { audience: 'self' });
        expect(src.map((s) => s.id)).toEqual([SOURCE_IDS.ME]);
        expect(sourceWritableForAudience(SOURCE_IDS.ME, 'self')).toBe(true);
        expect(audienceKeyFromPersonas({ loggedIn: true, isIt: false, isLehrer: false, isSchueler: false })).toBe(
            'self'
        );
    });

    it('erzwingt Mein Kalender für jede Audience auch ohne enabled me', () => {
        const settings = normalizeSettings({
            enabledSources: [SOURCE_IDS.SCHUL],
            schulkalenderUpn: 'schul@x.at'
        });
        expect(resolveEnabledSources(settings, { audience: 'lehrer' }).map((s) => s.id)).toEqual([
            SOURCE_IDS.ME,
            SOURCE_IDS.SCHUL
        ]);
        expect(resolveEnabledSources(settings, { audience: 'schueler' }).map((s) => s.id)).toEqual([
            SOURCE_IDS.ME,
            SOURCE_IDS.SCHUL
        ]);
    });

    it('filterEventsByVisibleSources blendet Layer aus', () => {
        const settings = normalizeSettings({
            enabledSources: [SOURCE_IDS.ME, SOURCE_IDS.SCHUL, SOURCE_IDS.LEHRER],
            visibleSources: [SOURCE_IDS.SCHUL],
            schulkalenderUpn: 'schul@x.at',
            lehrerkalenderUpn: 'lehrer@x.at'
        });
        const filtered = filterEventsByVisibleSources(
            [
                { id: '1', sourceId: 'me', title: 'Privat' },
                { id: '2', sourceId: 'schul', title: 'Schule' },
                { id: '3', sourceId: 'lehrer', title: 'Kollegium' }
            ],
            settings,
            { audience: 'lehrer' }
        );
        expect(filtered.map((e) => e.id)).toEqual(['2']);
        expect(
            resolveEnabledSources(settings, { audience: 'lehrer' }).find((s) => s.id === 'schul')
                ?.visible
        ).toBe(true);
        expect(
            resolveEnabledSources(settings, { audience: 'lehrer' }).find((s) => s.id === 'me')
                ?.visible
        ).toBe(false);
    });

    it('timedEventPlacement setzt Top/Höhe im Raster', () => {
        const place = timedEventPlacement(
            {
                title: 'A',
                start: '2026-11-10T09:00:00',
                end: '2026-11-10T10:30:00',
                allDay: false
            },
            '2026-11-10',
            { startHour: 7, endHour: 20 }
        );
        expect(place).toBeTruthy();
        // 09:00 = 2h nach 07:00 → 2/13
        expect(place.topPct).toBeCloseTo((2 / 13) * 100, 5);
        expect(place.heightPct).toBeCloseTo((1.5 / 13) * 100, 5);
    });

    it('layoutTimedEventsForDay vergibt Spalten bei Überlappung', () => {
        const laid = layoutTimedEventsForDay(
            [
                {
                    id: '1',
                    title: 'A',
                    start: '2026-11-10T09:00:00',
                    end: '2026-11-10T10:00:00'
                },
                {
                    id: '2',
                    title: 'B',
                    start: '2026-11-10T09:30:00',
                    end: '2026-11-10T10:30:00'
                }
            ],
            '2026-11-10',
            { startHour: 7, endHour: 20 }
        );
        expect(laid).toHaveLength(2);
        expect(laid[0].cols).toBe(2);
        expect(new Set(laid.map((x) => x.col)).size).toBe(2);
    });
});
