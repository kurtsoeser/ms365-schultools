import { describe, it, expect } from 'vitest';
import {
    buildLfrCalendarEvent,
    calendarEventsPath,
    formatLfrCalendarAccessError
} from '../src/tools/lehrer-freistellung-planer/lfr-calendar-sync.js';

describe('lfr-calendar-sync', () => {
    it('baut Graph-Event für genehmigten Antrag', () => {
        const ev = buildLfrCalendarEvent({
            titel: 'Fortbildung',
            lehrerName: 'Max Mustermann',
            lehrerEmail: 'max@schule.at',
            beginn: '2026-03-10T08:00:00',
            ende: '2026-03-10T16:00:00',
            kategorie: 'Fortbildung',
            antragId: 'lfr-1'
        });
        expect(ev.subject).toContain('Max Mustermann');
        expect(ev.start.dateTime).toContain('2026-03-10');
        expect(ev.categories).toContain('Lehrer-Freistellung');
    });

    it('nutzt /me wenn Ziel = angemeldetes Konto oder useMe', () => {
        expect(
            calendarEventsPath({
                calendarUser: 'kurt@schule.at',
                signedInEmail: 'kurt@schule.at'
            })
        ).toBe('/me/calendar/events');
        expect(calendarEventsPath({ useMe: true })).toBe('/me/calendar/events');
    });

    it('schreibt fremdes Postfach nur über /users/…', () => {
        expect(
            calendarEventsPath({
                calendarUser: 'demo-administrator@ms365.schule',
                signedInEmail: 'kurt@schule.at'
            })
        ).toBe('/users/demo-administrator%40ms365.schule/calendar/events');
    });

    it('erklärt AccessDenied verständlich', () => {
        const msg = formatLfrCalendarAccessError(
            { message: 'ErrorAccessDenied: Access is denied.' },
            'demo-administrator@ms365.schule'
        );
        expect(msg).toMatch(/freigegeben/i);
        expect(msg).toMatch(/Global Admin/i);
    });
});
