import { describe, it, expect } from 'vitest';
import { buildGroupCalendarEvent } from '../src/tools/schularbeiten-planer/schularbeiten-planer-calendar-sync.js';

describe('schularbeiten-planer-calendar-sync', () => {
    it('Graph-Termin ohne Erinnerung, ohne Online-Meeting, ohne Antwortanfrage', () => {
        const ev = buildGroupCalendarEvent(
            {
                schularbeitId: 'sa-test-1',
                datum: '2026-11-11',
                fachCode: 'MAM',
                klasseCode: '3AK',
                thema: 'Test'
            },
            { fach: 'Mathematik', klasse: '3AK' }
        );
        expect(ev.isReminderOn).toBe(false);
        expect(ev.reminderMinutesBeforeStart).toBe(0);
        expect(ev.responseRequested).toBe(false);
        expect(ev.isOnlineMeeting).toBe(false);
        expect(ev.attendees).toBeUndefined();
        expect(ev.isAllDay).toBe(true);
    });

    it('nutzt Beginn-Uhrzeit und Dauer für Terminende', () => {
        const ev = buildGroupCalendarEvent(
            {
                schularbeitId: 'sa-test-2',
                datum: '2026-11-11',
                beginnUhrzeit: '08:00',
                dauerMinuten: 100,
                fachCode: 'D',
                klasseCode: '3AK',
                titel: '1. SA'
            },
            { fach: 'Deutsch', klasse: '3AK' }
        );
        expect(ev.isAllDay).toBe(false);
        expect(ev.start.dateTime).toContain('2026-11-11T08:00:00');
        expect(ev.end.dateTime).toBe('2026-11-11T09:40:00.0000000');
        expect(ev.start.timeZone).toBe('Europe/Vienna');
        expect(ev.subject).toContain('1. SA');
        expect(ev.subject).toContain('Deutsch');
        expect(ev.body.content).toContain('08:00 – 09:40 Uhr');
    });
});
