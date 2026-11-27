import { describe, expect, it } from 'vitest';
import {
    eventLooksLikeBooking,
    parentFromEvent,
    readSetup,
    slotsFromEvents
} from '../src/tools/elternsprechtag/elternsprechtag-logic.js';

const setup = readSetup({
    eventDate: '2026-03-18',
    serviceName: 'Termin Elternsprechtag 2026',
    businessId: 'sprechtag@contoso.onmicrosoft.com'
});

describe('elternsprechtag slots', () => {
    it('erkennt Bookings am Organisator und lässt interne Termine weg', () => {
        const booking = {
            id: 'b',
            subject: 'Termin',
            organizer: { emailAddress: { address: 'Sprechtag@contoso.onmicrosoft.com' } },
            start: { dateTime: '2026-03-18T15:20:00' },
            end: { dateTime: '2026-03-18T15:30:00' },
            attendees: [
                { emailAddress: { address: 'lehrer@schule.at', name: 'Lehrkraft' } },
                { emailAddress: { address: 'anna.huber@mail.at', name: 'Anna Huber' } }
            ],
            bodyPreview: 'Kind: Lena Huber\nKlasse: 3A\nTelefon: +43 660 111'
        };
        const intern = {
            id: 'i',
            subject: 'Konferenz',
            organizer: { emailAddress: { address: 'direktion@schule.at' } },
            start: { dateTime: '2026-03-18T14:00:00' },
            end: { dateTime: '2026-03-18T15:00:00' }
        };
        expect(eventLooksLikeBooking(booking, setup)).toBe(true);
        expect(eventLooksLikeBooking(intern, setup)).toBe(false);
        const slots = slotsFromEvents([intern, booking], setup, 'lehrer@schule.at');
        expect(slots.map((s) => s.id)).toEqual(['b']);
        expect(slots[0].parentName).toBe('Anna Huber');
        expect(slots[0].parentEmail).toBe('anna.huber@mail.at');
        expect(slots[0].detail).toBe('Kind: Lena Huber · Klasse: 3A');
        expect(slots[0].phone).toMatch(/660/);
        expect(slots[0].startLabel).toBe('15:20');
    });

    it('liest den Elternnamen aus dem Text, wenn keine Teilnehmerliste da ist', () => {
        const parent = parentFromEvent(
            {
                organizer: { emailAddress: { address: 'bookings@schule.at' } },
                bodyPreview: 'Customer\nName: Markus Leitner\nmarkus.leitner@mail.at'
            },
            'lehrer@schule.at'
        );
        expect(parent.name).toBe('Markus Leitner');
        expect(parent.email).toBe('markus.leitner@mail.at');
    });

    it('zeigt auf Wunsch auch die übrigen Termine des Tages', () => {
        const slots = slotsFromEvents(
            [
                {
                    id: 'i',
                    subject: 'Konferenz',
                    organizer: { emailAddress: { address: 'direktion@schule.at' } },
                    start: { dateTime: '2026-03-18T14:00:00' }
                }
            ],
            setup,
            'lehrer@schule.at',
            { includeOther: true }
        );
        expect(slots).toHaveLength(1);
        expect(slots[0].booking).toBe(false);
    });
});
