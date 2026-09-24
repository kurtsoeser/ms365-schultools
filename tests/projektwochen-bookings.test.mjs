import { describe, expect, it } from 'vitest';
import {
    durationMinutesFromTimes,
    durationIsoFromMinutesPw,
    buildServicePayloadFromAngebot,
    normalizeAppointments,
    enrichAttendeesWithKlasse,
    attendeesToCsv,
    bookingOpenDateFromBuchungAb,
    defaultBusinessName
} from '../src/tools/projektwochen/projektwochen-bookings-logic.js';

describe('projektwochen-bookings-logic', () => {
    it('berechnet Dauer und ISO', () => {
        expect(durationMinutesFromTimes('08:00', '12:00')).toBe(240);
        expect(durationIsoFromMinutesPw(240)).toBe('PT4H');
        expect(durationIsoFromMinutesPw(90)).toBe('PT90M');
    });

    it('buildServicePayloadFromAngebot setzt Kapazität und Marker', () => {
        const r = buildServicePayloadFromAngebot(
            {
                title: 'Zoo',
                beschreibung: 'Ausflug',
                ort: 'Schönbrunn',
                datum: '2026-10-05',
                startzeit: '08:00',
                endzeit: '16:00',
                kapazitaet: 25,
                preisEuro: 12,
                angebotId: 'ang-abc',
                buchungAb: '2026-09-20T08:00'
            },
            { buchungAbDefault: '2026-09-01T08:00' }
        );
        expect(r.ok).toBe(true);
        expect(r.payload.maximumAttendeesCount).toBe(25);
        expect(r.payload.defaultPriceType).toBe('fixedPrice');
        expect(r.payload.description).toContain('[PW:ang-abc]');
        expect(r.maxAdvance).toMatch(/^P\d+D$/);
    });

    it('bookingOpenDateFromBuchungAb', () => {
        expect(bookingOpenDateFromBuchungAb('2026-09-20T08:00:00', '2026-10-05')).toBe('2026-09-20');
    });

    it('normalizeAppointments + CSV', () => {
        const angebote = [
            {
                angebotId: 'a1',
                title: 'Zoo',
                bookingsServiceId: 'svc1',
                kapazitaet: 20,
                datum: '2026-10-05'
            }
        ];
        const { rows, occupancy } = normalizeAppointments(
            [
                {
                    id: 'ap1',
                    serviceId: 'svc1',
                    serviceName: 'Zoo',
                    maximumAttendeesCount: 20,
                    filledAttendeesCount: 2,
                    customers: [
                        { name: 'Anna', emailAddress: 'anna@schule.at' },
                        { name: 'Ben', emailAddress: 'ben@schule.at' }
                    ]
                }
            ],
            angebote
        );
        expect(rows).toHaveLength(2);
        expect(occupancy.svc1.filled).toBe(2);
        const enriched = enrichAttendeesWithKlasse(rows, [
            { email: 'anna@schule.at', klasse: '1AK' }
        ]);
        expect(enriched[0].klasse).toBe('1AK');
        const csv = attendeesToCsv(enriched);
        expect(csv).toContain('Anna');
        expect(csv).toContain('1AK');
    });

    it('defaultBusinessName', () => {
        expect(defaultBusinessName({ title: 'Projektwoche 27' })).toBe('Projektwoche 27');
        expect(defaultBusinessName({ aktionId: 'pw-1' })).toContain('pw-1');
    });
});
