import { describe, expect, it } from 'vitest';
import {
    applyParentClass,
    childrenByParentEmail,
    fillFreeSlots,
    focusBooking,
    pickConferenceDate,
    presentDay
} from '../src/tools/elternsprechtag/elternsprechtag-day.js';
import {
    serviceWindow,
    windowFromItems,
    withServiceWindow
} from '../src/tools/elternsprechtag/elternsprechtag-spiegel-logic.js';

const booked = {
    id: 'b',
    booking: true,
    startLabel: '15:00',
    endLabel: '15:10',
    parentName: 'Anna Huber',
    parentEmail: 'anna.huber@mail.at',
    childName: '',
    className: ''
};

describe('elternsprechtag tagesleiste', () => {
    it('füllt freie Slots zwischen den Dienstzeiten', () => {
        const day = fillFreeSlots([booked], { start: '15:00', end: '15:30', minutes: 10 });
        expect(day.map((slot) => slot.startLabel + ' ' + (slot.free ? 'frei' : slot.parentName))).toEqual([
            '15:00 Anna Huber',
            '15:10 frei',
            '15:20 frei'
        ]);
        expect(day[1].endLabel).toBe('15:20');
    });

    it('setzt die Klasse aus der Eltern-Mail der Stammdaten', () => {
        const index = childrenByParentEmail(
            [
                { name: 'Lena Huber', klasse: '3A', guardianIds: ['g1'] },
                { name: 'Paul Huber', klasse: '1B', guardianIds: ['g1'] }
            ],
            [{ id: 'g1', email: 'Anna.Huber@mail.at' }]
        );
        const slot = applyParentClass(booked, index);
        expect(slot.className).toBe('3A, 1B');
        expect(slot.childName).toBe('');
        const only = applyParentClass(booked, childrenByParentEmail(
            [{ name: 'Lena Huber', klasse: '3A', parentPairs: [{ email: 'anna.huber@mail.at' }] }],
            []
        ));
        expect(only.childName).toBe('Lena Huber');
        expect(only.className).toBe('3A');
        const shown = presentDay([booked], { start: '15:00', end: '15:20', minutes: 10 }, {
            students: [{ name: 'Lena Huber', klasse: '3A', guardianIds: ['g1'] }],
            guardians: [{ id: 'g1', email: 'anna.huber@mail.at' }]
        });
        expect(shown[0].className).toBe('3A');
        expect(shown[1].free).toBe(true);
        const folded = presentDay([booked], { start: '15:00', end: '15:40', minutes: 10 });
        expect(folded.filter((slot) => slot.free)).toHaveLength(1);
        expect(folded[1].startLabel).toBe('15:10');
        expect(folded[1].endLabel).toBe('15:40');
    });

    it('zeigt den nächsten Termin und wählt den Sprechtag statt heute', () => {
        const later = { ...booked, id: 'c', startLabel: '15:20', endLabel: '15:30', parentName: 'Sara Novak' };
        const day = fillFreeSlots([booked, later], { start: '15:00', end: '15:40', minutes: 10 });
        expect(focusBooking(day, { isToday: true, nowMinutes: 15 * 60 + 5 }).kind).toBe('jetzt');
        expect(focusBooking(day, { isToday: true, nowMinutes: 15 * 60 + 12 }).slot.parentName).toBe('Sara Novak');
        expect(focusBooking(day, { isToday: false }).kind).toBe('erster');
        expect(pickConferenceDate(['2026-03-01', '2026-03-18', '2026-04-02'], '2026-03-10')).toBe('2026-03-18');
        expect(pickConferenceDate(['2026-03-18'], '2026-03-18')).toBe('2026-03-18');
    });

    it('merkt sich das Dienstfenster auch ohne Buchung', () => {
        expect(serviceWindow({ start: '15:00', end: '19:00', minutes: '10' })).toEqual({
            start: '15:00',
            end: '19:00',
            minutes: 10
        });
        const rows = withServiceWindow([], { start: '15:00', end: '19:00', minutes: 10 }, '2026-03-18');
        expect(rows).toHaveLength(1);
        expect(rows[0].lehrkraft).toBe('');
        expect(rows[0].fensterStart).toBe('15:00');
        const found = windowFromItems(
            [{ id: '1', fields: { Datum: '2026-03-18', FensterStart: '15:00', FensterEnde: '19:00', SlotMinuten: '10' } }],
            '2026-03-18'
        );
        expect(found.minutes).toBe(10);
    });
});
