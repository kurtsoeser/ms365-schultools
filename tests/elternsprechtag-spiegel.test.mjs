import { describe, expect, it } from 'vitest';
import {
    dayWasPublished,
    fieldsForRow,
    itemIdsForDate,
    resolveSpiegelSiteUrl,
    rowsFromAppointments,
    slotsFromSpiegelItems
} from '../src/tools/elternsprechtag/elternsprechtag-spiegel-logic.js';

const staff = [
    { id: 'st-anna', emailAddress: 'Anna.Lehrer@schule.at', displayName: 'Anna Lehrer' },
    { id: 'st-max', emailAddress: 'max.lehrer@schule.at', displayName: 'Max Lehrer' }
];

const appointment = {
    id: 'ap-1',
    serviceName: 'Termin Elternsprechtag 2026',
    staffMemberIds: ['st-anna'],
    start: { dateTime: '2026-03-18T15:00:00', timeZone: 'Europe/Vienna' },
    end: { dateTime: '2026-03-18T15:10:00', timeZone: 'Europe/Vienna' },
    joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/abc',
    serviceLocation: { displayName: 'Zimmer 12' },
    customers: [
        {
            name: 'Anna Huber',
            emailAddress: 'anna.huber@mail.at',
            phone: '+43 660 1234567',
            customQuestionAnswers: [
                { question: 'Name des Kindes', answer: 'Lena Huber' },
                { question: 'Klasse', answer: '3A' },
                { question: 'Anliegen', answer: 'Versetzung' }
            ]
        }
    ]
};

describe('elternsprechtag spiegel', () => {
    it('macht aus einer Buchung eine Zeile der zugeordneten Lehrkraft', () => {
        const built = rowsFromAppointments([appointment], staff, {
            eventDate: '2026-03-18',
            serviceName: 'Termin Elternsprechtag 2026'
        });
        expect(built.skippedNoTeacher).toBe(0);
        expect(built.rows).toHaveLength(1);
        const row = built.rows[0];
        expect(row.lehrkraft).toBe('anna.lehrer@schule.at');
        expect(row.eltern).toBe('Anna Huber');
        expect(row.kind).toBe('Lena Huber');
        expect(row.klasse).toBe('3A');
        expect(row.hinweis).toContain('Versetzung');
        expect(row.start).toBe('15:00');
        expect(row.link).toContain('teams.microsoft.com');
        const fields = fieldsForRow(row);
        expect(fields.Lehrkraft).toBe('anna.lehrer@schule.at');
        expect(fields.Datum).toBe('2026-03-18');
        expect(fields.Kind).toBe('Lena Huber');
    });

    it('lässt andere Dienste und Buchungen ohne Lehrkraft weg', () => {
        const other = { ...appointment, id: 'ap-2', serviceName: 'Aufnahme', staffMemberIds: ['st-anna'] };
        const orphan = { ...appointment, id: 'ap-3', staffMemberIds: ['unbekannt'] };
        const built = rowsFromAppointments([other, orphan], staff, {
            eventDate: '2026-03-18',
            serviceName: 'Termin Elternsprechtag 2026'
        });
        expect(built.rows).toHaveLength(0);
        expect(built.skippedNoTeacher).toBe(1);
    });

    it('zeigt der Lehrkraft nur die eigenen Zeilen des Tages', () => {
        const items = [
            {
                id: '1',
                fields: fieldsForRow({
                    terminId: 'a',
                    datum: '2026-03-18',
                    start: '15:00',
                    end: '15:10',
                    lehrkraft: 'anna.lehrer@schule.at',
                    eltern: 'Anna Huber',
                    email: 'anna.huber@mail.at',
                    kind: 'Lena Huber',
                    klasse: '3A'
                })
            },
            {
                id: '2',
                fields: fieldsForRow({
                    terminId: 'b',
                    datum: '2026-03-18',
                    start: '15:20',
                    lehrkraft: 'max.lehrer@schule.at',
                    eltern: 'Paul Leitner'
                })
            },
            {
                id: '3',
                fields: fieldsForRow({
                    terminId: 'c',
                    datum: '2026-03-19',
                    start: '15:00',
                    lehrkraft: 'anna.lehrer@schule.at',
                    eltern: 'Anderer Tag'
                })
            }
        ];
        const mine = slotsFromSpiegelItems(items, 'Anna.Lehrer@schule.at', '2026-03-18');
        expect(mine.map((s) => s.parentName)).toEqual(['Anna Huber']);
        expect(mine[0].detail).toContain('Lena Huber');
        expect(mine[0].detail).toContain('3A');
        expect(dayWasPublished(items, '2026-03-18')).toBe(true);
        expect(dayWasPublished(items, '2026-04-01')).toBe(false);
        expect(itemIdsForDate(items, '2026-03-18')).toEqual(['1', '2']);
    });

    it('nimmt die hinterlegte Website, sonst die Schul-Website', () => {
        expect(resolveSpiegelSiteUrl({ siteUrl: 'https://schule.sharepoint.com/sites/tools' }, '')).toBe(
            'https://schule.sharepoint.com/sites/tools'
        );
        expect(resolveSpiegelSiteUrl({}, 'https://schule.sharepoint.com/sites/planer')).toBe(
            'https://schule.sharepoint.com/sites/planer'
        );
        expect(resolveSpiegelSiteUrl({}, 'kein-link')).toBe('');
    });
});
