import { describe, expect, it } from 'vitest';
import {
    angeboteToCsv,
    buildAngeboteIcs,
    buildWeekPlanPrintHtml,
    buildAttendeesPrintHtml
} from '../src/tools/projektwochen/projektwochen-export.js';

describe('projektwochen-export', () => {
    const sample = [
        {
            title: 'Zoo',
            status: 'freigegeben',
            datum: '2026-10-05',
            tag: 'Mo',
            slot: 'ganztags',
            startzeit: '08:00',
            endzeit: '16:00',
            ort: 'Schönbrunn',
            kapazitaet: 25,
            preisEuro: 12,
            lehrerCode: 'MU',
            zielklassen: 'alle',
            kategorie: 'exkursion',
            angebotId: 'ang-1'
        }
    ];

    it('angeboteToCsv enthält Titel und Status', () => {
        const csv = angeboteToCsv(sample);
        expect(csv).toContain('Zoo');
        expect(csv).toContain('freigegeben');
        expect(csv.split('\r\n').length).toBe(2);
    });

    it('buildAngeboteIcs erzeugt VEVENT', () => {
        const ics = buildAngeboteIcs(sample, { title: 'PW Demo' });
        expect(ics).toContain('BEGIN:VEVENT');
        expect(ics).toContain('SUMMARY:Zoo');
        expect(ics).toContain('DTSTART:20261005T080000');
    });

    it('buildWeekPlanPrintHtml enthält Wochentage und spannt ganztags', () => {
        const html = buildWeekPlanPrintHtml(sample, { title: 'PW', startdatum: '2026-10-05', enddatum: '2026-10-09' });
        expect(html).toContain('Zoo');
        expect(html).toContain('<th>Mo</th>');
        expect(html).toContain('window.print');
        expect(html).toContain('colspan="3"');
        expect(html).toContain('<th>vormittag</th>');
        expect(html).not.toMatch(/<th>ganztags<\/th>/);
    });

    it('buildAttendeesPrintHtml listet TN', () => {
        const html = buildAttendeesPrintHtml(
            [{ name: 'Anna', email: 'a@x.at', klasse: '1AK', angebotTitle: 'Zoo', datum: '2026-10-05' }],
            { title: 'PW' }
        );
        expect(html).toContain('Anna');
        expect(html).toContain('1AK');
    });
});
