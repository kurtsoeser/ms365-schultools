import { describe, it, expect } from 'vitest';
import { buildIcs } from '../src/tools/schularbeiten-planer/schularbeiten-planer-export.js';

describe('schularbeiten-planer-export ICS', () => {
    it('nutzt Titel, TZID und Start/Ende-Uhrzeit', () => {
        const ics = buildIcs(
            [
                {
                    schularbeitId: 'sa-ics-1',
                    datum: '2026-11-11',
                    beginnUhrzeit: '08:35',
                    dauerMinuten: 105,
                    fachCode: 'ENWS',
                    klasseCode: '4AK',
                    titel: 'ENWS 4ABCDEK',
                    status: 'fixiert'
                }
            ],
            { subjects: [], classes: [], teachers: [] }
        );
        expect(ics).toContain('BEGIN:VTIMEZONE');
        expect(ics).toContain('TZID:Europe/Vienna');
        expect(ics).toContain('X-WR-TIMEZONE:Europe/Vienna');
        expect(ics).toContain('SUMMARY:ENWS 4ABCDEK');
        expect(ics).toContain('DTSTART;TZID=Europe/Vienna:20261111T083500');
        expect(ics).toContain('DTEND;TZID=Europe/Vienna:20261111T102000');
        expect(ics).toContain('Zeit: 08:35 – 10:20 Uhr');
    });

    it('Wanduhrzeit bleibt über DST-Grenzen gleich (Winter vs. Sommer)', () => {
        const winter = buildIcs(
            [
                {
                    schularbeitId: 'sa-w',
                    datum: '2026-01-15',
                    beginnUhrzeit: '08:00',
                    dauerMinuten: 50,
                    titel: 'Winter',
                    status: 'fixiert'
                }
            ],
            { subjects: [], classes: [], teachers: [] }
        );
        const summer = buildIcs(
            [
                {
                    schularbeitId: 'sa-s',
                    datum: '2026-06-15',
                    beginnUhrzeit: '08:00',
                    dauerMinuten: 50,
                    titel: 'Sommer',
                    status: 'fixiert'
                }
            ],
            { subjects: [], classes: [], teachers: [] }
        );
        expect(winter).toContain('DTSTART;TZID=Europe/Vienna:20260115T080000');
        expect(winter).toContain('DTEND;TZID=Europe/Vienna:20260115T085000');
        expect(summer).toContain('DTSTART;TZID=Europe/Vienna:20260615T080000');
        expect(summer).toContain('DTEND;TZID=Europe/Vienna:20260615T085000');
    });
});
