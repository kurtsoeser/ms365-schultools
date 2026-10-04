import { describe, it, expect } from 'vitest';
import {
    teachingSlotKey,
    plannedRowsFromWebuntisLines,
    buildKursteamAbgleichReport,
    collectPlannedRowsFromKursteamState
} from '../src/shared/kursteam-belegung-abgleich-logic.js';

describe('kursteam-belegung-abgleich-logic', () => {
    it('erzeugt Unterrichts-Slots aus WebUntis-Zeilen', () => {
        const rows = plannedRowsFromWebuntisLines([
            { klasse: '1A', fach: 'BW', lehrer: 'FRECH' },
            { klasse: '1A', fach: 'BW', lehrer: 'FRECH' },
            { klasse: '1A', fach: 'D', lehrer: 'DORFE' }
        ]);
        expect(rows).toHaveLength(2);
        expect(rows.map((r) => teachingSlotKey(r)).sort()).toEqual(['1A|DORFE|D|', '1A|FRECH|BW|']);
    });

    it('matched per gruppenmail', () => {
        const report = buildKursteamAbgleichReport(
            [{ klasse: '1A', fach: 'BW', lehrerCode: 'FRECH', gruppenmail: 'sj26-27-jg2031-hakb-bw-frech' }],
            [
                {
                    klasse: '1A',
                    fach: 'BW',
                    lehrerCode: 'FRECH',
                    gruppenmail: 'sj26-27-jg2031-hakb-bw-frech',
                    graphGroupId: 'g1',
                    teamName: 'SJ26-27 | 1A | BW-FRECH'
                }
            ]
        );
        expect(report.counts.matched).toBe(1);
        expect(report.counts.missingInM365).toBe(0);
        expect(report.counts.onlyInM365).toBe(0);
    });

    it('erkennt fehlende Teams in M365', () => {
        const report = buildKursteamAbgleichReport(
            [{ klasse: '1A', fach: 'GEO', lehrerCode: 'DORNE', gruppenmail: '' }],
            []
        );
        expect(report.counts.missingInM365).toBe(1);
    });

    it('matched per Unterrichts-Slot ohne Mail', () => {
        const report = buildKursteamAbgleichReport(
            [{ klasse: '1AK', fach: 'BESPK', lehrerCode: 'RINNE', gruppenmail: '' }],
            [
                {
                    klasse: '1AK',
                    fach: 'BESPK',
                    lehrerCode: 'RINNE',
                    gruppenmail: 'sj26-27-jg2031-ak-bespk-rinne',
                    graphGroupId: 'g2'
                }
            ]
        );
        expect(report.counts.matched).toBe(1);
        expect(report.matched[0].matchBy).toBe('unterricht');
    });

    it('nutzt teamsData aus Wizard-State', () => {
        const { source, rows } = collectPlannedRowsFromKursteamState({
            teamsGenerated: true,
            teamsData: [
                {
                    isValid: true,
                    originalClass: '1A',
                    fach: 'BW',
                    lehrerCode: 'FRECH',
                    gruppenmail: 'sj26-27-jg2031-hakb-bw-frech',
                    teamName: 'SJ26-27 | 1A | BW-FRECH',
                    besitzer: 'frech@school.at'
                }
            ]
        });
        expect(source).toBe('teamsData');
        expect(rows).toHaveLength(1);
        expect(rows[0].gruppenmail).toBe('sj26-27-jg2031-hakb-bw-frech');
    });
});
