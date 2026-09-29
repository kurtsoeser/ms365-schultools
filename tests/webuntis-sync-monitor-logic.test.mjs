import { describe, expect, it } from 'vitest';
import {
    buildSyncMonitorReport,
    parseKursteamState
} from '../src/shared/webuntis-sync-monitor-logic.js';

describe('webuntis-sync-monitor-logic', () => {
    it('parseKursteamState erkennt leeren und gültigen Stand', () => {
        expect(parseKursteamState('').ok).toBe(false);
        expect(parseKursteamState('{').ok).toBe(false);
        const ok = parseKursteamState({ teamsData: [], rawData: [{ lehrer: 'MUEL' }] });
        expect(ok.ok).toBe(true);
    });

    it('buildSyncMonitorReport listet fehlende Besitzer und Lehrer-Matches', () => {
        const report = buildSyncMonitorReport(
            {
                savedAt: '2026-09-01T10:00:00.000Z',
                kursteamEntryMode: 'webuntis',
                yearPrefix: 'SJ26',
                rawData: [{ lehrer: 'MUEL' }, { lehrer: 'NEU' }],
                filteredData: [{ lehrer: 'MUEL' }, { lehrer: 'NEU' }],
                teacherEmailMapping: { MUEL: 'mueller@schule.at' },
                teamsData: [
                    {
                        teamName: 'SJ26 | 1A | M',
                        besitzer: '',
                        gruppenmail: 'm@schule.at',
                        lehrerCode: 'NEU',
                        fach: 'M',
                        originalClass: '1A',
                        isValid: false
                    },
                    {
                        teamName: 'SJ26 | 1B | D',
                        besitzer: 'a@schule.at',
                        gruppenmail: 'd@schule.at',
                        lehrerCode: 'MUEL',
                        fach: 'D',
                        originalClass: '1B',
                        isValid: true
                    }
                ]
            },
            {
                teachers: [
                    { code: 'MUEL', email: 'mueller@schule.at', name: 'Müller' },
                    { code: 'NEU', name: 'Neu' }
                ],
                classes: [
                    { code: '1A', name: '1A' },
                    { code: '1B', name: '1B' },
                    { code: '2A', name: '2A' }
                ]
            }
        );

        expect(report.counts.teams).toBe(2);
        expect(report.counts.missingOwner).toBe(1);
        expect(report.missingOwner[0].teamName).toContain('1A');
        expect(report.teachersWithoutMatch.some((t) => t.code === 'NEU')).toBe(true);
        expect(report.classesWithoutTeam.some((c) => c.code === '2A')).toBe(true);
        expect(report.actionable).toBeGreaterThan(0);
    });
});
