import { describe, it, expect } from 'vitest';
import {
    buildBackupCoverageReport,
    compareSchooltoolSummaries,
    summarizeSchooltoolFromBackup
} from '../src/shared/stammdaten-sharepoint-sync-logic.js';

function v2Payload(overrides) {
    const core = Object.assign(
        {
            schoolName: 'Demo',
            domain: 'demo.at',
            teachers: [{ name: 'A' }],
            subjects: [{ code: 'D' }],
            verwaltungAudienceGroups: [{ id: 'schulleitung' }],
            adminAudienceMemberships: []
        },
        overrides && overrides.core ? overrides.core : {}
    );
    return {
        kind: 'ms365-browser-backup-v1',
        exportedAt: '2026-01-01T12:00:00.000Z',
        localStorage: {
            'ms365-schooltool-data-v2': {
                version: 4,
                core: core,
                years: {
                    current: '2025/26',
                    byLabel: {
                        '2025/26': {
                            students: [{ name: 'S' }],
                            classes: [{ name: '1A' }]
                        }
                    }
                },
                setup: { matched: { schuelerGroupId: 'g1' } }
            },
            'ms365-dashboard-tool-access-v1': { tools: {} }
        }
    };
}

describe('stammdaten-backup-coverage', () => {
    it('fasst Stammdaten v2 zusammen', () => {
        const s = summarizeSchooltoolFromBackup(v2Payload());
        expect(s.present).toBe(true);
        expect(s.schoolName).toBe('Demo');
        expect(s.teacherCount).toBe(1);
        expect(s.studentCount).toBe(1);
        expect(s.verwaltungAudienceGroupCount).toBe(1);
        expect(s.schuelerGroupLinked).toBe(true);
    });

    it('erkennt Abweichungen in der Stichprobe', () => {
        const local = v2Payload();
        const remote = v2Payload({ core: { schoolName: 'Andere Schule' } });
        const report = buildBackupCoverageReport(local, remote);
        expect(report.identical).toBe(false);
        expect(report.mismatchSchooltool).toBeGreaterThan(0);
        const nameRow = report.schooltoolRows.find((r) => r.id === 'schoolName');
        expect(nameRow.match).toBe(false);
        const v2Spot = report.spotlight.find((s) => s.key === 'ms365-schooltool-data-v2');
        expect(v2Spot.status).toBe('differs');
    });

    it('meldet identisch bei gleichem Fingerabdruck', () => {
        const a = v2Payload();
        a.contentFingerprint = 'abc:2';
        const b = v2Payload();
        b.contentFingerprint = 'abc:2';
        const cmp = { identical: true, changedKeys: [], onlyLocalKeys: [], onlyRemoteKeys: [] };
        const report = buildBackupCoverageReport(a, b, cmp);
        expect(report.readyToSync).toBe(true);
        expect(compareSchooltoolSummaries(a, b).every((r) => r.match)).toBe(true);
    });
});
