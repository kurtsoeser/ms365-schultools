import { describe, it, expect } from 'vitest';
import { buildMergePlan, applyLocalMerge } from '../src/tools/klassen-merge/klassen-merge-logic.js';
import { analyzeUserNaming, analyzeUsersNaming, buildAdExportCsv } from '../src/shared/naming-convention-audit.js';
import { diffStudentAttributes } from '../src/shared/stammdaten-health.js';
import { parseTermineCsvText, parseTerminTarget, terminToListFields } from '../src/tools/termin-import/termin-import-logic.js';

describe('klassen-merge', () => {
    const classes = [
        { code: '4HMA', name: '4HMA', year: '2027', headEmail: 'a@x.at' },
        { code: '4HMB', name: '4HMB', year: '2027', headEmail: 'b@x.at' }
    ];
    const students = [
        { name: 'A', email: 'a.s@x.at', klasse: '4HMA' },
        { name: 'B', email: 'b.s@x.at', klasse: '4HMB' }
    ];
    const classTeams = [
        { classCode: '4HMA', stableMailNickname: 'jg20274hma', graphGroupId: 'g1' },
        { classCode: '4HMB', stableMailNickname: 'jg20274hmb', graphGroupId: 'g2' }
    ];

    it('baut Plan und merged lokal', () => {
        const plan = buildMergePlan({
            survivor: classes[0],
            survivorOriginalCode: '4HMA',
            sources: [classes[1]],
            newCode: '4HM',
            newName: '4HM',
            newDisplayName: 'Klasse 4HM',
            students,
            classTeams,
            sourceAction: 'archive'
        });
        expect(plan.ok).toBe(true);
        expect(plan.memberEmails).toHaveLength(2);
        expect(plan.steps.some((s) => s.id === 'graph-members')).toBe(true);

        const applied = applyLocalMerge(plan, { classes, students }, classTeams);
        expect(applied.error).toBe('');
        expect(applied.settings.classes.map((c) => c.code)).toEqual(['4HM']);
        expect(applied.settings.students.every((s) => s.klasse === '4HM')).toBe(true);
        expect(applied.classTeams).toHaveLength(1);
        expect(applied.classTeams[0].classCode).toBe('4HM');
    });
});

describe('naming-convention-audit', () => {
    it('erkennt Nummern-DisplayName bei Cloud-User', () => {
        const r = analyzeUserNaming({
            id: '1',
            displayName: '8801',
            givenName: 'Anna',
            surname: 'Beispiel',
            userPrincipalName: '8801@x.at',
            onPremisesSyncEnabled: false
        });
        expect(r.severity).toBe('cloud_fix');
        expect(r.patch.displayName).toBe('Anna Beispiel');
    });

    it('AD-sync → Export statt Patch', () => {
        const r = analyzeUserNaming({
            id: '2',
            displayName: '8801',
            givenName: 'Max',
            surname: 'Muster',
            userPrincipalName: '8801@x.at',
            onPremisesSyncEnabled: true
        });
        expect(r.severity).toBe('ad_export');
        expect(r.action).toBe('export');
    });

    it('baut CSV', () => {
        const { rows } = analyzeUsersNaming([
            {
                id: '2',
                displayName: '8801',
                givenName: 'Max',
                surname: 'Muster',
                userPrincipalName: '8801@x.at',
                onPremisesSyncEnabled: true
            }
        ]);
        const csv = buildAdExportCsv(rows);
        expect(csv).toContain('Max');
        expect(csv.split(/\r?\n/).length).toBe(2);
    });
});

describe('stammdaten-health', () => {
    it('findet Klassen-Mismatch', () => {
        const diff = diffStudentAttributes(
            [{ name: 'A', email: 'a@x.at', klasse: '5BK' }],
            [{ id: 'u1', mail: 'a@x.at', department: '4BK', displayName: 'A' }]
        );
        expect(diff.summary.mismatch).toBe(1);
        expect(diff.rows[0].status).toBe('mismatch');
    });
});

describe('termin-import', () => {
    it('parsed CSV mit Ziel', () => {
        const text = 'Titel;Beginn;Ende;Ziel\nSchulanfang;08.09.2026;;S\nKonferenz;10.09.2026 14:00;10.09.2026 16:00;L';
        const { rows, summary } = parseTermineCsvText(text);
        expect(summary.ok).toBe(2);
        expect(rows[0].target).toBe('school');
        expect(rows[1].target).toBe('teachers');
        expect(parseTerminTarget('beide')).toBe('both');
        const fields = terminToListFields(rows[0]);
        expect(fields.Title).toBe('Schulanfang');
        expect(fields.Beginn).toContain('2026-09-08');
    });
});
