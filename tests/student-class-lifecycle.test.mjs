import { describe, expect, it } from 'vitest';
import {
    diffStudents,
    hasMembershipWork,
    previewMemberships,
    reconcileClassMembers,
    reconcileSammelgruppe,
    studentKey,
    summarizePreview
} from '../src/shared/student-class-lifecycle.js';

describe('student-class-lifecycle', () => {
    it('studentKey bevorzugt E-Mail', () => {
        expect(studentKey({ email: 'A@schule.at', name: 'Ada', klasse: '1A' })).toBe('e:a@schule.at');
        expect(studentKey({ name: 'Ada', klasse: '1A' })).toBe('n:ada|1a');
    });

    it('diffStudents erkennt Zu-, Abgänge und Klassenwechsel', () => {
        const prev = [
            { name: 'Ada', email: 'ada@s.at', klasse: '1A' },
            { name: 'Ben', email: 'ben@s.at', klasse: '1A' },
            { name: 'Cara', email: 'cara@s.at', klasse: '1B' }
        ];
        const next = [
            { name: 'Ada', email: 'ada@s.at', klasse: '2A' },
            { name: 'Cara', email: 'cara@s.at', klasse: '1B' },
            { name: 'Dana', email: 'dana@s.at', klasse: '1A' }
        ];
        const d = diffStudents(prev, next);
        expect(d.added.map((s) => s.email)).toEqual(['dana@s.at']);
        expect(d.removed.map((s) => s.email)).toEqual(['ben@s.at']);
        expect(d.classChanged).toHaveLength(1);
        expect(d.classChanged[0].fromClass).toBe('1A');
        expect(d.classChanged[0].toClass).toBe('2A');
        expect(d.classChanged[0].student.email).toBe('ada@s.at');
    });

    it('previewMemberships plant Join/Leave inkl. Sammelgruppe', () => {
        const diff = diffStudents(
            [
                { email: 'alt@s.at', klasse: '1A', name: 'Alt' },
                { email: 'move@s.at', klasse: '1A', name: 'Move' }
            ],
            [
                { email: 'neu@s.at', klasse: '1B', name: 'Neu' },
                { email: 'move@s.at', klasse: '1B', name: 'Move' }
            ]
        );
        const teams = [
            { classCode: '1A', graphGroupId: 'g-1a', displayName: 'Klasse 1A' },
            { classCode: '1B', graphGroupId: 'g-1b', displayName: 'Klasse 1B' }
        ];
        const preview = previewMemberships(diff, teams, 'g-alle');
        const byId = Object.fromEntries(preview.groups.map((g) => [g.groupId, g]));
        expect(byId['g-1a'].leave).toEqual(expect.arrayContaining(['alt@s.at', 'move@s.at']));
        expect(byId['g-1b'].join).toEqual(expect.arrayContaining(['neu@s.at', 'move@s.at']));
        expect(byId['g-alle'].join).toContain('neu@s.at');
        expect(byId['g-alle'].leave).toContain('alt@s.at');
        expect(byId['g-alle'].join).not.toContain('move@s.at');
        expect(hasMembershipWork(preview)).toBe(true);
        const sum = summarizePreview(preview);
        expect(sum.groupCount).toBe(3);
        expect(sum.join).toBeGreaterThan(0);
        expect(sum.leave).toBeGreaterThan(0);
    });

    it('reconcileClassMembers lässt Nicht-Schüler unberührt', () => {
        const r = reconcileClassMembers(
            ['ada@s.at'],
            ['ada@s.at', 'ben@s.at'],
            ['ada@s.at', 'ben@s.at', 'kv@s.at']
        );
        expect(r.join).toEqual([]);
        expect(r.leave).toEqual(['ben@s.at']);
    });

    it('reconcileSammelgruppe nimmt neue auf und entfernt ehemalige', () => {
        const r = reconcileSammelgruppe(['ada@s.at', 'neu@s.at'], ['ada@s.at', 'alt@s.at']);
        expect(r.join).toEqual(['neu@s.at']);
        expect(r.leave).toEqual(['alt@s.at']);
    });
});
