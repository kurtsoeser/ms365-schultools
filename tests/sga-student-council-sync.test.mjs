import { describe, expect, it } from 'vitest';
import {
    resolveStudentRepSyncSource,
    sgaStudentsFromStudentCouncil,
    studentCouncilFromSgaStudents
} from '../src/shared/sga-student-council-sync.js';

describe('sga-student-council-sync', () => {
    const students = [
        { klasse: '3A', name: 'Anna', email: 'anna@schule.at' },
        { klasse: '4B', name: 'Ben', email: 'ben@schule.at' },
        { klasse: '5C', name: 'Carla', email: 'carla@schule.at' }
    ];

    it('baut Schülervertretung aus SGA-Slots inkl. Klasse', () => {
        const rows = studentCouncilFromSgaStudents(
            [
                { name: 'Anna', email: 'anna@schule.at' },
                { name: '', email: 'ben@schule.at' },
                { name: '', email: '' }
            ],
            students,
            [{ klasse: '1A', name: 'Extra', email: 'extra@schule.at' }]
        );
        expect(rows).toEqual([
            { klasse: '3A', name: 'Anna', email: 'anna@schule.at' },
            { klasse: '4B', name: 'Ben', email: 'ben@schule.at' }
        ]);
    });

    it('bewahrt Extra-Plätze der Schülervertretung ab Slot 4', () => {
        const existing = [
            { klasse: 'x', name: 'Alt1', email: 'a@schule.at' },
            { klasse: 'x', name: 'Alt2', email: 'b@schule.at' },
            { klasse: 'x', name: 'Alt3', email: 'c@schule.at' },
            { klasse: '2A', name: 'Dana', email: 'dana@schule.at' }
        ];
        const rows = studentCouncilFromSgaStudents(
            [
                { name: 'Anna', email: 'anna@schule.at' },
                { name: 'Ben', email: 'ben@schule.at' },
                { name: 'Carla', email: 'carla@schule.at' }
            ],
            students,
            existing,
            3
        );
        expect(rows[3]).toEqual({ klasse: '2A', name: 'Dana', email: 'dana@schule.at' });
        expect(rows).toHaveLength(4);
    });

    it('füllt SGA-Slots aus Schülervertretung', () => {
        const slots = sgaStudentsFromStudentCouncil(
            [
                { name: 'Anna', email: 'anna@schule.at' },
                { name: 'Ben', email: 'ben@schule.at' }
            ],
            3
        );
        expect(slots).toEqual([
            { name: 'Anna', email: 'anna@schule.at' },
            { name: 'Ben', email: 'ben@schule.at' },
            { name: '', email: '' }
        ]);
    });

    it('wählt Sync-Quelle beim Laden', () => {
        expect(
            resolveStudentRepSyncSource([], [{ email: 'anna@schule.at' }])
        ).toBe('council');
        expect(
            resolveStudentRepSyncSource([{ email: 'anna@schule.at' }], [])
        ).toBe('sga');
        expect(
            resolveStudentRepSyncSource(
                [{ email: 'anna@schule.at' }],
                [{ email: 'anna@schule.at' }]
            )
        ).toBe('none');
        expect(
            resolveStudentRepSyncSource(
                [{ email: 'anna@schule.at' }],
                [{ email: 'ben@schule.at' }]
            )
        ).toBe('council');
    });
});
