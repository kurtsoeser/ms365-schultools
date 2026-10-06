import { describe, expect, it } from 'vitest';
import { studentsToLines, sortStudentRows } from '../src/shared/tenant-settings-ui-students.js';
import { classesToLines } from '../src/shared/tenant-settings-ui-classes.js';
import { summarizeKlassenM365 } from '../src/shared/tenant-panel-klassen.js';
import { summarizeStudentsByClass } from '../src/shared/tenant-panel-schueler.js';

describe('tenant-panel helpers', () => {
    it('studentsToLines inkl. Eltern und externalId', () => {
        const text = studentsToLines([
            {
                klasse: '1A',
                name: 'Max',
                email: 'max@schule.at',
                externalId: '99',
                parentPairs: [{ name: 'Eltern', email: 'e@schule.at' }]
            }
        ]);
        expect(text).toContain('#id:99');
        expect(text).toContain('e@schule.at');
    });

    it('sortStudentRows sortiert nach Name', () => {
        const sorted = sortStudentRows(
            [
                { name: 'Zed', klasse: '1A' },
                { name: 'Ann', klasse: '1A' }
            ],
            'name',
            1
        );
        expect(sorted[0].name).toBe('Ann');
    });

    it('classesToLines behält Abschlussjahr', () => {
        expect(classesToLines([{ code: '1A', year: '2030', name: 'Eins A' }])).toContain('2030');
    });

    it('summarizeKlassenM365 zählt Verknüpfungen', () => {
        const s = summarizeKlassenM365([{ code: '1A' }, { code: '1B' }], function (code) {
            if (code === '1A') return { groupId: 'g1' };
            return { notFound: true };
        });
        expect(s).toEqual({ total: 2, linked: 1, missing: 1, unchecked: 0 });
    });

    it('summarizeStudentsByClass findet Unzugeordnete', () => {
        const s = summarizeStudentsByClass({
            getClasses: () => [{ code: '1A' }],
            getStudents: () => [
                { klasse: '1A', name: 'A' },
                { klasse: '9Z', name: 'B' }
            ],
            normClassCode: (v) => String(v || '').trim()
        });
        expect(s.unassigned).toBe(1);
        expect(s.studentCount).toBe(2);
    });
});
