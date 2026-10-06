import { describe, it, expect } from 'vitest';
import { groupStudentsByClass, describeClassM365Match } from '../src/shared/tenant-klassen-hub-logic.js';

describe('tenant-klassen-hub-logic', () => {
    it('gruppiert Schüler nach Klassencode', () => {
        const classes = [{ code: '1AK' }, { code: '2BK' }];
        const students = [
            { klasse: '1AK', name: 'Anna' },
            { klasse: '1ak', name: 'Ben' },
            { klasse: '9ZZ', name: 'Ohne Klasse' }
        ];
        const g = groupStudentsByClass(classes, students, (v) => String(v).trim());
        expect(g.byClass).toHaveLength(2);
        expect(g.byClass[0].students).toHaveLength(2);
        expect(g.unassigned).toHaveLength(1);
    });

    it('beschreibt M365-Match', () => {
        expect(describeClassM365Match({ groupId: 'x', displayName: 'JG 1AK' }).kind).toBe('ok');
        expect(describeClassM365Match({ notFound: true }).kind).toBe('error');
        expect(describeClassM365Match(null).kind).toBe('muted');
    });
});
