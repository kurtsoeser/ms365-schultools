import { describe, it, expect } from 'vitest';
import {
    parseMs365DemoTeacherAccount,
    parseMs365DemoStudentAccount,
    peopleListHasEmail,
    teachersListHasEmail,
    upsertDemoTeacherInTeachersList,
    upsertDemoStudentInStudentsList
} from '../src/shared/demo-portal-teacher.js';

describe('demo-portal-teacher', () => {
    it('parst Lehrer01 DEMO @ms365.schule', () => {
        expect(parseMs365DemoTeacherAccount('Lehrer01.DEMO@ms365.schule')).toEqual({
            email: 'lehrer01.demo@ms365.schule',
            code: 'L01',
            name: 'Lehrer01 DEMO'
        });
        expect(parseMs365DemoTeacherAccount('schueler01.demo@ms365.schule')).toBeNull();
        expect(parseMs365DemoTeacherAccount('max@schule.at')).toBeNull();
    });

    it('parst Schüler01 DEMO @ms365.schule', () => {
        expect(parseMs365DemoStudentAccount('Schueler01.DEMO@ms365.schule')).toEqual({
            email: 'schueler01.demo@ms365.schule',
            name: 'Schüler01 DEMO',
            klasse: ''
        });
        expect(parseMs365DemoStudentAccount('lehrer01.demo@ms365.schule')).toBeNull();
        expect(parseMs365DemoStudentAccount('anna@schule.at')).toBeNull();
    });

    it('erkennt vorhandene E-Mail in Lehrerliste', () => {
        const teachers = [{ email: 'lehrer01.demo@ms365.schule', code: 'L01' }];
        expect(teachersListHasEmail(teachers, 'Lehrer01.DEMO@ms365.schule')).toBe(true);
        expect(peopleListHasEmail(teachers, 'lehrer02.demo@ms365.schule')).toBe(false);
    });

    it('fügt Demo-Lehrer nur einmal hinzu', () => {
        const account = parseMs365DemoTeacherAccount('lehrer01.demo@ms365.schule');
        const first = upsertDemoTeacherInTeachersList([], account);
        expect(first.changed).toBe(true);
        expect(first.teachers).toHaveLength(1);
        expect(first.teachers[0].email).toBe('lehrer01.demo@ms365.schule');

        const second = upsertDemoTeacherInTeachersList(first.teachers, account);
        expect(second.changed).toBe(false);
        expect(second.teachers).toHaveLength(1);
    });

    it('fügt Demo-Schüler nur einmal hinzu', () => {
        const account = parseMs365DemoStudentAccount('schueler01.demo@ms365.schule');
        const first = upsertDemoStudentInStudentsList([], account);
        expect(first.changed).toBe(true);
        expect(first.students).toHaveLength(1);
        expect(first.students[0]).toEqual({
            name: 'Schüler01 DEMO',
            email: 'schueler01.demo@ms365.schule',
            klasse: ''
        });

        const second = upsertDemoStudentInStudentsList(first.students, account);
        expect(second.changed).toBe(false);
        expect(second.students).toHaveLength(1);
    });
});
