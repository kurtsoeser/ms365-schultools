import { describe, expect, it } from 'vitest';
import {
    addSchulOrgPersonToArea,
    removeSchulOrgPersonFromArea,
    listPeopleForSchulOrgNode,
    normalizeSchulOrgAreaPersonRole,
    ensureClassHeadsNamedFromTeachers,
    healClassHeadsFromAreaPeople
} from '../src/tools/schul-organigramm/schul-organigramm-edit.js';
import {
    seedSchulOrganigrammFromStammdaten,
    isClassNode,
    upsertSchulOrganigrammAreaPerson,
    NODE
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';

function classAreaSettings(extra) {
    const base = {
        teachers: [{ code: 'PAGE', name: 'Jimmy Page', email: 'jimmy.page@kurtrocks.com' }],
        classes: [
            {
                code: '4A',
                name: 'DEMO Klasse 4A',
                headName: '',
                headEmail: ''
            },
            {
                code: '4B',
                name: 'DEMO Klasse 4B',
                headName: '',
                headEmail: 'jimmy.page@kurtrocks.com'
            }
        ],
        students: [],
        arges: [],
        admin: [],
        adminRoles: []
    };
    const settings = Object.assign(base, extra || {});
    settings.schulOrganigramm = seedSchulOrganigrammFromStammdaten(settings);
    return settings;
}

describe('Klasse: Funktion/Rolle beim Personen-Zuordnen', () => {
    it('normalizeSchulOrgAreaPersonRole erkennt KV und Schüler', () => {
        const area = { type: 'area', refKind: 'class', refKey: '4A', id: 'class_4A' };
        expect(normalizeSchulOrgAreaPersonRole('kv', area)).toBe('kv');
        expect(normalizeSchulOrgAreaPersonRole('Klassenvorstand', area)).toBe('kv');
        expect(normalizeSchulOrgAreaPersonRole('student', area)).toBe('student');
        expect(normalizeSchulOrgAreaPersonRole('', area)).toBe('student');
    });

    it('setzt Person mit role=kv als Klassenvorstand in Stammdaten', () => {
        const settings = classAreaSettings();
        const area = settings.schulOrganigramm.nodes.find(
            (n) => isClassNode(n) && n.refKey === '4A'
        );
        const res = addSchulOrgPersonToArea(settings, area, {
            name: 'Jimmy Page',
            email: 'jimmy.page@kurtrocks.com',
            role: 'kv'
        });
        expect(res.ok).toBe(true);
        expect(res.role).toBe('kv');
        expect(res.settings.classes.find((c) => c.code === '4A')).toMatchObject({
            headName: 'Jimmy Page',
            headEmail: 'jimmy.page@kurtrocks.com'
        });
        const people = listPeopleForSchulOrgNode(res.settings, area);
        expect(people[0]).toMatchObject({
            kind: 'head',
            seatLabel: 'Klassenvorstand',
            email: 'jimmy.page@kurtrocks.com'
        });
    });

    it('setzt Person mit role=student als Schüler:in der Klasse', () => {
        const settings = classAreaSettings();
        const area = settings.schulOrganigramm.nodes.find(
            (n) => isClassNode(n) && n.refKey === '4A'
        );
        const res = addSchulOrgPersonToArea(settings, area, {
            name: 'Schüler19 DEMO',
            email: 'schueler19.demo@ms365.schule',
            role: 'student'
        });
        expect(res.ok).toBe(true);
        expect(res.role).toBe('student');
        expect(res.settings.students).toEqual([
            {
                name: 'Schüler19 DEMO',
                email: 'schueler19.demo@ms365.schule',
                klasse: '4A'
            }
        ]);
        const people = listPeopleForSchulOrgNode(res.settings, area);
        const student = people.find((p) => p.kind === 'student');
        expect(student).toMatchObject({
            seatLabel: 'Schüler',
            email: 'schueler19.demo@ms365.schule'
        });
    });

    it('entfernt Klassenvorstand und Schüler:in korrekt', () => {
        const settings = classAreaSettings({
            classes: [
                {
                    code: '4A',
                    name: 'DEMO Klasse 4A',
                    headName: 'Jimmy Page',
                    headEmail: 'jimmy.page@kurtrocks.com'
                }
            ],
            students: [
                { name: 'Schüler19 DEMO', email: 's19@ms365.schule', klasse: '4A' }
            ]
        });
        const area = settings.schulOrganigramm.nodes.find(
            (n) => isClassNode(n) && n.refKey === '4A'
        );
        const rmKv = removeSchulOrgPersonFromArea(
            settings,
            area,
            'jimmy.page@kurtrocks.com'
        );
        expect(rmKv.ok).toBe(true);
        expect(rmKv.settings.classes[0].headEmail).toBe('');
        const rmStu = removeSchulOrgPersonFromArea(rmKv.settings, area, 's19@ms365.schule');
        expect(rmStu.settings.students[0].klasse).toBe('');
    });

    it('ensureClassHeadsNamedFromTeachers füllt fehlende headNames', () => {
        const settings = {
            teachers: [{ name: 'Jimmy Page', email: 'jimmy.page@kurtrocks.com' }],
            classes: [
                {
                    code: '4B',
                    name: 'DEMO Klasse 4B',
                    headName: '',
                    headEmail: 'jimmy.page@kurtrocks.com'
                }
            ]
        };
        const named = ensureClassHeadsNamedFromTeachers(settings);
        expect(named.changed).toBe(1);
        expect(named.settings.classes[0].headName).toBe('Jimmy Page');
    });

    it('healClassHeadsFromAreaPeople befördert Lehrkraft aus Area-Slot zum KV', () => {
        let settings = classAreaSettings();
        const area = settings.schulOrganigramm.nodes.find(
            (n) => isClassNode(n) && n.refKey === '4A'
        );
        settings.schulOrganigramm = upsertSchulOrganigrammAreaPerson(
            settings.schulOrganigramm,
            area.id,
            { email: 'jimmy.page@kurtrocks.com', name: 'Jimmy Page' }
        );
        const healed = healClassHeadsFromAreaPeople(settings);
        expect(healed.changed).toBeGreaterThanOrEqual(1);
        const cls = healed.settings.classes.find((c) => c.code === '4A');
        expect(cls).toMatchObject({
            headEmail: 'jimmy.page@kurtrocks.com',
            headName: 'Jimmy Page'
        });
        const people = listPeopleForSchulOrgNode(healed.settings, area);
        expect(people[0]).toMatchObject({
            kind: 'head',
            seatLabel: 'Klassenvorstand',
            email: 'jimmy.page@kurtrocks.com'
        });
        const leftover = (healed.settings.schulOrganigramm.nodes || []).filter(
            (n) =>
                n.type === NODE.PERSON &&
                n.parentId === area.id &&
                String(n.refKey || '').toLowerCase() === 'jimmy.page@kurtrocks.com'
        );
        expect(leftover).toHaveLength(0);
    });
});
