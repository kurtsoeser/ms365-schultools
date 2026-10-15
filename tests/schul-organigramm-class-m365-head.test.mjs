import { describe, expect, it } from 'vitest';
import {
    applyClassHeadFromGraphOwners,
    listPeopleForSchulOrgNode
} from '../src/tools/schul-organigramm/schul-organigramm-edit.js';
import { seedSchulOrganigrammFromStammdaten, isClassNode } from '../src/tools/schul-organigramm/schul-organigramm-logic.js';

describe('Klasse = M365-Gruppe: Besitzer → Klassenvorstand', () => {
    it('setzt Klassenvorstand aus Graph-Besitzer (Lehrkraft bevorzugen)', () => {
        const settings = {
            teachers: [{ code: 'PAGE', name: 'Jimmy Page', email: 'page@schule.at' }],
            classes: [
                {
                    code: '4A',
                    name: 'DEMO Klasse 4A',
                    headName: '',
                    headEmail: ''
                }
            ],
            arges: [],
            admin: [],
            adminRoles: []
        };
        const applied = applyClassHeadFromGraphOwners(settings, '4A', [
            { displayName: 'Extern Graph', mail: 'extern@schule.at' },
            { displayName: 'Jimmy Graph', mail: 'page@schule.at' }
        ]);
        expect(applied.changed).toBe(true);
        expect(applied.head).toEqual({ name: 'Jimmy Page', email: 'page@schule.at' });
        expect(applied.settings.classes[0]).toMatchObject({
            headName: 'Jimmy Page',
            headEmail: 'page@schule.at'
        });
    });

    it('ändert nichts bei gleichem Stand', () => {
        const settings = {
            teachers: [{ name: 'Jimmy Page', email: 'page@schule.at' }],
            classes: [
                {
                    code: '4A',
                    name: '4A',
                    headName: 'Jimmy Page',
                    headEmail: 'page@schule.at'
                }
            ]
        };
        const applied = applyClassHeadFromGraphOwners(settings, '4A', [
            { displayName: 'Jimmy Page', mail: 'page@schule.at' }
        ]);
        expect(applied.changed).toBe(false);
    });

    it('Landkarte listet Klassenvorstand als head mit Sitz-Label', () => {
        const settings = {
            teachers: [{ code: 'PAGE', name: 'Jimmy Page', email: 'page@schule.at' }],
            classes: [
                {
                    code: '4A',
                    name: 'DEMO Klasse 4A',
                    headName: 'Jimmy Page',
                    headEmail: 'page@schule.at'
                }
            ],
            students: [
                { name: 'Schüler19 DEMO', email: 's19@schule.at', klasse: '4A' }
            ],
            arges: [],
            admin: [],
            adminRoles: []
        };
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const classNode = layout.nodes.find((n) => isClassNode(n) && n.refKey === '4A');
        expect(classNode).toBeTruthy();
        const people = listPeopleForSchulOrgNode(settings, classNode);
        expect(people[0]).toMatchObject({
            kind: 'head',
            seatLabel: 'Klassenvorstand',
            email: 'page@schule.at',
            name: 'Jimmy Page'
        });
        expect(people[0].vacant).toBeFalsy();
        const student = people.find((p) => p.kind === 'student');
        expect(student).toMatchObject({
            email: 's19@schule.at',
            seatLabel: 'Schüler'
        });
    });
});
