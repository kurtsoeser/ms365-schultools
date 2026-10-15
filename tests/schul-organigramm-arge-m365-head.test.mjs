import { describe, expect, it } from 'vitest';
import {
    applyArgeHeadFromGraphOwners,
    listPeopleForSchulOrgNode
} from '../src/tools/schul-organigramm/schul-organigramm-edit.js';
import {
    collectArgeLeiterPeopleFromSettings,
    isArgeAreaNode,
    seedSchulOrganigrammFromStammdaten
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';

describe('ARGE-Bereich = M365-Gruppe: Besitzer → Leitung', () => {
    it('setzt ARGE-Leitung aus Graph-Besitzer (Lehrkraft bevorzugen)', () => {
        const settings = {
            teachers: [
                { code: 'MUEL', name: 'Müller', email: 'mueller@schule.at' },
                { code: 'SCHM', name: 'Schmidt', email: 'schmidt@schule.at' }
            ],
            arges: [{ code: 'MATHE', name: 'DEMO-ARGE Mathematik', headName: '', headEmail: '' }],
            classes: [],
            admin: [],
            adminRoles: []
        };
        const applied = applyArgeHeadFromGraphOwners(settings, 'MATHE', [
            { displayName: 'Extern', mail: 'extern@schule.at' },
            { displayName: 'Müller Graph', mail: 'mueller@schule.at' }
        ]);
        expect(applied.changed).toBe(true);
        expect(applied.head).toEqual({ name: 'Müller', email: 'mueller@schule.at' });
        expect(applied.settings.arges[0]).toMatchObject({
            headName: 'Müller',
            headEmail: 'mueller@schule.at'
        });
    });

    it('ändert nichts bei gleichem Stand', () => {
        const settings = {
            teachers: [{ name: 'Müller', email: 'mueller@schule.at' }],
            arges: [
                {
                    code: 'MATHE',
                    name: 'Mathe',
                    headName: 'Müller',
                    headEmail: 'mueller@schule.at'
                }
            ]
        };
        const applied = applyArgeHeadFromGraphOwners(settings, 'MATHE', [
            { displayName: 'Müller', mail: 'mueller@schule.at' }
        ]);
        expect(applied.changed).toBe(false);
    });

    it('ARGE-Leiter-Bereich listet gesetzte Leitungen', () => {
        const settings = {
            teachers: [{ name: 'Müller', email: 'mueller@schule.at' }],
            arges: [
                {
                    code: 'MATHE',
                    name: 'DEMO-ARGE Mathematik',
                    headName: 'Müller',
                    headEmail: 'mueller@schule.at'
                },
                { code: 'MUSIK', name: 'DEMO-ARGE Musik', headName: '', headEmail: '' }
            ],
            classes: [],
            admin: [],
            adminRoles: []
        };
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const leiter = layout.nodes.find((n) => n.id === 'sec_arge_leiter');
        expect(leiter).toBeTruthy();
        const people = collectArgeLeiterPeopleFromSettings(settings);
        expect(people.map((p) => p.email)).toEqual(['mueller@schule.at']);
        expect(people[0].argeName).toMatch(/Mathematik/i);
    });

    it('ARGE-Bereich listet Leitung mit Sitz-Label ARGE-Leitung', () => {
        const settings = {
            teachers: [{ name: 'Müller', email: 'mueller@schule.at' }],
            arges: [
                {
                    code: 'MATHE',
                    name: 'DEMO-ARGE Mathematik',
                    headName: 'Müller',
                    headEmail: 'mueller@schule.at'
                }
            ],
            classes: [],
            admin: [],
            adminRoles: []
        };
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const argeArea = layout.nodes.find(
            (n) => isArgeAreaNode(n) && String(n.refKey || '').toUpperCase() === 'MATHE'
        );
        expect(argeArea).toBeTruthy();
        const people = listPeopleForSchulOrgNode(settings, argeArea);
        expect(people[0]).toMatchObject({
            kind: 'head',
            seatLabel: 'ARGE-Leitung',
            email: 'mueller@schule.at',
            name: 'Müller'
        });
    });

    it('ARGE-Fachlehrer aus Stammdaten tragen Funktion ARGE-Mitglieder', () => {
        const settings = {
            teachers: [
                { name: 'Müller', email: 'mueller@schule.at' },
                { name: 'Lea', email: 'lea@schule.at' }
            ],
            arges: [
                {
                    code: 'MATHE',
                    name: 'DEMO-ARGE Mathematik',
                    headName: 'Müller',
                    headEmail: 'mueller@schule.at',
                    subjects: ['MAT']
                }
            ],
            unterrichtsbelegung: {
                rows: [
                    {
                        klasse: '1A',
                        fach: 'MAT',
                        lehrerEmail: 'lea@schule.at',
                        lehrerCode: 'LEA'
                    }
                ]
            },
            classes: [],
            admin: [],
            adminRoles: []
        };
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const argeArea = layout.nodes.find(
            (n) => isArgeAreaNode(n) && String(n.refKey || '').toUpperCase() === 'MATHE'
        );
        const people = listPeopleForSchulOrgNode(settings, argeArea);
        const lea = people.find((p) => p.email === 'lea@schule.at');
        expect(lea).toMatchObject({
            kind: 'teacher',
            seatLabel: 'ARGE-Mitglieder'
        });
    });
});
