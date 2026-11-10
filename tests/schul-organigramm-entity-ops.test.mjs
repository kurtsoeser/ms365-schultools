import { describe, expect, it } from 'vitest';
import {
    NODE,
    elternClassNodeId,
    klassenchatClassNodeId,
    removeArgeEntityFromSettings,
    removeClassEntityFromSettings,
    renameClassEntityInSettings,
    renameSchulOrganigrammClassEntity,
    seedSchulOrganigrammFromStammdaten,
    stripSchulOrganigrammClassEntity
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';

describe('schul-organigramm entity ops (P3)', () => {
    const baseSettings = {
        classes: [
            { code: '1A', name: '1A', headEmail: 'kv@schule.at' },
            { code: '2B', name: '2B' }
        ],
        students: [
            { id: 's1', klasse: '1A', name: 'Lea', email: 'lea@schule.at', guardianIds: ['g1'] },
            { id: 's2', klasse: '2B', name: 'Tom', email: 'tom@schule.at', guardianIds: [] }
        ],
        studentCouncil: [{ klasse: '1A', name: 'Lea', email: 'lea@schule.at' }],
        guardians: [{ id: 'g1', name: 'Eltern', email: 'e@familie.at' }],
        teachers: [],
        arges: [{ code: 'MATHE', name: 'Mathematik' }],
        admin: [],
        adminRoles: [],
        sga: [],
        classChats: {
            version: 1,
            items: [{ klasse: '1A', kind: 'groupChat', chatId: 'c1', topic: '1A' }]
        },
        unterrichtsbelegung: {
            rows: [
                { klasse: '1A', fach: 'MAT', lehrerEmail: 'l@schule.at' },
                { klasse: '2B', fach: 'D', lehrerEmail: 'l@schule.at' }
            ]
        }
    };

    it('renameClassEntityInSettings hängt Schüler, Council, Chats, Belegung und Landkarte um', () => {
        const seeded = seedSchulOrganigrammFromStammdaten(baseSettings);
        const settings = Object.assign({}, baseSettings, { schulOrganigramm: seeded });
        const result = renameClassEntityInSettings(settings, '1A', '1AX', { name: '1AX neu' });
        expect(result.renamedStudents).toBe(1);
        expect(result.renamedCouncil).toBe(1);
        expect(result.renamedChats).toBe(1);
        expect(result.renamedBelegung).toBe(1);
        expect(result.settings.classes.find((c) => c.code === '1AX')?.name).toBe('1AX neu');
        expect(result.settings.students.find((s) => s.id === 's1')?.klasse).toBe('1AX');
        expect(result.settings.classChats.items[0].klasse).toBe('1AX');
        expect(result.settings.unterrichtsbelegung.rows[0].klasse).toBe('1AX');
        const nodes = result.settings.schulOrganigramm.nodes;
        expect(nodes.some((n) => n.id === 'class_1AX' && n.refKey === '1AX')).toBe(true);
        expect(nodes.some((n) => n.id === 'class_1A')).toBe(false);
        expect(nodes.some((n) => n.id === elternClassNodeId('1AX'))).toBe(true);
        expect(nodes.some((n) => n.id === klassenchatClassNodeId('1AX'))).toBe(true);
    });

    it('removeClassEntityFromSettings entfernt Klasse/Chat/Knoten, behält Schüler', () => {
        const seeded = seedSchulOrganigrammFromStammdaten(baseSettings);
        const settings = Object.assign({}, baseSettings, { schulOrganigramm: seeded });
        const result = removeClassEntityFromSettings(settings, '1A');
        expect(result.studentCount).toBe(1);
        expect(result.removedChat).toBe(true);
        expect(result.settings.classes.map((c) => c.code)).toEqual(['2B']);
        expect(result.settings.students.find((s) => s.id === 's1')?.klasse).toBe('1A');
        expect(result.settings.classChats.items).toHaveLength(0);
        const nodes = result.settings.schulOrganigramm.nodes;
        expect(nodes.some((n) => n.id === 'class_1A')).toBe(false);
        expect(nodes.some((n) => n.id === elternClassNodeId('1A'))).toBe(false);
        expect(nodes.some((n) => n.id === klassenchatClassNodeId('1A'))).toBe(false);
        expect(nodes.some((n) => n.id === 'class_2B')).toBe(true);
    });

    it('stripSchulOrganigrammClassEntity entfernt Kursteams unter der Klasse', () => {
        const layout = {
            version: 2,
            nodes: [
                {
                    id: 'class_1A',
                    type: NODE.SECTION,
                    refKind: 'class',
                    refKey: '1A',
                    parentId: 'sec_jahrgang_1',
                    label: '1A',
                    order: 1
                },
                {
                    id: 'kurs_x@1A',
                    type: NODE.KURSTEAM,
                    refKind: 'kursteam',
                    refKey: 'x@1A',
                    parentId: 'class_1A',
                    label: 'MAT',
                    order: 2,
                    meta: { placementClass: '1A' }
                }
            ]
        };
        const next = stripSchulOrganigrammClassEntity(layout, '1A');
        expect(next.nodes).toHaveLength(0);
    });

    it('renameSchulOrganigrammClassEntity schreibt Parent-IDs um', () => {
        const layout = {
            version: 2,
            nodes: [
                {
                    id: 'class_1A',
                    type: NODE.SECTION,
                    refKind: 'class',
                    refKey: '1A',
                    parentId: null,
                    label: '1A',
                    order: 1
                },
                {
                    id: 'child',
                    type: NODE.SECTION,
                    refKind: 'none',
                    refKey: '',
                    parentId: 'class_1A',
                    label: 'Unter',
                    order: 2
                }
            ]
        };
        const next = renameSchulOrganigrammClassEntity(layout, '1A', '9Z', '9Z');
        expect(next.nodes.find((n) => n.id === 'class_9Z')?.refKey).toBe('9Z');
        expect(next.nodes.find((n) => n.id === 'child')?.parentId).toBe('class_9Z');
    });

    it('removeArgeEntityFromSettings entfernt Liste und Knoten', () => {
        const seeded = seedSchulOrganigrammFromStammdaten(baseSettings);
        const settings = Object.assign({}, baseSettings, { schulOrganigramm: seeded });
        const next = removeArgeEntityFromSettings(settings, 'MATHE');
        expect(next.arges).toHaveLength(0);
        expect(next.schulOrganigramm.nodes.some((n) => n.refKind === 'arge' && n.refKey === 'MATHE')).toBe(
            false
        );
    });
});
