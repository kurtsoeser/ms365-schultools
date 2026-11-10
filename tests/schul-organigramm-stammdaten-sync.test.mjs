import { describe, expect, it } from 'vitest';
import {
    projectStudentParentPairs,
    settingsFromAppContainer,
    syncSchulOrganigrammFromAppData
} from '../src/shared/schul-organigramm-stammdaten-sync.js';
import {
    NODE,
    seedSchulOrganigrammFromStammdaten,
    setSchulOrganigrammNodePosition
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';

describe('schul-organigramm-stammdaten-sync', () => {
    it('settingsFromAppContainer mappt Core + Jahr + Belegung', () => {
        const settings = settingsFromAppContainer({
            core: {
                teachers: [{ code: 'MUEL', name: 'Müller', email: 'm@schule.at' }],
                arges: [{ code: 'MATHE', name: 'Mathe' }],
                schulOrganigramm: { version: 2, nodes: [] }
            },
            years: {
                current: '2025/26',
                byLabel: {
                    '2025/26': {
                        classes: [{ code: '1AK', name: '1AK' }],
                        students: [{ klasse: '1AK', name: 'Lea', email: 'lea@schule.at' }],
                        unterrichtsbelegung: {
                            rows: [
                                {
                                    klasse: '1AK',
                                    fach: 'MAT',
                                    lehrerEmail: 'm@schule.at',
                                    teamName: 'MAT 1AK',
                                    gruppenmail: 'jg-1ak-mat'
                                }
                            ]
                        }
                    }
                }
            }
        });
        expect(settings.classes).toHaveLength(1);
        expect(settings.teachers[0].code).toBe('MUEL');
        expect(settings.unterrichtsbelegung.rows).toHaveLength(1);
        expect(settings.arges[0].code).toBe('MATHE');
    });

    it('settingsFromAppContainer liefert guardians, classChats und parentPairs-Projektion', () => {
        const settings = settingsFromAppContainer({
            core: { teachers: [], arges: [], schulOrganigramm: { version: 2, nodes: [] } },
            years: {
                current: '2025/26',
                byLabel: {
                    '2025/26': {
                        classes: [{ code: '1A', name: '1A' }],
                        students: [
                            {
                                id: 's1',
                                klasse: '1A',
                                name: 'Lea',
                                email: 'lea@schule.at',
                                guardianIds: ['g1', 'g2']
                            }
                        ],
                        guardians: [
                            { id: 'g1', name: 'Eltern Lea', email: 'eltern.lea@familie.at' },
                            { id: 'g2', name: 'Vater Lea', email: 'vater.lea@familie.at' }
                        ],
                        classChats: {
                            version: 1,
                            items: [{ klasse: '1A', kind: 'groupChat', chatId: 'chat-1' }]
                        }
                    }
                }
            }
        });
        expect(settings.guardians).toHaveLength(2);
        expect(settings.classChats?.items?.[0]?.chatId).toBe('chat-1');
        expect(settings.students[0].parentPairs).toEqual([
            { name: 'Eltern Lea', email: 'eltern.lea@familie.at' },
            { name: 'Vater Lea', email: 'vater.lea@familie.at' }
        ]);
    });

    it('projectStudentParentPairs belässt vorhandene parentPairs', () => {
        const out = projectStudentParentPairs(
            [
                {
                    id: 's1',
                    klasse: '1A',
                    guardianIds: ['g1'],
                    parentPairs: [{ name: 'Alt', email: 'alt@familie.at' }]
                }
            ],
            [{ id: 'g1', name: 'Neu', email: 'neu@familie.at' }]
        );
        expect(out[0].parentPairs).toEqual([{ name: 'Alt', email: 'alt@familie.at' }]);
    });

    it('Materializer schreibt privateChannel-Meta an KlassenChat-Knoten', () => {
        const container = {
            version: 4,
            core: {
                schoolName: 'Test',
                domain: 'schule.at',
                subjects: [],
                arges: [],
                teachers: [],
                admin: [],
                adminRoles: [],
                sga: [],
                schulOrganigramm: { version: 2, nodes: [] }
            },
            years: {
                current: '2025/26',
                byLabel: {
                    '2025/26': {
                        classes: [{ code: '2B', name: '2B' }],
                        students: [],
                        studentCouncil: [],
                        classChats: {
                            version: 1,
                            items: [
                                {
                                    klasse: '2B',
                                    kind: 'privateChannel',
                                    channelId: 'ch-2b',
                                    teamId: 'team-x',
                                    topic: '2B'
                                }
                            ]
                        }
                    }
                }
            },
            setup: {}
        };
        let saved = null;
        globalThis.window = {
            ms365AppDataV2: {
                getContainer: () => container,
                setContainer: (c) => {
                    saved = c;
                    return c;
                }
            }
        };
        const result = syncSchulOrganigrammFromAppData({ reason: 'classChats' });
        expect(result.ok).toBe(true);
        const chat = saved.core.schulOrganigramm.nodes.find(
            (n) => n.refKind === 'klassenchatClass' && n.refKey === '2B'
        );
        expect(chat?.meta?.chatKind).toBe('privateChannel');
        expect(chat?.meta?.channelId).toBe('ch-2b');
        expect(chat?.meta?.teamId).toBe('team-x');
    });

    it('syncSchulOrganigrammFromAppData schreibt Kursteams ins Organigramm', () => {
        const container = {
            version: 4,
            core: {
                schoolName: 'Test',
                domain: 'schule.at',
                subjects: [],
                arges: [],
                teachers: [{ code: 'MUEL', name: 'Müller', email: 'm@schule.at' }],
                admin: [],
                adminRoles: [],
                sga: [],
                schulOrganigramm: { version: 2, nodes: [] }
            },
            years: {
                current: '2025/26',
                byLabel: {
                    '2025/26': {
                        classes: [
                            { code: '1AK', name: '1AK', year: '2034' },
                            { code: '1BK', name: '1BK', year: '2034' }
                        ],
                        students: [],
                        studentCouncil: [],
                        unterrichtsbelegung: {
                            rows: [
                                {
                                    klasse: '1AK1BK',
                                    fach: 'MAT',
                                    lehrerEmail: 'm@schule.at',
                                    teamName: 'MAT geteilt',
                                    gruppenmail: 'jg-shared-mat',
                                    graphGroupId: 'gid-1'
                                }
                            ]
                        }
                    }
                }
            },
            setup: {}
        };

        let saved = null;
        globalThis.window = {
            ms365AppDataV2: {
                getContainer: () => container,
                setContainer: (c) => {
                    saved = c;
                    return c;
                }
            }
        };

        const result = syncSchulOrganigrammFromAppData({ reason: 'test' });
        expect(result.ok).toBe(true);
        expect(saved).toBeTruthy();
        const kurs = saved.core.schulOrganigramm.nodes.filter((n) => n.type === NODE.KURSTEAM);
        expect(kurs.length).toBe(2);
        expect(kurs.every((n) => n.parentId === 'class_1AK' || n.parentId === 'class_1BK')).toBe(true);
        expect(
            saved.core.schulOrganigramm.nodes.some(
                (n) =>
                    (n.type === NODE.CLASS || n.refKind === 'class') &&
                    (n.refKey === '1AK' || n.refKey === '1BK')
            )
        ).toBe(true);
    });

    it('sync behält Root-Landkarten-Orte (Speichern darf nicht zurücksetzen)', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [{ code: 'MAT', name: 'Mathe' }],
            classes: [{ code: '1A', name: '1A' }],
            teachers: [],
            students: [],
            sga: [],
            studentCouncil: []
        });
        layout = setSchulOrganigrammNodePosition(layout, 'sec_arge', 1600, 220);
        layout = setSchulOrganigrammNodePosition(layout, 'sec_schueler', 80, 640);
        layout = setSchulOrganigrammNodePosition(layout, 'sec_lehrer', 1100, 40);

        const container = {
            version: 4,
            core: {
                schoolName: 'Test',
                domain: 'schule.at',
                subjects: [],
                arges: [{ code: 'MAT', name: 'Mathe' }],
                teachers: [],
                admin: [],
                adminRoles: [],
                sga: [],
                schulOrganigramm: layout
            },
            years: {
                current: '2025/26',
                byLabel: {
                    '2025/26': {
                        classes: [{ code: '1A', name: '1A' }],
                        students: [],
                        studentCouncil: []
                    }
                }
            },
            setup: {}
        };

        let saved = null;
        globalThis.window = {
            ms365AppDataV2: {
                getContainer: () => container,
                setContainer: (c) => {
                    saved = c;
                    Object.assign(container, c);
                    return c;
                }
            }
        };

        const result = syncSchulOrganigrammFromAppData({ reason: 'stammdaten' });
        expect(result.ok).toBe(true);
        expect(saved.core.schulOrganigramm.nodes.find((n) => n.id === 'sec_arge')?.x).toBe(1600);
        expect(saved.core.schulOrganigramm.nodes.find((n) => n.id === 'sec_arge')?.y).toBe(220);
        expect(saved.core.schulOrganigramm.nodes.find((n) => n.id === 'sec_schueler')?.y).toBe(640);
        expect(saved.core.schulOrganigramm.nodes.find((n) => n.id === 'sec_lehrer')?.x).toBe(1100);
    });
});
