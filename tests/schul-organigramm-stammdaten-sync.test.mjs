import { describe, expect, it } from 'vitest';
import {
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
