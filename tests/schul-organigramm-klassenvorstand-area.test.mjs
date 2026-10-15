import { describe, expect, it } from 'vitest';
import { listPeopleForSchulOrgNode } from '../src/tools/schul-organigramm/schul-organigramm-edit.js';
import {
    SCHUL_ORG_SEC_KLASSENVORSTAND,
    SCHUL_ORG_SEC_SCHUELER,
    collectKlassenvorstandPeopleFromSettings,
    ensureSchulOrganigrammKlassenvorstandStructure,
    isKlassenvorstandArea,
    seedSchulOrganigrammFromStammdaten,
    syncSchulOrganigrammWithStammdaten
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';
import { computeLandkarteLayout } from '../src/tools/schul-organigramm/schul-organigramm-canvas.js';

describe('Bereich Klassenvorstände (Root, analog ARGE-Leiter)', () => {
    it('sammelt Klassenvorstände aus Stammdaten (ohne Doppelung)', () => {
        const settings = {
            teachers: [
                { name: 'Jimmy Page', email: 'page@schule.at' },
                { name: 'Brian May', email: 'may@schule.at' }
            ],
            classes: [
                {
                    code: '4A',
                    name: 'DEMO Klasse 4A',
                    headName: 'Jimmy Page',
                    headEmail: 'page@schule.at'
                },
                {
                    code: '4B',
                    name: 'DEMO Klasse 4B',
                    headName: 'Jimmy Page',
                    headEmail: 'page@schule.at'
                },
                {
                    code: '5A',
                    name: 'DEMO Klasse 5A',
                    headName: 'Brian May',
                    headEmail: 'may@schule.at'
                },
                { code: '5B', name: 'Ohne KV', headName: '', headEmail: '' }
            ]
        };
        const people = collectKlassenvorstandPeopleFromSettings(settings);
        expect(people.map((p) => p.email).sort()).toEqual(['may@schule.at', 'page@schule.at']);
        const page = people.find((p) => p.email === 'page@schule.at');
        expect(page.className).toMatch(/4A/);
        expect(page.className).toMatch(/4B/);
    });

    it('seedet Root-Bereich Klassenvorstände außerhalb Schüler', () => {
        const settings = {
            teachers: [{ name: 'Jimmy Page', email: 'page@schule.at' }],
            classes: [
                {
                    code: '4A',
                    name: 'DEMO Klasse 4A',
                    headName: 'Jimmy Page',
                    headEmail: 'page@schule.at'
                }
            ],
            arges: [],
            admin: [],
            adminRoles: [],
            students: []
        };
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const board = layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_KLASSENVORSTAND);
        expect(board).toBeTruthy();
        expect(isKlassenvorstandArea(board)).toBe(true);
        expect(board.parentId == null || board.parentId === '').toBe(true);
        expect(board.label).toBe('Klassenvorstände');
        expect(board.parentId).not.toBe(SCHUL_ORG_SEC_SCHUELER);

        const people = listPeopleForSchulOrgNode(settings, board);
        expect(people).toHaveLength(1);
        expect(people[0]).toMatchObject({
            email: 'page@schule.at',
            seatLabel: 'Klassenvorstand',
            kind: 'teacher'
        });

        const map = computeLandkarteLayout(layout, {
            peopleCountForArea: function (area) {
                return area && area.id === SCHUL_ORG_SEC_KLASSENVORSTAND ? 1 : 0;
            }
        });
        const box = map.boxes.find(
            (b) => b.kind === 'area' && b.node.id === SCHUL_ORG_SEC_KLASSENVORSTAND
        );
        expect(box?.peoplePreview).toBe(true);
        expect(box?.parentAreaId == null || box?.parentAreaId === '').toBe(true);
    });

    it('Migration legt Klassenvorstände-Bereich nach und behält M365', () => {
        const legacy = {
            version: 2,
            nodes: [
                {
                    id: 'custom_kv',
                    type: 'area',
                    label: 'Klassenvorstände',
                    parentId: SCHUL_ORG_SEC_SCHUELER,
                    m365: {
                        groupMode: 'match',
                        groupId: 'gid-kv-demo',
                        groupLabel: 'KV-Gruppe'
                    }
                }
            ]
        };
        const ensured = ensureSchulOrganigrammKlassenvorstandStructure(legacy, {});
        const board = ensured.nodes.find((n) => n.id === SCHUL_ORG_SEC_KLASSENVORSTAND);
        expect(board).toBeTruthy();
        expect(board.parentId == null || board.parentId === '').toBe(true);
        expect(board.m365?.groupId).toBe('gid-kv-demo');
        expect(ensured.nodes.some((n) => n.id === 'custom_kv')).toBe(false);

        const synced = syncSchulOrganigrammWithStammdaten(legacy, {
            classes: [
                {
                    code: '1A',
                    name: '1A',
                    headName: 'May',
                    headEmail: 'may@schule.at'
                }
            ],
            teachers: [{ name: 'May', email: 'may@schule.at' }],
            arges: [],
            admin: [],
            adminRoles: []
        });
        expect(
            synced.layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_KLASSENVORSTAND)
        ).toBeTruthy();
    });
});
