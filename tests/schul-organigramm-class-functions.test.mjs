import { describe, expect, it } from 'vitest';
import {
    addSchulOrgPersonToArea,
    removeSchulOrgPersonFromArea,
    listPeopleForSchulOrgNode,
    normalizeSchulOrgAreaPersonRole
} from '../src/tools/schul-organigramm/schul-organigramm-edit.js';
import {
    SCHUL_ORG_DEFAULT_CLASS_FUNCTIONS,
    classFunctionBoardId,
    collectClassFunctionPeopleFromSettings,
    ensureSchulOrganigrammClassFunctionStructure,
    isClassFunctionArea,
    isClassNode,
    listSchulOrgClassFunctions,
    seedSchulOrganigrammFromStammdaten,
    syncSchulOrganigrammWithStammdaten
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';
import { computeLandkarteLayout } from '../src/tools/schul-organigramm/schul-organigramm-canvas.js';

function classAreaSettings(extra) {
    const base = {
        teachers: [{ code: 'PAGE', name: 'Jimmy Page', email: 'jimmy.page@kurtrocks.com' }],
        classes: [
            {
                code: '1A',
                name: 'DEMO 1a',
                headName: '',
                headEmail: ''
            },
            {
                code: '1B',
                name: 'DEMO 1b',
                headName: '',
                headEmail: ''
            }
        ],
        students: [
            { name: 'Anna Demo', email: 'anna.demo@ms365.schule', klasse: '1A' },
            { name: 'Ben Demo', email: 'ben.demo@ms365.schule', klasse: '1A' },
            { name: 'Clara Demo', email: 'clara.demo@ms365.schule', klasse: '1B' }
        ],
        arges: [],
        admin: [],
        adminRoles: []
    };
    const settings = Object.assign(base, extra || {});
    settings.schulOrganigramm = seedSchulOrganigrammFromStammdaten(settings);
    return settings;
}

describe('Klassenfunktionen (Klassensprecher)', () => {
    it('liefert Builtin Klassensprecher', () => {
        const defs = listSchulOrgClassFunctions();
        expect(defs.some((d) => d.id === 'klassensprecher')).toBe(true);
        expect(SCHUL_ORG_DEFAULT_CLASS_FUNCTIONS[0].id).toBe('klassensprecher');
    });

    it('normalizeSchulOrgAreaPersonRole erkennt Klassensprecher', () => {
        const area = { type: 'area', refKind: 'class', refKey: '1A', id: 'class_1A' };
        expect(normalizeSchulOrgAreaPersonRole('fn:klassensprecher', area)).toBe('fn:klassensprecher');
        expect(normalizeSchulOrgAreaPersonRole('Klassensprecher', area)).toBe('fn:klassensprecher');
        expect(normalizeSchulOrgAreaPersonRole('kv', area)).toBe('kv');
    });

    it('setzt Klassensprecher an der Klasse und listet Sitz', () => {
        const settings = classAreaSettings();
        const area = settings.schulOrganigramm.nodes.find(
            (n) => isClassNode(n) && n.refKey === '1A'
        );
        const res = addSchulOrgPersonToArea(settings, area, {
            name: 'Anna Demo',
            email: 'anna.demo@ms365.schule',
            role: 'fn:klassensprecher'
        });
        expect(res.ok).toBe(true);
        expect(res.role).toBe('fn:klassensprecher');
        expect(res.settings.classes.find((c) => c.code === '1A').functions).toEqual({
            klassensprecher: { name: 'Anna Demo', email: 'anna.demo@ms365.schule' }
        });
        const people = listPeopleForSchulOrgNode(res.settings, area);
        const seat = people.find((p) => p.functionId === 'klassensprecher');
        expect(seat).toMatchObject({
            seatLabel: 'Klassensprecher',
            email: 'anna.demo@ms365.schule'
        });
        expect(seat.vacant).toBeFalsy();
        expect(people.filter((p) => p.email === 'anna.demo@ms365.schule')).toHaveLength(1);
    });

    it('wechselt Klassensprecher und aktualisiert Sammelbereich', () => {
        const settings = classAreaSettings();
        const area = settings.schulOrganigramm.nodes.find(
            (n) => isClassNode(n) && n.refKey === '1A'
        );
        let cur = addSchulOrgPersonToArea(settings, area, {
            name: 'Anna Demo',
            email: 'anna.demo@ms365.schule',
            role: 'fn:klassensprecher'
        }).settings;
        cur = addSchulOrgPersonToArea(cur, area, {
            name: 'Ben Demo',
            email: 'ben.demo@ms365.schule',
            role: 'fn:klassensprecher'
        }).settings;
        expect(cur.classes.find((c) => c.code === '1A').functions.klassensprecher.email).toBe(
            'ben.demo@ms365.schule'
        );
        expect(cur.students.filter((s) => s.klasse === '1A').map((s) => s.email).sort()).toEqual([
            'anna.demo@ms365.schule',
            'ben.demo@ms365.schule'
        ]);

        const boardId = classFunctionBoardId('klassensprecher');
        const board = cur.schulOrganigramm.nodes.find((n) => n.id === boardId);
        expect(isClassFunctionArea(board)).toBe(true);
        const boardPeople = listPeopleForSchulOrgNode(cur, board);
        expect(boardPeople.map((p) => p.email)).toEqual(['ben.demo@ms365.schule']);
    });

    it('sammelt Klassensprecher schulweit ohne Doppelung', () => {
        const settings = {
            classes: [
                {
                    code: '1A',
                    name: 'DEMO 1a',
                    functions: {
                        klassensprecher: { name: 'Anna', email: 'anna@schule.at' }
                    }
                },
                {
                    code: '1B',
                    name: 'DEMO 1b',
                    functions: {
                        klassensprecher: { name: 'Clara', email: 'clara@schule.at' }
                    }
                },
                {
                    code: '2A',
                    name: 'DEMO 2a',
                    functions: {
                        klassensprecher: { name: 'Anna', email: 'anna@schule.at' }
                    }
                }
            ]
        };
        const people = collectClassFunctionPeopleFromSettings(settings, 'klassensprecher');
        expect(people.map((p) => p.email).sort()).toEqual(['anna@schule.at', 'clara@schule.at']);
        const anna = people.find((p) => p.email === 'anna@schule.at');
        expect(anna.className).toMatch(/1a/i);
        expect(anna.className).toMatch(/2a/i);
    });

    it('seedet Root-Bereich Klassensprecher außerhalb Schüler', () => {
        const settings = classAreaSettings({
            classes: [
                {
                    code: '1A',
                    name: 'DEMO 1a',
                    functions: {
                        klassensprecher: {
                            name: 'Anna Demo',
                            email: 'anna.demo@ms365.schule'
                        }
                    }
                }
            ]
        });
        const boardId = classFunctionBoardId('klassensprecher');
        const board = settings.schulOrganigramm.nodes.find((n) => n.id === boardId);
        expect(board).toBeTruthy();
        expect(board.parentId == null || board.parentId === '').toBe(true);
        expect(board.label).toBe('Klassensprecher');
        expect(board.refKind).toBe('classFunction');

        const people = listPeopleForSchulOrgNode(settings, board);
        expect(people).toHaveLength(1);
        expect(people[0].email).toBe('anna.demo@ms365.schule');

        const map = computeLandkarteLayout(settings.schulOrganigramm, {
            peopleCountForArea: function (area) {
                return area && area.id === boardId ? 1 : 0;
            }
        });
        const box = map.boxes.find((b) => b.kind === 'area' && b.node.id === boardId);
        expect(box?.peoplePreview).toBe(true);
    });

    it('Migration legt Klassensprecher-Bereich nach und behält M365', () => {
        const legacy = {
            version: 2,
            nodes: [
                {
                    id: 'custom_ks',
                    type: 'area',
                    label: 'Klassensprecher',
                    parentId: null,
                    order: 3,
                    m365: { groupMode: 'match', groupId: 'gid-ks', groupKind: 'distribution' }
                }
            ]
        };
        const next = ensureSchulOrganigrammClassFunctionStructure(legacy, {});
        const board = next.nodes.find((n) => n.id === classFunctionBoardId('klassensprecher'));
        expect(board).toBeTruthy();
        expect(board.m365.groupId).toBe('gid-ks');
        expect(board.refKind).toBe('classFunction');
    });

    it('entfernt nur die Funktion, Schüler bleibt in der Klasse', () => {
        const settings = classAreaSettings();
        const area = settings.schulOrganigramm.nodes.find(
            (n) => isClassNode(n) && n.refKey === '1A'
        );
        let cur = addSchulOrgPersonToArea(settings, area, {
            name: 'Anna Demo',
            email: 'anna.demo@ms365.schule',
            role: 'fn:klassensprecher'
        }).settings;
        const res = removeSchulOrgPersonFromArea(cur, area, 'anna.demo@ms365.schule');
        expect(res.ok).toBe(true);
        expect(res.message).toMatch(/Klassensprecher/i);
        expect(res.settings.classes.find((c) => c.code === '1A').functions).toBeUndefined();
        expect(
            res.settings.students.find((s) => s.email === 'anna.demo@ms365.schule').klasse
        ).toBe('1A');
    });

    it('sync zieht fehlenden Klassensprecher-Bereich nach', () => {
        const settings = classAreaSettings();
        settings.schulOrganigramm.nodes = settings.schulOrganigramm.nodes.filter(
            (n) => n.id !== classFunctionBoardId('klassensprecher')
        );
        const synced = syncSchulOrganigrammWithStammdaten(settings.schulOrganigramm, settings);
        const board = synced.layout.nodes.find(
            (n) => n.id === classFunctionBoardId('klassensprecher')
        );
        expect(board).toBeTruthy();
        expect(isClassFunctionArea(board)).toBe(true);
    });
});
