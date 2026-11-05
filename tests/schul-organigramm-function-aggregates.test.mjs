import { describe, expect, it } from 'vitest';
import { listPeopleForSchulOrgNode } from '../src/tools/schul-organigramm/schul-organigramm-edit.js';
import {
    SCHUL_ORG_SEC_ARGE_LEITER,
    SCHUL_ORG_SEC_KLASSENVORSTAND,
    argeFunctionBoardId,
    classFunctionBoardId,
    collectFunctionAggregatePeople,
    ensureSchulOrganigrammFunctionAggregateStructure,
    isFunctionAggregateArea,
    listSchulOrgFunctionAggregates,
    seedSchulOrganigrammFromStammdaten,
    syncSchulOrganigrammWithStammdaten
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';

describe('Funktions-Sammelbereiche (Registry)', () => {
    it('listet KV, ARGE-Leiter und Klassensprecher als Aggregate', () => {
        const defs = listSchulOrgFunctionAggregates();
        const ids = defs.map((d) => d.id);
        expect(ids).toContain('klassenvorstand');
        expect(ids).toContain('argeLeiter');
        expect(ids).toContain('klassensprecher');
        expect(defs.every((d) => d.aggregate)).toBe(true);
    });

    it('nimmt zusätzliche Klassen- und ARGE-Funktionen auf', () => {
        const defs = listSchulOrgFunctionAggregates({
            classFunctions: [{ id: 'kassier', label: 'Kassier', labelPlural: 'Kassiere' }],
            argeFunctions: [
                { id: 'stellvertretung', label: 'ARGE-Stv.', labelPlural: 'ARGE-Stellvertretungen' }
            ]
        });
        expect(defs.some((d) => d.id === 'kassier' && d.refKind === 'classFunction')).toBe(true);
        expect(defs.some((d) => d.seatKey === 'stellvertretung' && d.refKind === 'argeFunction')).toBe(
            true
        );
    });

    it('collectFunctionAggregatePeople bündelt Sitze schulweit', () => {
        const settings = {
            classes: [
                {
                    code: '1A',
                    name: '1A',
                    headName: 'KV A',
                    headEmail: 'kv.a@schule.at',
                    functions: {
                        klassensprecher: { name: 'Anna', email: 'anna@schule.at' },
                        kassier: { name: 'Ben', email: 'ben@schule.at' }
                    }
                },
                {
                    code: '1B',
                    name: '1B',
                    headName: 'KV B',
                    headEmail: 'kv.b@schule.at',
                    functions: {
                        klassensprecher: { name: 'Clara', email: 'clara@schule.at' }
                    }
                }
            ],
            arges: [
                {
                    code: 'MATHE',
                    name: 'Mathe',
                    headName: 'May',
                    headEmail: 'may@schule.at',
                    functions: {
                        stellvertretung: { name: 'Page', email: 'page@schule.at' }
                    }
                }
            ],
            classFunctions: [{ id: 'kassier', label: 'Kassier', labelPlural: 'Kassiere' }],
            argeFunctions: [
                { id: 'stellvertretung', label: 'ARGE-Stv.', labelPlural: 'ARGE-Stellvertretungen' }
            ]
        };
        expect(
            collectFunctionAggregatePeople(settings, 'klassenvorstand').map((p) => p.email).sort()
        ).toEqual(['kv.a@schule.at', 'kv.b@schule.at']);
        expect(
            collectFunctionAggregatePeople(settings, 'klassensprecher').map((p) => p.email).sort()
        ).toEqual(['anna@schule.at', 'clara@schule.at']);
        expect(collectFunctionAggregatePeople(settings, 'kassier').map((p) => p.email)).toEqual([
            'ben@schule.at'
        ]);
        expect(collectFunctionAggregatePeople(settings, 'argeLeiter').map((p) => p.email)).toEqual([
            'may@schule.at'
        ]);
        const stv = listSchulOrgFunctionAggregates(settings).find(
            (d) => d.seatKey === 'stellvertretung'
        );
        expect(collectFunctionAggregatePeople(settings, stv).map((p) => p.email)).toEqual([
            'page@schule.at'
        ]);
    });

    it('ensure legt alle Aggregate an und behält M365', () => {
        const legacy = {
            version: 2,
            nodes: [
                {
                    id: 'custom_ks',
                    type: 'area',
                    label: 'Klassensprecher',
                    parentId: null,
                    m365: { groupMode: 'match', groupId: 'gid-ks', groupKind: 'distribution' }
                }
            ]
        };
        const settings = {
            classFunctions: [{ id: 'kassier', label: 'Kassier', labelPlural: 'Kassiere' }],
            argeFunctions: [
                { id: 'stellvertretung', label: 'ARGE-Stv.', labelPlural: 'ARGE-Stellvertretungen' }
            ]
        };
        const next = ensureSchulOrganigrammFunctionAggregateStructure(legacy, settings);
        expect(next.nodes.find((n) => n.id === SCHUL_ORG_SEC_KLASSENVORSTAND)).toBeTruthy();
        expect(next.nodes.find((n) => n.id === SCHUL_ORG_SEC_ARGE_LEITER)).toBeTruthy();
        const ks = next.nodes.find((n) => n.id === classFunctionBoardId('klassensprecher'));
        expect(ks.m365.groupId).toBe('gid-ks');
        expect(isFunctionAggregateArea(ks)).toBe(true);
        expect(next.nodes.find((n) => n.id === classFunctionBoardId('kassier'))).toBeTruthy();
        expect(next.nodes.find((n) => n.id === argeFunctionBoardId('stellvertretung'))).toBeTruthy();
    });

    it('seed + listPeople: Sammelbereiche sind M365-fähig und personen-gelesen', () => {
        const settings = {
            teachers: [{ name: 'May', email: 'may@schule.at' }],
            classes: [
                {
                    code: '1A',
                    name: 'DEMO 1a',
                    headName: 'May',
                    headEmail: 'may@schule.at',
                    functions: {
                        klassensprecher: { name: 'Anna', email: 'anna@schule.at' }
                    }
                }
            ],
            students: [{ name: 'Anna', email: 'anna@schule.at', klasse: '1A' }],
            arges: [{ code: 'MATHE', name: 'Mathe', headName: 'May', headEmail: 'may@schule.at' }],
            admin: [],
            adminRoles: []
        };
        settings.schulOrganigramm = seedSchulOrganigrammFromStammdaten(settings);
        const kv = settings.schulOrganigramm.nodes.find((n) => n.id === SCHUL_ORG_SEC_KLASSENVORSTAND);
        const ks = settings.schulOrganigramm.nodes.find(
            (n) => n.id === classFunctionBoardId('klassensprecher')
        );
        const al = settings.schulOrganigramm.nodes.find((n) => n.id === SCHUL_ORG_SEC_ARGE_LEITER);
        expect(listPeopleForSchulOrgNode(settings, kv)[0].email).toBe('may@schule.at');
        expect(listPeopleForSchulOrgNode(settings, ks)[0].email).toBe('anna@schule.at');
        expect(listPeopleForSchulOrgNode(settings, al)[0].email).toBe('may@schule.at');
    });

    it('sync zieht neue Funktions-Aggregate nach', () => {
        const settings = {
            teachers: [],
            classes: [{ code: '1A', name: '1A' }],
            students: [],
            arges: [],
            admin: [],
            adminRoles: [],
            classFunctions: [{ id: 'sportwart', label: 'Sportwart', labelPlural: 'Sportwarte' }]
        };
        settings.schulOrganigramm = seedSchulOrganigrammFromStammdaten({
            teachers: [],
            classes: [{ code: '1A', name: '1A' }],
            students: [],
            arges: [],
            admin: [],
            adminRoles: []
        });
        const synced = syncSchulOrganigrammWithStammdaten(settings.schulOrganigramm, settings);
        expect(
            synced.layout.nodes.find((n) => n.id === classFunctionBoardId('sportwart'))
        ).toBeTruthy();
    });
});
