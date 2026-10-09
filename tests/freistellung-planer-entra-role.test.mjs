import { describe, it, expect } from 'vitest';
import {
    entraGroupsConfigured,
    listRolesFromEntraGroups,
    listRolesFromStammdaten,
    resolveActivePlanerRole,
    finalizePlanerRoles,
    planerEntraGroupIds,
    schuelerEntraGroupIdsForCheck,
    direktionEntraGroupIdsForCheck,
    kvEntraGroupIdsForCheck,
    isPlanerDemoRoleUiEnabled
} from '../src/tools/freistellung-planer/freistellung-planer-entra-role.js';
import { matchKvByClassHeadEmail } from '../src/tools/freistellung-planer/freistellung-planer-state.js';

const DIR = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
const KV = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';
const SCH = 'cccccccc-cccc-cccc-cccc-cccccccccccc';

const cfg = {
    groupDirektion: 'Direktion',
    groupDirektionId: DIR,
    groupKv: 'KV',
    groupKvId: KV,
    groupSchueler: 'Schüler',
    groupSchuelerId: SCH
};

describe('freistellung-planer-entra-role', () => {
    it('isPlanerDemoRoleUiEnabled nur mit demoRole-Override', () => {
        expect(isPlanerDemoRoleUiEnabled(false, false)).toBe(false);
        expect(isPlanerDemoRoleUiEnabled(false, true)).toBe(true);
        expect(isPlanerDemoRoleUiEnabled(true, false)).toBe(false);
    });

    it('entraGroupsConfigured', () => {
        expect(entraGroupsConfigured({})).toBe(false);
        expect(entraGroupsConfigured({ groupKvId: KV })).toBe(true);
        expect(
            entraGroupsConfigured({ schuelerUsers: [{ mail: 's@schule.at', id: 'x' }] })
        ).toBe(true);
    });

    it('listRolesFromEntraGroups priorisiert direktion > kv > schueler in Liste', () => {
        const all = new Set([DIR, KV, SCH].map((x) => x.toLowerCase()));
        expect(listRolesFromEntraGroups(all, cfg)).toEqual(['direktion', 'kv', 'schueler']);
        expect(listRolesFromEntraGroups(new Set([KV.toLowerCase()]), cfg)).toEqual(['kv']);
    });

    it('erkennt Schüler über Stammdaten-Sammelgruppe auch ohne groupSchuelerId in Planer-Config', () => {
        const stammSch = 'dddddddd-dddd-dddd-dddd-dddddddddddd';
        const cfgOhneSch = { groupDirektionId: DIR, groupKvId: KV, groupSchuelerId: '' };
        const member = new Set([stammSch.toLowerCase()]);
        const orig = globalThis.localStorage;
        const store = { 'ms365-dashboard-audience-groups-v1': JSON.stringify({ groupSchuelerId: stammSch }) };
        globalThis.localStorage = {
            getItem: (k) => store[k] || null,
            setItem: () => {}
        };
        try {
            expect(schuelerEntraGroupIdsForCheck(cfgOhneSch)).toContain(stammSch);
            expect(listRolesFromEntraGroups(member, cfgOhneSch)).toEqual(['schueler']);
            expect(planerEntraGroupIds(cfgOhneSch)).toContain(stammSch);
        } finally {
            globalThis.localStorage = orig;
        }
    });

    it('erkennt Direktion und KV über zusätzliche Gruppen', () => {
        const extraDir = '11111111-1111-1111-1111-111111111111';
        const extraKv = '22222222-2222-2222-2222-222222222222';
        const cfgExtra = {
            groupDirektionId: DIR,
            groupKvId: KV,
            direktionGroups: [{ groupId: extraDir, groupLabel: 'Sek' }],
            kvGroups: [{ groupId: extraKv, groupLabel: 'KV Extra' }]
        };
        expect(direktionEntraGroupIdsForCheck(cfgExtra)).toEqual(expect.arrayContaining([DIR, extraDir]));
        expect(kvEntraGroupIdsForCheck(cfgExtra)).toEqual(expect.arrayContaining([KV, extraKv]));
        expect(listRolesFromEntraGroups(new Set([extraKv.toLowerCase()]), cfgExtra)).toEqual(['kv']);
        expect(listRolesFromEntraGroups(new Set([extraDir.toLowerCase()]), cfgExtra)).toEqual(['direktion']);
    });

    it('erkennt Schüler über zusätzliche schuelerGroups', () => {
        const extra = 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee';
        const cfgExtra = {
            groupDirektionId: DIR,
            groupKvId: KV,
            groupSchuelerId: SCH,
            schuelerGroups: [{ groupId: extra, groupLabel: 'Extra' }]
        };
        expect(schuelerEntraGroupIdsForCheck(cfgExtra)).toEqual(expect.arrayContaining([SCH, extra]));
        expect(listRolesFromEntraGroups(new Set([extra.toLowerCase()]), cfgExtra)).toEqual(['schueler']);
    });

    it('listRolesFromStammdaten', () => {
        expect(
            listRolesFromStammdaten({ direktionMatch: true, kvMatch: {}, studentMatch: {} })
        ).toEqual(['direktion', 'kv', 'schueler']);
        expect(listRolesFromStammdaten({ studentMatch: {} })).toEqual(['schueler']);
    });

    it('finalizePlanerRoles ergänzt Schüler für Direktion', () => {
        const { roles, sources } = finalizePlanerRoles(['direktion'], { direktion: 'entra' });
        expect(roles).toEqual(['direktion', 'schueler']);
        expect(sources.schueler).toBe('direktion-schueler');
    });

    it('resolveActivePlanerRole', () => {
        expect(resolveActivePlanerRole(['kv', 'schueler'], { preferredActiveRole: 'schueler' })).toBe('schueler');
    });

    it('matchKvByClassHeadEmail', () => {
        const hit = matchKvByClassHeadEmail('kv@test.at', [{ code: '3A', headEmail: 'kv@test.at' }]);
        expect(hit && hit.classCode).toBe('3A');
    });

    it('planerEntraGroupIds', () => {
        expect(planerEntraGroupIds(cfg)).toHaveLength(3);
    });
});
