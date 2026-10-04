import { describe, it, expect } from 'vitest';
import {
    entraGroupsConfigured,
    listRolesFromEntraGroups,
    listRolesFromStammdaten,
    resolveActivePlanerRole,
    finalizePlanerRoles,
    planerEntraGroupIds,
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
