import { describe, it, expect } from 'vitest';
import {
    remotePayloadToPermissions,
    permissionsToRemotePayload
} from '../src/tools/freistellung-planer/freistellung-planer-remote-config.js';
import {
    normalizePermissionsConfig,
    normalizePlannerExtraEntraGroups,
    normalizeSchuelerGroups,
    entraGroupsConfigured
} from '../src/tools/freistellung-planer/freistellung-planer-permissions.js';

describe('freistellung-planer-permissions / remote', () => {
    it('remotePayloadToPermissions', () => {
        const p = remotePayloadToPermissions({
            groupSchuelerId: 'cccccccc-cccc-cccc-cccc-cccccccccccc',
            groupSchueler: 'Schüler'
        });
        expect(entraGroupsConfigured(p)).toBe(true);
    });

    it('permissionsToRemotePayload enthält version', () => {
        const raw = permissionsToRemotePayload(
            normalizePermissionsConfig({ groupKvId: 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' })
        );
        expect(raw.version).toBe(1);
        expect(raw.groupKvId).toContain('bbbb');
    });

    it('normalizeSchuelerGroups und Remote-Payload', () => {
        const extra = 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee';
        const c = normalizePermissionsConfig({
            groupSchuelerId: 'cccccccc-cccc-cccc-cccc-cccccccccccc',
            schuelerGroups: [{ groupId: extra, groupLabel: 'Extra SuS' }, { groupId: 'bad' }]
        });
        expect(normalizeSchuelerGroups(c.schuelerGroups)).toEqual([
            { groupId: extra, groupLabel: 'Extra SuS' }
        ]);
        const raw = permissionsToRemotePayload(c);
        expect(raw.schuelerGroups).toEqual([{ groupId: extra, groupLabel: 'Extra SuS' }]);
        expect(entraGroupsConfigured({ schuelerGroups: [{ groupId: extra }] })).toBe(true);
    });
});
