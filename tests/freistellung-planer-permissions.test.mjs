import { describe, it, expect } from 'vitest';
import {
    remotePayloadToPermissions,
    permissionsToRemotePayload
} from '../src/tools/freistellung-planer/freistellung-planer-remote-config.js';
import {
    normalizePermissionsConfig,
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
});
