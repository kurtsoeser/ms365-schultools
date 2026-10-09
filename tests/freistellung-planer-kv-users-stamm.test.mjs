import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { kvUsersFromSchoolStammdaten } from '../src/tools/freistellung-planer/freistellung-planer-list-permissions.js';

describe('kvUsersFromSchoolStammdaten', () => {
    beforeEach(() => {
        vi.stubGlobal('ms365TenantSettingsLoad', () => ({
            classes: [{ code: '1A', headName: 'Brian May', headEmail: 'brian@schule.at' }]
        }));
    });

    afterEach(() => {
        vi.unstubAllGlobals();
    });

    it('sammelt eindeutige KV-E-Mails aus Tenant-Klassen', () => {
        const users = kvUsersFromSchoolStammdaten();
        expect(users).toEqual([
            { id: '', displayName: 'Brian May', mail: 'brian@schule.at' }
        ]);
    });
});
