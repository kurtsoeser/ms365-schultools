import { describe, it, expect } from 'vitest';
import {
    normalizeDirektionUsers,
    mergeDirektionUser,
    accountIsDirektionPlannerUser
} from '../src/tools/freistellung-planer/freistellung-planer-direktion-users.js';
import { normalizePermissionsConfig } from '../src/tools/freistellung-planer/freistellung-planer-permissions.js';

describe('freistellung-planer-direktion-users', () => {
    it('normalizeDirektionUsers dedupliziert nach E-Mail', () => {
        const list = normalizeDirektionUsers([
            { mail: 'a@schule.at', displayName: 'A' },
            { email: 'a@schule.at', name: 'A2' },
            { mail: 'b@schule.at', displayName: 'B' }
        ]);
        expect(list).toHaveLength(2);
        expect(list[0].mail).toBe('a@schule.at');
    });

    it('accountIsDirektionPlannerUser', () => {
        const users = [{ mail: 'sek@schule.at', displayName: 'Sek' }];
        expect(accountIsDirektionPlannerUser('sek@schule.at', users)).toBe(true);
        expect(accountIsDirektionPlannerUser('other@schule.at', users)).toBe(false);
    });

    it('mergeDirektionUser', () => {
        const merged = mergeDirektionUser([{ mail: 'a@x.at', displayName: 'A' }], {
            mail: 'b@x.at',
            displayName: 'B'
        });
        expect(merged).toHaveLength(2);
    });

    it('normalizePermissionsConfig behält direktionUsers', () => {
        const cfg = normalizePermissionsConfig({
            groupKvId: 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb',
            direktionUsers: [{ mail: 'dir@schule.at', displayName: 'Dir' }]
        });
        expect(cfg.direktionUsers[0].mail).toBe('dir@schule.at');
    });
});
