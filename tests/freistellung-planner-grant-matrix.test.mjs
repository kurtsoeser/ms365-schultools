import { describe, it, expect } from 'vitest';
import {
    migrateLegacyFreistellungPerms,
    compileGrantRowsToLegacyFields,
    mergeDuplicateGrantRows
} from '../src/tools/freistellung-planer/freistellung-planner-grant-matrix.js';

describe('freistellung-planner-grant-matrix', () => {
    it('migrateLegacy erzeugt zusammengeführte Zeilen', () => {
        const rows = migrateLegacyFreistellungPerms({
            groupKvId: 'kv-1',
            groupKv: 'KV Gruppe',
            kvUsers: [{ mail: 'kv@schule.at', displayName: 'KV' }]
        });
        expect(rows.length).toBe(2);
        const merged = mergeDuplicateGrantRows([
            ...rows,
            {
                principalType: 'group',
                groupId: 'kv-1',
                groupLabel: 'KV Gruppe',
                mail: '',
                displayName: '',
                roles: { direktion: true, kv: true, schueler: false }
            }
        ]);
        expect(merged.length).toBe(2);
        expect(merged.find((r) => r.groupId === 'kv-1').roles.direktion).toBe(true);
        expect(merged.find((r) => r.groupId === 'kv-1').roles.kv).toBe(true);
    });

    it('compile schreibt Legacy-Felder zurück', () => {
        const compiled = compileGrantRowsToLegacyFields([
            {
                principalType: 'group',
                groupId: 'a1',
                groupLabel: 'IT',
                mail: '',
                displayName: '',
                roles: { admin: true, direktion: false, kv: false, schueler: false }
            },
            {
                principalType: 'group',
                groupId: 'd1',
                groupLabel: 'Dir',
                mail: '',
                displayName: '',
                roles: { admin: false, direktion: true, kv: false, schueler: false }
            },
            {
                principalType: 'user',
                groupId: '',
                groupLabel: '',
                mail: 'a@b.at',
                displayName: 'A',
                roles: { admin: false, direktion: false, kv: true, schueler: false }
            }
        ]);
        expect(compiled.groupAdminId).toBe('a1');
        expect(compiled.groupDirektionId).toBe('d1');
        expect(compiled.kvUsers.length).toBe(1);
        expect(compiled.kvUsers[0].mail).toBe('a@b.at');
    });
});
