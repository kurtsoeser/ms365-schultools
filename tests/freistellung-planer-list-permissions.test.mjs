import { describe, it, expect } from 'vitest';
import { mapFreistellungConfigForSpo } from '../src/tools/freistellung-planer/freistellung-planer-list-permissions.js';

describe('freistellung-planer-list-permissions', () => {
    it('mapFreistellungConfigForSpo', () => {
        const m = mapFreistellungConfigForSpo({
            groupDirektion: 'Dir',
            groupDirektionId: 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa',
            groupKv: 'KV',
            groupKvId: 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb',
            groupSchueler: 'SuS',
            groupSchuelerId: 'cccccccc-cccc-cccc-cccc-cccccccccccc'
        });
        expect(m.groupAdminId).toBe('aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa');
        expect(m.groupLehrerId).toBe('bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb');
        expect(m.groupSchuelerId).toBe('cccccccc-cccc-cccc-cccc-cccccccccccc');
    });
});
