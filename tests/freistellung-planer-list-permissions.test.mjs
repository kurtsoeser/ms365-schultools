import { describe, it, expect } from 'vitest';
import {
    mapFreistellungConfigForSpo,
    FREISTELLUNG_LIST_PROFILE,
    FREISTELLUNG_LIST_ITEM_LEVEL,
    grantFreistellungFlowServiceAccountOnList
} from '../src/tools/freistellung-planer/freistellung-planer-list-permissions.js';

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

    it('Schüler nur eigene Elemente, KV Bearbeiten', () => {
        expect(FREISTELLUNG_LIST_PROFILE.schueler).toBe('contribute');
        expect(FREISTELLUNG_LIST_PROFILE.lehrer).toBe('edit');
        expect(FREISTELLUNG_LIST_ITEM_LEVEL.readSecurity).toBe(2);
        expect(FREISTELLUNG_LIST_ITEM_LEVEL.writeSecurity).toBe(2);
    });

    it('grantFreistellungFlowServiceAccountOnList ohne Konto überspringt', async () => {
        const logs = [];
        const r = await grantFreistellungFlowServiceAccountOnList(
            'https://x.sharepoint.com/sites/a',
            'Freistellungen',
            '',
            {},
            (m) => logs.push(m)
        );
        expect(r.skipped).toBe(true);
        expect(r.reason).toBe('no-account');
        expect(logs.some((l) => /Technik-Konto fehlt/.test(l))).toBe(true);
    });
});
