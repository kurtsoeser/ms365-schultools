import { describe, it, expect } from 'vitest';
import {
    permissionsToListDescriptionPayload,
    parseListDescriptionMarkerFull,
    LIST_DESCRIPTION_MARKER
} from '../src/tools/freistellung-planer/freistellung-planer-remote-config.js';
import { normalizePermissionsConfig } from '../src/tools/freistellung-planer/freistellung-planer-permissions.js';
import { classesForStudentPicker } from '../src/tools/freistellung-planer/freistellung-planer-state.js';

describe('freistellung class team links (SharePoint)', () => {
    const gid = '11111111-2222-3333-4444-555555555555';

    it('serialisiert ct in Listen-Beschreibung und expandiert zurück', () => {
        const cfg = normalizePermissionsConfig({
            groupSchuelerId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
            classTeamLinks: [{ code: '1A', groupId: gid, name: '1A' }],
            classCatalog: [{ code: '1A', name: '1A' }, { code: '1B', name: '1B' }]
        });
        const compact = permissionsToListDescriptionPayload(cfg, {
            siteWebUrl: 'https://x.sharepoint.com/sites/s',
            listId: 'list-1'
        });
        expect(compact.ct).toEqual([{ c: '1A', g: gid }]);
        expect(compact.cl).toEqual([{ c: '1A' }, { c: '1B' }]);
        const parsed = parseListDescriptionMarkerFull(
            LIST_DESCRIPTION_MARKER + JSON.stringify(compact)
        );
        expect(parsed.permissions.classTeamLinks[0].code).toBe('1A');
        expect(parsed.permissions.classTeamLinks[0].groupId).toBe(gid);
        expect(parsed.permissions.classCatalog.map((c) => c.code)).toEqual(['1A', '1B']);
    });

    it('classesForStudentPicker nutzt classCatalog und ignoriert 1AHW', () => {
        const orig = globalThis.localStorage;
        const store = {};
        globalThis.localStorage = {
            getItem(k) {
                return store[k] ?? null;
            },
            setItem(k, v) {
                store[k] = v;
            }
        };
        store['ms365-freistellung-perms-v1'] = JSON.stringify({
            groupSchuelerId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
            classTeamLinks: [{ code: '1A', groupId: gid }],
            classCatalog: [{ code: '1A', name: '1A' }, { code: '1B', name: '1B' }]
        });
        const list = classesForStudentPicker({
            stammdaten: { classes: [], students: [] },
            klasseColumnChoices: [
                { code: '1AHW', name: '1AHW' },
                { code: '2AHW', name: '2AHW' }
            ],
            items: []
        });
        globalThis.localStorage = orig;
        expect(list.map((c) => c.code)).toEqual(['1A', '1B']);
    });
});
