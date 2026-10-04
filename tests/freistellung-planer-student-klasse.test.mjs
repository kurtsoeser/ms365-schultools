import { describe, it, expect } from 'vitest';
import {
    listClassGraphGroupIds,
    matchStudentKlasseFromMemberGroupIds
} from '../src/tools/freistellung-planer/freistellung-planer-student-klasse.js';
import { prefillStudentFreistellungForm, emptyForm } from '../src/tools/freistellung-planer/freistellung-planer-state.js';

describe('freistellung student klasse entra', () => {
    it('matchStudentKlasseFromMemberGroupIds über classTeams', () => {
        const gid = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
        const hit = matchStudentKlasseFromMemberGroupIds(
            new Set([gid]),
            [{ graphGroupId: gid, classCode: '1A', displayName: '1A' }],
            {}
        );
        expect(hit && hit.klasse).toBe('1A');
    });

    it('listClassGraphGroupIds sammelt Teams und classGroupMatchByKey', () => {
        const id1 = '11111111-2222-3333-4444-555555555555';
        const id2 = '66666666-7777-8888-9999-aaaaaaaaaaaa';
        const ids = listClassGraphGroupIds(
            { classGroupMatchByKey: { '2B': { groupId: id2 } } },
            [{ graphGroupId: id1 }]
        );
        expect(ids).toContain(id1);
        expect(ids).toContain(id2);
    });

    it('prefill überschreibt Demo-Klasse bei Stammdaten-Treffer', () => {
        const state = {
            role: 'schueler',
            demoKlasseCode: '3A',
            studentMatch: { email: 's@schule.at', klasse: '1A' },
            stammdaten: {
                classes: [{ code: '1A', name: '1A', headEmail: 'kv@schule.at' }]
            },
            form: { ...emptyForm(), klasse: '3A' }
        };
        prefillStudentFreistellungForm(state);
        expect(state.form.klasse).toBe('1A');
        expect(state.demoKlasseCode).toBe('');
    });
});
