import { describe, it, expect } from 'vitest';
import {
    normalizeLfrPermissions,
    accountIsLfrDirektionAccess,
    accountIsLfrLehrerAccess,
    resolveLfrApproverEmail,
    lfrPermsStepConfigured
} from '../src/tools/lehrer-freistellung-planer/lfr-permissions.js';

describe('lfr-permissions', () => {
    it('normalisiert Gruppen und Einzelpersonen', () => {
        const n = normalizeLfrPermissions({
            groupLehrerId: 'g1',
            groupLehrerName: 'Lehrer',
            lehrerUsers: [{ mail: 'a@schule.at', displayName: 'A' }],
            direktionUsers: [{ email: 'dir@schule.at', name: 'Dir' }],
            emailDirektion: 'Dir@schule.at'
        });
        expect(n.groupLehrerId).toBe('g1');
        expect(n.lehrerUsers[0].mail).toBe('a@schule.at');
        expect(n.direktionUsers[0].mail).toBe('dir@schule.at');
        expect(n.emailDirektion).toBe('dir@schule.at');
    });

    it('erkennt Direktion über Genehmiger-Mail und Einzelpersonen', () => {
        const perms = normalizeLfrPermissions({
            emailDirektion: 'chef@schule.at',
            direktionUsers: [{ mail: 'verwaltung@schule.at', displayName: 'Verw' }]
        });
        expect(accountIsLfrDirektionAccess('chef@schule.at', perms)).toBe(true);
        expect(accountIsLfrDirektionAccess('verwaltung@schule.at', perms)).toBe(true);
        expect(accountIsLfrDirektionAccess('lehrer@schule.at', perms)).toBe(false);
    });

    it('erkennt Lehrkraft-Einzelperson', () => {
        const perms = normalizeLfrPermissions({
            lehrerUsers: [{ mail: 'max@schule.at', displayName: 'Max' }]
        });
        expect(accountIsLfrLehrerAccess('max@schule.at', perms)).toBe(true);
        expect(accountIsLfrLehrerAccess('other@schule.at', perms)).toBe(false);
    });

    it('permsStepConfigured zählt auch nur Personen', () => {
        expect(lfrPermsStepConfigured({})).toBe(false);
        expect(
            lfrPermsStepConfigured({
                direktionUsers: [{ mail: 'a@b.at', displayName: 'A' }]
            })
        ).toBe(true);
    });

    it('resolveLfrApproverEmail bevorzugt Perms', () => {
        const perms = normalizeLfrPermissions({ emailDirektion: 'a@schule.at' });
        expect(resolveLfrApproverEmail(perms)).toBe('a@schule.at');
    });
});
