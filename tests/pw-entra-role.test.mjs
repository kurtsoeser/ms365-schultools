import { describe, it, expect } from 'vitest';
import {
    canUsePwRoleSwitcher,
    isPwDemoRoleUiEnabled,
    sortPwRoles
} from '../src/tools/projektwochen/projektwochen-entra-role.js';

describe('projektwochen-entra-role policy', () => {
    it('Demo nur mit Override', () => {
        expect(isPwDemoRoleUiEnabled(true, false)).toBe(false);
        expect(isPwDemoRoleUiEnabled(false, true)).toBe(true);
    });

    it('Switcher nur Admin+Mehrfach oder Demo', () => {
        expect(canUsePwRoleSwitcher(['lehrer', 'schueler'], false)).toBe(false);
        expect(canUsePwRoleSwitcher(['admin', 'lehrer'], false)).toBe(true);
        expect(canUsePwRoleSwitcher(['schueler'], true)).toBe(true);
    });

    it('sortiert Rollen', () => {
        expect(sortPwRoles(['schueler', 'admin', 'lehrer'])).toEqual([
            'admin',
            'lehrer',
            'schueler'
        ]);
    });
});
