import { describe, it, expect } from 'vitest';
import {
    canUseLfrRoleSwitcher,
    isLfrDemoRoleUiEnabled,
    sortLfrRoles,
    resolveActiveLfrRole
} from '../src/tools/lehrer-freistellung-planer/lfr-entra-role.js';

describe('lfr-entra-role policy', () => {
    it('Demo nur mit Override', () => {
        expect(isLfrDemoRoleUiEnabled(true, false)).toBe(false);
        expect(isLfrDemoRoleUiEnabled(false, true)).toBe(true);
    });

    it('Switcher nur Direktion+Mehrfach oder Demo', () => {
        expect(canUseLfrRoleSwitcher(['lehrer'], false)).toBe(false);
        expect(canUseLfrRoleSwitcher(['direktion'], false)).toBe(false);
        expect(canUseLfrRoleSwitcher(['direktion', 'lehrer'], false)).toBe(true);
        expect(canUseLfrRoleSwitcher(['lehrer'], true)).toBe(true);
    });

    it('sortiert und wählt aktive Rolle', () => {
        expect(sortLfrRoles(['lehrer', 'direktion'])).toEqual(['direktion', 'lehrer']);
        expect(resolveActiveLfrRole(['lehrer', 'direktion'], { preferredActiveRole: 'lehrer' })).toBe(
            'lehrer'
        );
    });
});
