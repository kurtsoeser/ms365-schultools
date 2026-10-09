import { describe, expect, it } from 'vitest';
import {
    ADMIN_ROLE_SEAT_SINGLE,
    adminRoleSeatModeLabel,
    canAddPersonToAdminRole,
    inferAdminRoleSeatMode,
    normalizeAdminRoleM365Kind,
    normalizeAdminRoleSeatMode
} from '../src/shared/administration-role-policy.js';

describe('administration-role-policy', () => {
    it('inferiert Einzelplatz für Direktion und Schularzt', () => {
        expect(inferAdminRoleSeatMode({ name: 'Direktion' })).toBe(ADMIN_ROLE_SEAT_SINGLE);
        expect(inferAdminRoleSeatMode({ name: 'Schularzt' })).toBe(ADMIN_ROLE_SEAT_SINGLE);
        expect(inferAdminRoleSeatMode({ name: 'Sekretariat' })).not.toBe(ADMIN_ROLE_SEAT_SINGLE);
    });

    it('begrenzt Einzelrollen auf eine Person', () => {
        expect(canAddPersonToAdminRole('single', 0)).toBe(true);
        expect(canAddPersonToAdminRole('single', 1)).toBe(false);
        expect(canAddPersonToAdminRole('multi', 3)).toBe(true);
    });

    it('normalisiert M365-Ressourcentyp', () => {
        expect(normalizeAdminRoleM365Kind('sharedMailbox')).toBe('sharedMailbox');
        expect(normalizeAdminRoleM365Kind('group')).toBe('group');
        expect(normalizeAdminRoleM365Kind('')).toBe('none');
    });

    it('liefert lesbare Besetzungs-Kurzlabels', () => {
        expect(adminRoleSeatModeLabel('single')).toBe('1 Platz');
        expect(adminRoleSeatModeLabel('multi')).toBe('Team');
        expect(normalizeAdminRoleSeatMode('team', null)).toBe('multi');
    });
});
