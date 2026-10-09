import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest';
import {
    loadSchoolAdminGroups,
    overlaySchoolAdminGroupsOnPermissions,
    preferredItStammdatenGroupId
} from '../src/shared/school-admin-groups.js';

const SL = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
const VW = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';

describe('school-admin-groups', () => {
    beforeEach(() => {
        vi.stubGlobal('ms365AppDataV2', {
            getSetup: () => ({
                matched: { schulleitungGroupId: SL, verwaltungGroupId: VW },
                catalogLinks: [
                    { kind: 'sammelgruppe', code: 'schulleitung', displayName: 'Schulleitung' },
                    { kind: 'sammelgruppe', code: 'verwaltung', displayName: 'Personal' }
                ]
            }),
            getCatalogLink: () => null
        });
    });

    afterEach(() => {
        vi.unstubAllGlobals();
    });

    it('overlay setzt groupAdmin und groupDirektion auf Schulleitung', () => {
        const out = overlaySchoolAdminGroupsOnPermissions({ groupAdminId: '', groupDirektionId: '' });
        expect(out.groupAdminId).toBe(SL);
        expect(out.groupDirektionId).toBe(SL);
        expect(out.groupVerwaltungStaffId).toBe(VW);
    });

    it('bestehende groupAdminId bleibt', () => {
        const out = overlaySchoolAdminGroupsOnPermissions({ groupAdminId: 'existing-id' });
        expect(out.groupAdminId).toBe('existing-id');
    });

    it('preferredItStammdatenGroupId bevorzugt Schulleitung', () => {
        expect(preferredItStammdatenGroupId()).toBe(SL);
    });

    it('loadSchoolAdminGroups liest Namen', () => {
        const g = loadSchoolAdminGroups();
        expect(g.schulleitungGroupName).toBe('Schulleitung');
    });
});
