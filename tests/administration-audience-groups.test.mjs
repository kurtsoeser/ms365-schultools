import { describe, expect, it } from 'vitest';
import {
    addCustomVerwaltungAudienceGroup,
    buildMembershipsFromRoleTiers,
    collectEmailsForAudienceGroup,
    ensureAdminAudienceOnSettings,
    findAdminDisplayRowByEmail,
    setAudienceGroupsForEmail
} from '../src/shared/administration-audience-groups.js';
import { ADMIN_TIER_SCHULLEITUNG } from '../src/shared/administration-audience-logic.js';

describe('administration-audience-groups', () => {
    it('migriert Tier-Modell in Mitgliedschaften', () => {
        const admin = [
            { role: 'Direktion', email: 'dir@schule.at', name: 'Dir' },
            { role: 'Sekretariat', email: 'sek@schule.at', name: 'Sek' }
        ];
        const roles = [
            { name: 'Direktion', code: 'DIR', tier: ADMIN_TIER_SCHULLEITUNG },
            { name: 'Sekretariat', code: 'SEK', tier: 'verwaltung' }
        ];
        const m = buildMembershipsFromRoleTiers(admin, roles);
        expect(collectEmailsForAudienceGroup(m, 'schulleitung')).toEqual(['dir@schule.at']);
        expect(collectEmailsForAudienceGroup(m, 'verwaltung')).toEqual(['sek@schule.at']);
    });

    it('eine Person in mehreren Gruppen', () => {
        const added = addCustomVerwaltungAudienceGroup('Schupersonal', []);
        const groups = added.groups;
        const id = added.id;
        let m = setAudienceGroupsForEmail('it@schule.at', ['verwaltung', id], 'IT', []);
        expect(collectEmailsForAudienceGroup(m, 'verwaltung')).toContain('it@schule.at');
        expect(collectEmailsForAudienceGroup(m, id)).toContain('it@schule.at');
    });

    it('findAdminDisplayRowByEmail', () => {
        const rows = [
            { name: 'Sekretariat', code: 'SEK', personName: 'Anna', email: 'sek@schule.at' }
        ];
        expect(findAdminDisplayRowByEmail(rows, 'sek@schule.at')?.code).toBe('SEK');
        expect(findAdminDisplayRowByEmail(rows, 'x@schule.at')).toBeNull();
    });

    it('ensureAdminAudienceOnSettings füllt leere Mitgliedschaften aus admin', () => {
        const st = ensureAdminAudienceOnSettings({
            admin: [{ role: 'Direktion', email: 'dir@schule.at' }],
            adminRoles: [{ name: 'Direktion', code: 'DIR', tier: ADMIN_TIER_SCHULLEITUNG }]
        });
        expect(st.memberships.length).toBeGreaterThan(0);
        expect(st.groups.some((g) => g.id === 'schulleitung')).toBe(true);
    });
});
