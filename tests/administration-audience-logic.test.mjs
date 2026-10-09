import { describe, expect, it } from 'vitest';
import {
    ADMIN_TIER_SCHULLEITUNG,
    ADMIN_TIER_VERWALTUNG,
    collectAdminEmailsForTier,
    inferAdminTierForRole,
    splitAdminEmailsByAudienceTier
} from '../src/shared/administration-audience-logic.js';

describe('administration-audience-logic', () => {
    const catalog = [
        { code: 'DIREKTION', name: 'Direktion', tier: ADMIN_TIER_SCHULLEITUNG },
        { code: 'SEKRETARIAT', name: 'Sekretariat', tier: ADMIN_TIER_VERWALTUNG }
    ];

    it('Direktion ist Schulleitung, andere Rollen Verwaltung', () => {
        expect(inferAdminTierForRole({ name: 'Direktion' })).toBe(ADMIN_TIER_SCHULLEITUNG);
        expect(inferAdminTierForRole({ name: 'Bibliothek' })).toBe(ADMIN_TIER_VERWALTUNG);
    });

    it('splitAdminEmailsByAudienceTier trennt nach Rolle', () => {
        const admin = [
            { role: 'Direktion', email: 'dir@schule.at' },
            { role: 'Sekretariat', email: 'sek@schule.at' },
            { role: 'IT-Support', email: 'it@schule.at' }
        ];
        const split = splitAdminEmailsByAudienceTier(admin, catalog);
        expect(split.schulleitung).toEqual(['dir@schule.at']);
        expect(split.verwaltung).toContain('sek@schule.at');
        expect(split.verwaltung).toContain('it@schule.at');
    });

    it('collectAdminEmailsForTier dedupliziert', () => {
        const admin = [
            { role: 'Sekretariat', email: 'sek@schule.at' },
            { role: 'Sekretariat', email: 'sek@schule.at' }
        ];
        expect(collectAdminEmailsForTier(admin, catalog, ADMIN_TIER_VERWALTUNG)).toEqual(['sek@schule.at']);
    });
});
