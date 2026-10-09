import { describe, it, expect } from 'vitest';
import {
    emailDomainFromAddress,
    collectAllowedKursteamOwnerDomains,
    findForeignOwnerDomains,
    validateKursteamBackendTenantContext,
    parseVerifiedEmailDomainsField
} from '../src/shared/kursteam-owner-domain-policy.js';

describe('kursteam-owner-domain-policy', () => {
    it('emailDomainFromAddress', () => {
        expect(emailDomainFromAddress('kurt@kurtsoeser.at')).toBe('kurtsoeser.at');
        expect(emailDomainFromAddress('')).toBe('');
    });

    it('collectAllowedKursteamOwnerDomains merges login, school and stammdaten', () => {
        const allowed = collectAllowedKursteamOwnerDomains({
            loginUpn: 'kurt@kurtsoeser.at',
            schoolDomainNoAt: 'ms365.schule',
            tenantSettings: {
                domain: 'kurtrocks.com',
                teachers: [{ code: 'MAY', email: 'brian.may@kurtrocks.com' }],
                classes: [{ code: '1A', headEmail: 'brian.may@kurtrocks.com' }]
            },
            extraEmails: ['lehrer01.demo@ms365.schule']
        });
        expect(allowed.has('kurtsoeser.at')).toBe(true);
        expect(allowed.has('kurtrocks.com')).toBe(true);
        expect(allowed.has('ms365.schule')).toBe(true);
    });

    it('parseVerifiedEmailDomainsField', () => {
        expect(parseVerifiedEmailDomainsField('kurtrocks.com, kurtsoeser.at')).toEqual([
            'kurtrocks.com',
            'kurtsoeser.at'
        ]);
    });

    it('validate allows multi-domain owners in same tenant', () => {
        const r = validateKursteamBackendTenantContext({
            tenantId: 'tid-1',
            loginUpn: 'kurt@kurtsoeser.at',
            schoolDomainNoAt: 'ms365.schule',
            tenantSettings: {
                teachers: [
                    { email: 'brian.may@kurtrocks.com' },
                    { email: 'lehrer01.demo@ms365.schule' }
                ]
            },
            teams: [
                { besitzer: 'brian.may@kurtrocks.com' },
                { besitzer: 'lehrer01.demo@ms365.schule' }
            ]
        });
        expect(r.ok).toBe(true);
    });

    it('validate blocks unknown external domain', () => {
        const r = validateKursteamBackendTenantContext({
            tenantId: 'tid-1',
            loginUpn: 'kurt@kurtsoeser.at',
            schoolDomainNoAt: 'kurtsoeser.at',
            tenantSettings: { teachers: [] },
            teams: [{ besitzer: 'x@gmail.com' }]
        });
        expect(r.ok).toBe(false);
        expect(r.foreignDomains).toContain('gmail.com');
    });

    it('findForeignOwnerDomains', () => {
        const foreign = findForeignOwnerDomains(
            ['a@kurtrocks.com', 'b@kurtsoeser.at'],
            new Set(['kurtrocks.com'])
        );
        expect(foreign).toEqual(['kurtsoeser.at']);
    });
});
