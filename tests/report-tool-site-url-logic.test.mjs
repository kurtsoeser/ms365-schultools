import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const {
    canReportToolSiteUrl,
    normalizeReportedToolSiteUrl,
    findLicenseItemForTenant
} = require('../backend/license-api/src/lib/evaluate-license.js');

describe('Tool-Site melden an Betreiber', () => {
    it('autorisiert designierte IT-Kontakte', () => {
        expect(
            canReportToolSiteUrl({
                allowed: true,
                contactEmails: ['it@schule.at'],
                upn: 'it@schule.at'
            }).ok
        ).toBe(true);
        expect(
            canReportToolSiteUrl({
                allowed: true,
                contactEmails: ['it@schule.at'],
                upn: 'lehrer@schule.at'
            }).ok
        ).toBe(false);
        expect(
            canReportToolSiteUrl({
                allowed: true,
                contactEmails: [],
                upn: 'ga@schule.at'
            }).reason
        ).toBe('no_contacts_bootstrap');
    });

    it('nimmt nur https SharePoint-Sites an', () => {
        expect(normalizeReportedToolSiteUrl('https://a.sharepoint.com/sites/MS365-Schultools/')).toEqual({
            ok: true,
            url: 'https://a.sharepoint.com/sites/MS365-Schultools'
        });
        expect(normalizeReportedToolSiteUrl('https://evil.example/app').ok).toBe(false);
        expect(normalizeReportedToolSiteUrl('http://a.sharepoint.com/sites/X').ok).toBe(false);
    });

    it('findet Lizenzzeile inkl. Item-ID', () => {
        const hit = findLicenseItemForTenant(
            [{ id: '9', fields: { TenantId: 'tid-1', Title: 'HAK' } }],
            'TID-1'
        );
        expect(hit && hit.id).toBe('9');
    });
});
