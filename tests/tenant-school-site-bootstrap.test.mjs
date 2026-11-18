import { describe, it, expect } from 'vitest';
import {
    normalizeToolSiteUrl,
    toolSiteUrlFromLicenseResult
} from '../src/shared/tenant-school-site-bootstrap.js';

describe('tenant-school-site-bootstrap', () => {
    it('normalisiert SharePoint-URLs', () => {
        expect(normalizeToolSiteUrl('https://schule.sharepoint.com/sites/MS365-Schultools/')).toBe(
            'https://schule.sharepoint.com/sites/MS365-Schultools'
        );
        expect(normalizeToolSiteUrl('')).toBe('');
        expect(normalizeToolSiteUrl('not-a-url')).toBe('');
    });

    it('liest toolSiteUrl aus license/me-Form', () => {
        expect(
            toolSiteUrlFromLicenseResult({
                allowed: true,
                sites: { toolSiteUrl: 'https://a.sharepoint.com/sites/X/' },
                license: { toolSiteUrl: 'https://b.sharepoint.com/sites/Y' }
            })
        ).toBe('https://a.sharepoint.com/sites/X');

        expect(
            toolSiteUrlFromLicenseResult({
                allowed: true,
                license: { toolSiteUrl: 'https://b.sharepoint.com/sites/Y' }
            })
        ).toBe('https://b.sharepoint.com/sites/Y');
    });
});
