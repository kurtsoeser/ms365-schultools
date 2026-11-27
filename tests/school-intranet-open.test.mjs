import { describe, expect, it } from 'vitest';
import {
    intranetHubAudienceDecision,
    normalizeHttpsUrl,
    resolveSchoolIntranetOpenUrl
} from '../src/shared/school-intranet-open.js';

describe('school-intranet-open', () => {
    it('nimmt nur https-Adressen', () => {
        expect(normalizeHttpsUrl('https://kurtrocks.sharepoint.com/sites/intranet')).toBe(
            'https://kurtrocks.sharepoint.com/sites/intranet'
        );
        expect(normalizeHttpsUrl('http://kurtrocks.sharepoint.com/sites/intranet')).toBe('');
        expect(normalizeHttpsUrl('javascript:alert(1)')).toBe('');
        expect(normalizeHttpsUrl('')).toBe('');
    });

    it('bevorzugt die öffentliche Intranet-Seite vor der Listen-Site', () => {
        expect(
            resolveSchoolIntranetOpenUrl({
                schoolIntranetSiteUrl: 'https://kurtrocks.sharepoint.com/sites/intranet',
                intranetSiteUrl: 'https://kurtrocks.sharepoint.com/sites/listen'
            })
        ).toBe('https://kurtrocks.sharepoint.com/sites/intranet');
        expect(
            resolveSchoolIntranetOpenUrl({
                intranetSiteUrl: 'https://kurtrocks.sharepoint.com/sites/intranet'
            })
        ).toBe('https://kurtrocks.sharepoint.com/sites/intranet');
    });

    it('lässt die Schul-IT im Assistenten und schickt andere auf die SharePoint-Seite', () => {
        const setup = { schoolIntranetSiteUrl: 'https://kurtrocks.sharepoint.com/sites/intranet' };
        expect(intranetHubAudienceDecision({ loggedIn: true, isIt: true }, setup)).toEqual({
            action: 'stay'
        });
        expect(intranetHubAudienceDecision({ loggedIn: true, isIt: false }, setup)).toEqual({
            action: 'open',
            url: 'https://kurtrocks.sharepoint.com/sites/intranet'
        });
        expect(intranetHubAudienceDecision({ loggedIn: true, isIt: false }, {})).toEqual({
            action: 'missing'
        });
        expect(intranetHubAudienceDecision({ loggedIn: false }, setup)).toEqual({ action: 'wait' });
    });
});
