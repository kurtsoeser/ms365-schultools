import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import {
    normalizeSchoolToolSitesSetup,
    resolveSchoolToolSiteUrl,
    suggestPrimarySiteUrl
} from '../src/shared/school-tool-sites.js';
import { resolveSiteUrls } from '../src/tools/datenlandkarte/datenlandkarte-spo-metrics.js';

describe('school-tool-sites', () => {
    const prevWindow = globalThis.window;
    const store = new Map();

    beforeEach(() => {
        store.clear();
        globalThis.window = {
            ms365AppDataV2: {
                getSetup: () => ({
                    intranetSiteUrl: 'https://contoso.sharepoint.com/sites/intranet',
                    planerDefaultSiteUrl: 'https://contoso.sharepoint.com/sites/planer',
                    planerSiteUrls: { freistellung: 'https://contoso.sharepoint.com/sites/fr' },
                    schoolMailboxes: {}
                })
            }
        };
        globalThis.localStorage = {
            getItem: (k) => (store.has(k) ? store.get(k) : null),
            setItem: (k, v) => store.set(k, String(v)),
            removeItem: (k) => store.delete(k)
        };
    });

    afterEach(() => {
        globalThis.window = prevWindow;
    });

    it('normalizeSchoolToolSitesSetup füllt Defaults', () => {
        const n = normalizeSchoolToolSitesSetup({});
        expect(n.planerDefaultSiteUrl).toBe('');
        expect(n.planerSiteUrls.schularbeiten).toBe('');
        expect(n.schoolMailboxes.elternbriefName).toBe('Elternbrief Schule');
    });

    it('resolveSchoolToolSiteUrl: Override > Default > Intranet', () => {
        expect(resolveSchoolToolSiteUrl('freistellung')).toBe('https://contoso.sharepoint.com/sites/fr');
        expect(resolveSchoolToolSiteUrl('schularbeiten')).toBe('https://contoso.sharepoint.com/sites/planer');
        store.set('ms365-sa-site-url', 'https://contoso.sharepoint.com/sites/sa-local');
        expect(resolveSchoolToolSiteUrl('schularbeiten')).toBe('https://contoso.sharepoint.com/sites/sa-local');
        store.set('ms365-srdp-planer-site-v1', 'https://contoso.sharepoint.com/sites/srdp-local');
        expect(resolveSchoolToolSiteUrl('srdp')).toBe('https://contoso.sharepoint.com/sites/srdp-local');
        store.set(
            'ms365-krankmeldung-sp-settings-v1',
            JSON.stringify({ siteUrl: 'https://contoso.sharepoint.com/sites/km-local' })
        );
        expect(resolveSchoolToolSiteUrl('krankmeldung')).toBe(
            'https://contoso.sharepoint.com/sites/km-local'
        );
        expect(resolveSchoolToolSiteUrl('einsammler')).toBe('https://contoso.sharepoint.com/sites/planer');
        store.set(
            'ms365-einsammler-sp-settings-v1',
            JSON.stringify({ siteUrl: 'https://contoso.sharepoint.com/sites/einsammler-local' })
        );
        expect(resolveSchoolToolSiteUrl('einsammler')).toBe(
            'https://contoso.sharepoint.com/sites/einsammler-local'
        );
        expect(resolveSchoolToolSiteUrl('elternbrief')).toBe(
            'https://contoso.sharepoint.com/sites/planer'
        );
        store.set(
            'ms365-elternbrief-sp-settings-v1',
            JSON.stringify({ siteUrl: 'https://contoso.sharepoint.com/sites/elternbrief-local' })
        );
        expect(resolveSchoolToolSiteUrl('elternbrief')).toBe(
            'https://contoso.sharepoint.com/sites/elternbrief-local'
        );
        expect(resolveSchoolToolSiteUrl('technik')).toBe('https://contoso.sharepoint.com/sites/planer');
        store.set(
            'ms365-technik-meldungen-sp-settings-v1',
            JSON.stringify({ siteUrl: 'https://contoso.sharepoint.com/sites/technik-local' })
        );
        expect(resolveSchoolToolSiteUrl('technik')).toBe(
            'https://contoso.sharepoint.com/sites/technik-local'
        );
        expect(resolveSchoolToolSiteUrl('ausgaben')).toBe(
            'https://contoso.sharepoint.com/sites/planer'
        );
        store.set(
            'ms365-ausgaben-sp-settings-v1',
            JSON.stringify({ siteUrl: 'https://contoso.sharepoint.com/sites/ausgaben-local' })
        );
        expect(resolveSchoolToolSiteUrl('ausgaben')).toBe(
            'https://contoso.sharepoint.com/sites/ausgaben-local'
        );
        expect(resolveSchoolToolSiteUrl('raumplan')).toBe(
            'https://contoso.sharepoint.com/sites/planer'
        );
        store.set(
            'ms365-raumplan-sp-settings-v1',
            JSON.stringify({ siteUrl: 'https://contoso.sharepoint.com/sites/raumplan-local' })
        );
        expect(resolveSchoolToolSiteUrl('raumplan')).toBe(
            'https://contoso.sharepoint.com/sites/raumplan-local'
        );
        expect(resolveSchoolToolSiteUrl('lehrerFreistellung')).toBe(
            'https://contoso.sharepoint.com/sites/planer'
        );
        store.set('ms365-lfr-planer-site-v1', 'https://contoso.sharepoint.com/sites/lfr-local');
        expect(resolveSchoolToolSiteUrl('lehrerFreistellung')).toBe(
            'https://contoso.sharepoint.com/sites/lfr-local'
        );
    });

    it('suggestPrimarySiteUrl bevorzugt planerDefault', () => {
        expect(suggestPrimarySiteUrl()).toBe('https://contoso.sharepoint.com/sites/planer');
    });

    it('datenlandkarte resolveSiteUrls nutzt planerDefault', () => {
        const sites = resolveSiteUrls({
            intranetSiteUrl: 'https://contoso.sharepoint.com/sites/intranet',
            planerDefaultSiteUrl: 'https://contoso.sharepoint.com/sites/planer'
        });
        expect(sites.schularbeiten).toBe('https://contoso.sharepoint.com/sites/planer');
        expect(sites.intranet).toBe('https://contoso.sharepoint.com/sites/intranet');
    });
});
