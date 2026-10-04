import { describe, it, expect } from 'vitest';
import { resolveSiteUrls, SPO_LIST_PROBES } from '../src/tools/datenlandkarte/datenlandkarte-spo-metrics.js';
import { mergeDatenMetrics, collectDatenMetrics } from '../src/tools/datenlandkarte/datenlandkarte-metrics.js';

describe('datenlandkarte-spo-metrics', () => {
    it('resolveSiteUrls nutzt Setup und Fallback auf Intranet', () => {
        const sites = resolveSiteUrls({ intranetSiteUrl: 'https://contoso.sharepoint.com/sites/intranet' });
        expect(sites.intranet).toBe('https://contoso.sharepoint.com/sites/intranet');
        expect(sites.schularbeiten).toBe('https://contoso.sharepoint.com/sites/intranet');
        expect(sites.freistellung).toBe('https://contoso.sharepoint.com/sites/intranet');
    });

    it('mergeDatenMetrics überschreibt SP-Keys', () => {
        const local = collectDatenMetrics();
        const merged = mergeDatenMetrics(local, { spoListKlassen: { value: 42, hint: 'SharePoint · Klassen' } });
        expect(merged.spoListKlassen.value).toBe(42);
        expect(merged.classes).toEqual(local.classes);
    });

    it('SPO_LIST_PROBES deckt alle spoList countKeys im Katalog ab', () => {
        const probeKeys = new Set(SPO_LIST_PROBES.map((p) => p.countKey));
        expect(probeKeys.has('spoListKlassen')).toBe(true);
        expect(probeKeys.has('spoListPwAktionen')).toBe(true);
    });
});
