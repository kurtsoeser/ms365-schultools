import { describe, it, expect } from 'vitest';
import {
    resolveSiteUrls,
    SPO_LIST_PROBES,
    resolveSpoProbeListTitle,
    findGraphListOnSite,
    fetchListItemCount
} from '../src/tools/datenlandkarte/datenlandkarte-spo-metrics.js';
import { mergeDatenMetrics, collectDatenMetrics } from '../src/tools/datenlandkarte/datenlandkarte-metrics.js';

describe('datenlandkarte-spo-metrics', () => {
    it('resolveSiteUrls nutzt Setup und Fallback auf Intranet', () => {
        const sites = resolveSiteUrls({ intranetSiteUrl: 'https://contoso.sharepoint.com/sites/intranet' });
        expect(sites.intranet).toBe('https://contoso.sharepoint.com/sites/intranet');
        expect(sites.schularbeiten).toBe('https://contoso.sharepoint.com/sites/intranet');
        expect(sites.freistellung).toBe('https://contoso.sharepoint.com/sites/intranet');
    });

    it('resolveSpoProbeListTitle nutzt intranetListTitles aus dem Setup', () => {
        const probe = SPO_LIST_PROBES.find((p) => p.countKey === 'spoListKlassen');
        expect(probe).toBeTruthy();
        const title = resolveSpoProbeListTitle(probe, {
            intranetListTitles: { klassen: 'Meine Klassenliste' }
        });
        expect(title).toBe('Meine Klassenliste');
        expect(resolveSpoProbeListTitle(probe, {})).toBe('Klassen');
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

    it('findGraphListOnSite matcht Listen per Enumeration (case-insensitive)', async () => {
        const calls = [];
        const fakeG = {
            graphPathSite: (id) => '/sites/' + id,
            graphJson: async (method, path) => {
                calls.push(path);
                if (path.indexOf('$filter=') >= 0) return { value: [] };
                if (path.indexOf('/lists?') >= 0) {
                    return {
                        value: [{ id: 'list-guid-1', displayName: 'Fächer', webUrl: 'https://x' }]
                    };
                }
                return { value: [] };
            }
        };
        const prev = globalThis.window;
        globalThis.window = { ms365SpoGraph: fakeG };
        try {
            const row = await findGraphListOnSite('tok', 'site-id', 'Fächer');
            expect(row && row.id).toBe('list-guid-1');
            expect(calls.some((p) => p.indexOf('/lists?') >= 0 && p.indexOf('$filter') < 0)).toBe(true);
        } finally {
            globalThis.window = prev;
        }
    });

    it('fetchListItemCount nutzt @odata.count als Fallback', async () => {
        const prev = globalThis.window;
        const prevFetch = globalThis.fetch;
        globalThis.window = {
            ms365SpoGraph: {
                graphPathSite: (id) => '/sites/' + id
            }
        };
        let n = 0;
        globalThis.fetch = async () => {
            n += 1;
            if (n === 1) {
                return { ok: false, status: 403, text: async () => 'denied' };
            }
            return {
                ok: true,
                status: 200,
                text: async () => JSON.stringify({ '@odata.count': 17, value: [] })
            };
        };
        try {
            const r = await fetchListItemCount('tok', 'site', 'list-id');
            expect(r.ok).toBe(true);
            expect(r.count).toBe(17);
            expect(r.via).toBe('@odata.count');
        } finally {
            globalThis.fetch = prevFetch;
            globalThis.window = prev;
        }
    });
});
