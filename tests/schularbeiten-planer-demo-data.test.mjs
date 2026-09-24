import { describe, it, expect } from 'vitest';
import { getDemoSeedPackage, DEMO_SITE_DEFAULT, buildDemoSchularbeiten } from '../src/tools/schularbeiten-planer/schularbeiten-planer-demo-data.js';

describe('schularbeiten demo 2026/27', () => {
    it('liefert umfassendes Paket', () => {
        const p = getDemoSeedPackage();
        expect(p.siteDefault).toContain('kurtrocks.sharepoint.com');
        expect(p.counts.schularbeiten).toBeGreaterThanOrEqual(50);
        expect(p.counts.terminfenster).toBeGreaterThanOrEqual(8);
        expect(p.counts.fachMeta).toBe(12);
        expect(p.stammdaten.classes.length).toBe(10);
        expect(p.regelwerk.Aktiv).toBe(true);
    });

    it('Schularbeiten haben stabile IDs und gemischte Status', () => {
        const rows = buildDemoSchularbeiten();
        const ids = new Set(rows.map((r) => r.SchularbeitId));
        expect(ids.size).toBe(rows.length);
        expect(rows.some((r) => r.Status === 'fixiert')).toBe(true);
        expect(rows.some((r) => r.Status === 'beantragt')).toBe(true);
        expect(rows.some((r) => r.Status === 'abgelehnt')).toBe(true);
        expect(rows.some((r) => r.Semester === 'WS')).toBe(true);
        expect(rows.some((r) => r.Semester === 'SS')).toBe(true);
    });

    it('keine SA in gesperrten Kernferien (Stichprobe)', () => {
        const rows = buildDemoSchularbeiten().filter((r) => r.Status !== 'abgelehnt');
        const inXmas = rows.filter((r) => r.Datum >= '2026-12-24' && r.Datum <= '2027-01-06');
        expect(inXmas).toHaveLength(0);
        expect(DEMO_SITE_DEFAULT).toMatch(/MS365-Schultools/);
    });
});
