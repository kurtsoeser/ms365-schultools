import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { getDemoSeedPackage, DEMO_SITE_DEFAULT, buildDemoSchularbeiten } from '../src/tools/schularbeiten-planer/schularbeiten-planer-demo-data.js';
import { mapSchularbeitToFields } from '../src/tools/schularbeiten-planer/schularbeiten-planer-graph.js';
import { parseDemoImportJson } from '../src/tools/schularbeiten-planer/schularbeiten-planer-demo-seed.js';

const hakJsonPath = join(dirname(fileURLToPath(import.meta.url)), '../docs/demo-data/schularbeiten-hak-steyr-2026-27.json');

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

    it('HAK-Import: leeres FixiertAm wird nicht an Graph gesendet', () => {
        const pack = parseDemoImportJson(readFileSync(hakJsonPath, 'utf8'));
        const row = pack.schularbeiten[0];
        expect(row.FixiertAm).toBe('');
        const fields = mapSchularbeitToFields(
            {
                titel: row.Titel || row.Title,
                schularbeitId: row.SchularbeitId,
                fachCode: row.FachCode,
                klasseCode: row.KlasseCode,
                datum: row.Datum,
                status: row.Status,
                fixiertAm: row.FixiertAm || undefined
            },
            { labels: { fach: { MAM: 'Mathematik' }, klasse: { '3AK': '3AK' } } }
        );
        expect(fields.FixiertAm).toBeUndefined();
        expect(fields.Titel).toBeTruthy();
        expect(fields.Title).toMatch(/^Schularbeit - /);
        expect(fields.Datum).toBe('2026-11-11');
        expect(String(fields.SchularbeitId).length).toBeLessThanOrEqual(40);
    });

    it('keine SA in gesperrten Kernferien (Stichprobe)', () => {
        const rows = buildDemoSchularbeiten().filter((r) => r.Status !== 'abgelehnt');
        const inXmas = rows.filter((r) => r.Datum >= '2026-12-24' && r.Datum <= '2027-01-06');
        expect(inXmas).toHaveLength(0);
        expect(DEMO_SITE_DEFAULT).toMatch(/MS365-Schultools/);
    });
});
