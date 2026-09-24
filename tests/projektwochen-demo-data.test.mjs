import { describe, it, expect } from 'vitest';
import {
    getDemoSeedPackage,
    buildLocalDemoState,
    DEMO_SITE_DEFAULT,
    DEMO_SEED_TAG,
    buildDemoAngebotFields
} from '../src/tools/projektwochen/projektwochen-demo-data.js';
import { parseDemoImportJson } from '../src/tools/projektwochen/projektwochen-demo-seed.js';
import {
    resolveWeekdayFieldName,
    adaptAngebotFieldsForList
} from '../src/tools/projektwochen/projektwochen-graph.js';

describe('projektwochen demo kurtrocks', () => {
    it('liefert umfassendes Paket mit kurtrocks-Site', () => {
        const p = getDemoSeedPackage();
        expect(p.siteDefault).toContain('kurtrocks.sharepoint.com');
        expect(p.seedTag).toBe(DEMO_SEED_TAG);
        expect(p.counts.angebote).toBeGreaterThanOrEqual(18);
        expect(p.counts.freigegeben).toBeGreaterThanOrEqual(8);
        expect(p.counts.beantragt).toBeGreaterThanOrEqual(3);
        expect(p.stammdaten.teachers.length).toBe(8);
        expect(p.stammdaten.classes.length).toBe(10);
        expect(p.aktion.AktionId).toBeTruthy();
        expect(DEMO_SITE_DEFAULT).toMatch(/MS365-Schultools/);
    });

    it('Angebote haben stabile IDs und gemischte Status', () => {
        const rows = buildDemoAngebotFields();
        const ids = new Set(rows.map((r) => r.AngebotId));
        expect(ids.size).toBe(rows.length);
        expect(rows.some((r) => r.Status === 'freigegeben')).toBe(true);
        expect(rows.some((r) => r.Status === 'beantragt')).toBe(true);
        expect(rows.some((r) => r.Status === 'abgelehnt')).toBe(true);
        expect(rows.some((r) => r.Status === 'entwurf')).toBe(true);
        expect(rows.every((r) => String(r.NotizIntern || '').includes(DEMO_SEED_TAG) || r.NotizIntern)).toBe(true);
    });

    it('buildLocalDemoState und parseDemoImportJson', () => {
        const pack = getDemoSeedPackage();
        const local = buildLocalDemoState(pack);
        expect(local.aktionen).toHaveLength(1);
        expect(local.angebote.length).toBe(pack.angebote.length);
        expect(local.angebote[0].angebotId).toBeTruthy();
        const parsed = parseDemoImportJson(JSON.stringify(pack));
        expect(parsed.counts.angebote).toBe(pack.angebote.length);
    });

    it('Legacy-Spalte Tag statt Wochentag beim Schreiben', () => {
        expect(resolveWeekdayFieldName(['Tag', 'Title'])).toBe('Tag');
        expect(resolveWeekdayFieldName(['Wochentag', 'Tag'])).toBe('Wochentag');
        const adapted = adaptAngebotFieldsForList({ Title: 'Zoo', Wochentag: 'Mo', AngebotId: 'x' }, 'Tag');
        expect(adapted.Tag).toBe('Mo');
        expect(adapted.Wochentag).toBeUndefined();
    });
});
