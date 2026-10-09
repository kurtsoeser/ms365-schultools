import { describe, it, expect } from 'vitest';
import {
    buildDemoFreistellungen,
    getDemoSeedPackage,
    extractDemoId,
    isDemoFreistellungFields,
    buildLocalDemoItems,
    itemsFromDemoPack,
    DEMO_SEED_TAG,
    DEMO_SCHOOL_YEAR
} from '../src/tools/freistellung-planer/freistellung-planer-demo.js';

describe('freistellung-planer-demo', () => {
    it('liefert 20–30 Freistellungen für SJ 2026/27', () => {
        const rows = buildDemoFreistellungen();
        expect(rows.length).toBeGreaterThanOrEqual(20);
        expect(rows.length).toBeLessThanOrEqual(30);
        expect(DEMO_SCHOOL_YEAR).toBe('2026/27');
        rows.forEach((r) => {
            expect(r.Beschreibung).toContain(DEMO_SEED_TAG);
            expect(extractDemoId(r.Beschreibung)).toMatch(/^fr-demo-\d+$/);
            expect(['Ausstehend', 'Genehmigt', 'Abgelehnt']).toContain(r.Status);
        });
    });

    it('Paket enthält Stammdaten und Counts', () => {
        const pack = getDemoSeedPackage();
        expect(pack.freistellungen.length).toBe(26);
        expect(pack.counts.gesamt).toBe(26);
        expect(pack.stammdaten.classes.length).toBeGreaterThan(5);
        expect(pack.counts.ausstehend + pack.counts.genehmigt + pack.counts.abgelehnt).toBe(26);
    });

    it('itemsFromDemoPack formatiert Datumsfelder', () => {
        const items = itemsFromDemoPack({
            freistellungen: [
                {
                    Title: 'Test (3A)',
                    Beginn: '2026-10-15',
                    Ende: '2026-10-15',
                    Status: 'Ausstehend',
                    Klasse: '3A',
                    GenehmigtAmKV: '2026-09-08'
                }
            ]
        });
        expect(items[0].beginn).toBe('2026-10-15T00:00');
        expect(items[0].genehmigtAmKv).toBe('2026-09-08');
    });

    it('erkennt Demo-Felder und baut lokale UI-Items', () => {
        const row = buildDemoFreistellungen()[0];
        expect(isDemoFreistellungFields(row)).toBe(true);
        const items = buildLocalDemoItems({});
        expect(items.length).toBe(26);
        expect(items[0].dayCount).toBeGreaterThanOrEqual(1);
    });
});
