import { describe, it, expect } from 'vitest';
import {
    getDemoSeedPackage,
    buildLocalDemoState,
    buildDemoAktivitaeten,
    DEMO_SEED_TAG,
    DEMO_SCHOOL_YEAR
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-demo-data.js';
import { parseDemoImportJson } from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-demo-seed.js';

describe('schulaktivitaeten-planer-demo', () => {
    it('builds 30–40 activities for 2026/27', () => {
        const rows = buildDemoAktivitaeten();
        expect(rows.length).toBeGreaterThanOrEqual(30);
        expect(rows.length).toBeLessThanOrEqual(40);
        const ids = new Set(rows.map((r) => r.AktivitaetId));
        expect(ids.size).toBe(rows.length);
        rows.forEach((r) => {
            expect(r.AktivitaetId).toMatch(/^akt-demo-/);
            expect(r.Notiz).toContain(DEMO_SEED_TAG);
            expect(r.Startdatum).toMatch(/^202[67]-/);
            expect(['Exkursion', 'Schulaktivitaet', 'Veranstaltung', 'Sonstiges']).toContain(r.Typ);
            expect(['beantragt', 'genehmigt', 'abgelehnt']).toContain(r.Status);
        });
    });

    it('getDemoSeedPackage has counts and stammdaten', () => {
        const pack = getDemoSeedPackage();
        expect(pack.schoolYear).toBe(DEMO_SCHOOL_YEAR);
        expect(pack.aktivitaeten.length).toBe(pack.counts.aktivitaeten);
        expect(pack.counts.beantragt).toBeGreaterThan(0);
        expect(pack.counts.genehmigt).toBeGreaterThan(0);
        expect(pack.stammdaten.classes.length).toBe(10);
        expect(pack.stammdaten.teachers.length).toBe(8);
    });

    it('buildLocalDemoState maps fields', () => {
        const local = buildLocalDemoState();
        expect(local.items.length).toBeGreaterThanOrEqual(30);
        expect(local.items[0].titel).toBeTruthy();
        expect(local.items[0].startdatum).toMatch(/^\d{4}-\d{2}-\d{2}$/);
        expect(local.rules.minVorlaufTage).toBe(7);
    });

    it('parseDemoImportJson accepts package', () => {
        const pack = parseDemoImportJson(getDemoSeedPackage());
        expect(pack.aktivitaeten.length).toBeGreaterThanOrEqual(30);
        expect(pack.seedTag).toBe(DEMO_SEED_TAG);
    });
});
