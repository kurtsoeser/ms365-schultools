import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

function loadSis() {
    const sandbox = { console };
    sandbox.window = sandbox;
    createContext(sandbox);
    const full = join(projectRoot, 'src/shared/school-sis-import.js');
    runInContext(readFileSync(full, 'utf8'), sandbox, { filename: full });
    return sandbox;
}

describe('school-sis-import', () => {
    let ctx;

    beforeEach(() => {
        ctx = loadSis();
    });

    it('MS365-Vorlage mappt Elternspalten', () => {
        const rows = [
            {
                Klasse: '1A',
                Name: 'Anna Beispiel',
                'E-Mail': 'anna@schule.at',
                Eltern1: 'Maria',
                Eltern1Mail: 'maria@mail.com',
                Eltern2: 'Tom',
                Eltern2Mail: 'tom@mail.com'
            }
        ];
        const r = ctx.ms365SchoolSisImport.importStudentsAndGuardians({ objectRows: rows, source: 'ms365' });
        expect(r.meta.studentCount).toBe(1);
        expect(r.meta.withParents).toBe(1);
        expect(r.records[0].parentPairs).toHaveLength(2);
        expect(r.lines).toContain('maria@mail.com');
    });

    it('Sokrates-AOA aggregiert mehrere Elternzeilen pro Schülerkennzahl', () => {
        const aoa = [
            [
                'Klasse',
                'Schülerkennzahl',
                'Familienname',
                'Vorname',
                'Titel',
                'Akad. Grad',
                'Vorname',
                'Familienname',
                'Mailadresse'
            ],
            ['1A', '10001', 'Beispiel', 'Anna', '', '', 'Maria', 'Beispiel', 'maria@mail.com'],
            ['1A', '10001', 'Beispiel', 'Anna', '', '', 'Thomas', 'Beispiel', 'thomas@mail.com'],
            ['1B', '10002', 'Grohl', 'Dave', '', '', 'Jane', 'Grohl', 'jane@mail.com']
        ];
        const r = ctx.ms365SchoolSisImport.importStudentsAndGuardians({ aoa, source: 'sokrates' });
        expect(r.meta.studentCount).toBe(2);
        const anna = r.records.find((x) => x.externalId === '10001');
        expect(anna.parentPairs).toHaveLength(2);
        expect(anna.name).toBe('Anna Beispiel');
        expect(anna.klasse).toBe('1A');
    });

    it('erkennt Sokrates an Headern', () => {
        expect(
            ctx.ms365SchoolSisImport.detectSourceFromHeaders(['Klasse', 'Schülerkennzahl', 'Familienname', 'Mailadresse'])
        ).toBe('sokrates');
    });

    it('liefert CSV-Vorlage mit Elternspalten', () => {
        const aoa = ctx.ms365SchoolSisImport.ms365TemplateAoa();
        expect(aoa[0]).toContain('Eltern1Mail');
        expect(aoa.length).toBeGreaterThan(2);
    });

    it('diffSisImport wertet geteilte Elternmails bei Geschwistern nicht als Konflikt', () => {
        const sis = ctx.ms365SchoolSisImport;
        const incoming = [
            {
                klasse: '1A',
                name: 'Anna',
                email: 'anna@schule.at',
                parentPairs: [{ name: 'Maria', email: 'eltern@mail.com' }]
            },
            {
                klasse: '3B',
                name: 'Ben',
                email: 'ben@schule.at',
                parentPairs: [{ name: 'Maria', email: 'eltern@mail.com' }]
            }
        ];
        const diff = sis.diffSisImport([], incoming);
        expect(diff.conflicts).toHaveLength(0);
        expect(diff.counts.added).toBe(2);
    });

    it('diffSisImport erkennt neu, geändert und Konflikte', () => {
        const sis = ctx.ms365SchoolSisImport;
        const existing = [
            { klasse: '1A', name: 'Anna', email: 'anna@schule.at', externalId: '10001' },
            { klasse: '1B', name: 'Ben', email: 'ben@schule.at' }
        ];
        const incoming = [
            { klasse: '2A', name: 'Anna', email: 'anna@schule.at', externalId: '10001' },
            { klasse: '1C', name: 'Clara', email: 'clara@schule.at' },
            { klasse: '1A', name: 'Anna neu', email: 'anna@schule.at', externalId: '99999' }
        ];
        const diff = sis.diffSisImport(existing, incoming);
        expect(diff.counts.added).toBe(1);
        expect(diff.counts.updated).toBeGreaterThanOrEqual(1);
        expect(diff.counts.removed).toBe(1);
        expect(diff.conflicts.some((c) => c.type === 'email-id')).toBe(true);
        const merged = sis.applySisImport(existing, incoming, { mode: 'merge' });
        expect(merged.some((r) => r.email === 'ben@schule.at')).toBe(true);
        expect(merged.some((r) => r.email === 'clara@schule.at')).toBe(true);
        const replaced = sis.applySisImport(existing, incoming, { mode: 'replace' });
        expect(replaced.some((r) => r.email === 'ben@schule.at')).toBe(false);
        expect(sis.summarizeSisDiff(diff)).toMatch(/neu/);
    });

    it('diffSisImport markiert E-Mail- und Elternänderungen getrennt', () => {
        const sis = ctx.ms365SchoolSisImport;
        const existing = [
            {
                klasse: '1A',
                name: 'Anna',
                email: 'anna@schule.at',
                externalId: '10001',
                parentPairs: [{ name: 'Maria', email: 'alt@mail.com' }]
            }
        ];
        const incoming = [
            {
                klasse: '1A',
                name: 'Anna',
                email: 'anna.neu@schule.at',
                externalId: '10001',
                parentPairs: [{ name: 'Maria', email: 'neu@mail.com' }]
            }
        ];
        const diff = sis.diffSisImport(existing, incoming);
        expect(diff.updated).toHaveLength(1);
        expect(diff.updated[0].emailChanged).toBe(true);
        expect(diff.updated[0].parentsChanged).toBe(true);
        expect(diff.updated[0].klasseChanged).toBe(false);
    });
});
