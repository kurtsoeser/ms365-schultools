import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

describe('Jahrgangsgruppen-Modul Seite', () => {
    it('hängt SLG-Live-Details und das Listen-Skript ein', () => {
        const html = readFileSync(join(projectRoot, 'tools/jahrgangsgruppen.html'), 'utf8');
        expect(html).toContain('MS365-Schulverwaltung – Jahrgangsgruppen');
        expect(html).toContain('src/tools/schueler-lehrer-gruppen/slg-live-details.js');
        expect(html).toContain('src/tools/jahrgangsgruppen/jahrgangsgruppen.js');
        expect(html).toContain('id="slgLiveName"');
        expect(html).toContain('id="jgListItems"');
        expect(html).toContain('href="jahrgang.html"');
        expect(html).toContain('Wizard: alle neu erstellen');
    });

    it('lässt den Bulk-Wizard unter der alten URL erreichbar', () => {
        const wizard = readFileSync(join(projectRoot, 'tools/jahrgang.html'), 'utf8');
        expect(wizard).toContain('src/tools/jahrgang/jahrgang.js');
        expect(wizard).toContain('jahrgangsgruppen.html');
    });
});
