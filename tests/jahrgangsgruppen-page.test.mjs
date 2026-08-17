import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

describe('Jahrgangsgruppen-Modul Seite', () => {
    it('hängt die zentrale Gruppen-Detailansicht und das Listen-Skript ein', () => {
        const html = readFileSync(join(projectRoot, 'tools/jahrgangsgruppen.html'), 'utf8');
        expect(html).toContain('MS365-Schulverwaltung – Klassengruppen');
        expect(html).toContain('src/shared/group-detail/group-detail.css');
        expect(html).toContain('src/shared/group-detail/group-detail.js');
        expect(html).toContain('src/tools/schueler-lehrer-gruppen/slg-live-details.js');
        expect(html).toContain('src/tools/jahrgangsgruppen/jahrgangsgruppen.js');
        expect(html).toContain('id="groupDetailHost"');
        expect(html).not.toContain('id="slgLiveName"');
        expect(html).toContain('id="jgListItems"');
        expect(html).toContain('href="jahrgang.html"');
        expect(html).toContain('id="jgBtnSmtpAll"');
        expect(html).toContain('id="jgBtnBulkOwner"');
        expect(html).toContain('id="jgBtnBulkDelete"');
        expect(html).toContain('id="jgBtnSelectMatched"');
        expect(html).toContain('id="jgBulkOwnerApply"');
        expect(html).toContain('id="jgBtnAddClass"');
        expect(html).toContain('id="jgBtnEditClass"');
        expect(html).toContain('id="jgBtnDeleteClass"');
        expect(html).toContain('id="jgClassModal"');
        expect(html).toContain('Sammelaktionen');
        expect(html).toContain('id="jgPageInfo"');
        const js = readFileSync(join(projectRoot, 'src/tools/jahrgangsgruppen/jahrgangsgruppen.js'), 'utf8');
        expect(js).toContain('ms365GroupDetail');
        expect(js).toContain("mount('#groupDetailHost'");
        expect(js).toContain('runBulkSetOwner');
        expect(js).toContain('runBulkDelete');
        expect(js).toContain('deleteUnifiedGroup');
        const graph = readFileSync(join(projectRoot, 'src/shared/graph-unified-groups.js'), 'utf8');
        expect(graph).toContain('async function deleteUnifiedGroup');
    });

    it('lässt den Bulk-Wizard unter der alten URL erreichbar', () => {
        const wizard = readFileSync(join(projectRoot, 'tools/jahrgang.html'), 'utf8');
        expect(wizard).toContain('src/tools/jahrgang/jahrgang.js');
        expect(wizard).toContain('jahrgangsgruppen.html');
        expect(wizard).toContain('einmalige Neuanlage');
        expect(wizard).toContain('Nicht der Alltagsweg');
    });
});
