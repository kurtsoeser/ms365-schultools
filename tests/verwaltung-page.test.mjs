import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

describe('Verwaltungs-Modul Seite', () => {
    it('hängt SLG-Live-Details und das Verwaltungs-Skript ein', () => {
        const html = readFileSync(join(projectRoot, 'tools/verwaltung.html'), 'utf8');
        expect(html).toContain('MS365-Schulverwaltung – Verwaltung');
        expect(html).toContain('src/shared/group-detail/group-detail.js');
        expect(html).toContain('src/tools/schueler-lehrer-gruppen/slg-live-details.js');
        expect(html).toContain('src/tools/verwaltung/verwaltung-gruppenverwaltung.js');
        expect(html).toContain('id="slgVerwaltungCount"');
        expect(html).toContain('data-slg-kind="verwaltung"');
        expect(html).toContain('id="vwRoleList"');
        expect(html).toContain('id="groupDetailHost"');
        expect(html).toContain('id="vwRolePanel"');
        expect(html).toContain('id="vwBtnAddRole"');
        expect(html).toContain('id="vwRolePeopleBody"');
    });
});
