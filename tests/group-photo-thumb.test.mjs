import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const projectRoot = join(dirname(fileURLToPath(import.meta.url)), '..');

function read(rel) {
    return readFileSync(join(projectRoot, rel), 'utf8');
}

describe('Gruppenbild-Thumbnails', () => {
    it('stellt gemeinsame Thumb-API bereit', () => {
        const js = read('src/shared/group-photo-thumb.js');
        expect(js).toContain('window.ms365GroupPhotoThumb');
        expect(js).toContain('createThumb');
        expect(js).toContain('hydrate');
        expect(js).toContain('invalidate');
        expect(js).toContain('IntersectionObserver');
    });

    it('synchronisiert Gruppenfotos optional mit Teams', () => {
        const js = read('src/shared/graph-unified-groups.js');
        expect(js).toContain('setTeamPhoto');
        expect(js).toContain('deleteTeamPhoto');
        expect(js).toContain('syncTeamPhotoForGroup');
        expect(js).toContain('TeamSettings.ReadWrite.All');
    });

    it('bindet Thumbnails in Listen-Tools ein', () => {
        for (const rel of [
            'src/tools/jahrgangsgruppen/jahrgangsgruppen.js',
            'src/tools/arge-fachgruppen/arge-fachgruppen.js',
            'src/tools/organisations-assistent/organisations-assistent-cohorts.js',
            'src/tools/schulstruktur-sync/schulstruktur-sync.js',
            'src/tools/schueler-lehrer-gruppen/slg-gruppenverwaltung.js'
        ]) {
            const js = read(rel);
            expect(js, rel).toContain('ms365GroupPhotoThumb');
            expect(js, rel).toContain('.hydrate');
        }
    });

    it('lädt group-photo-thumb.js auf Group-Detail-Seiten', () => {
        for (const html of [
            'tools/jahrgangsgruppen.html',
            'tools/arge-fachgruppen.html',
            'tools/gruppenerstellung.html'
        ]) {
            expect(read(html), html).toContain('src/shared/group-photo-thumb.js');
        }
    });

    it('stylet Listen-Thumbnails in group-detail.css', () => {
        const css = read('src/shared/group-detail/group-detail.css');
        expect(css).toContain('.gd-group-photo-thumb');
        expect(css).toContain('.slg-search-result-row');
    });
});
