import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

function read(rel) {
    return readFileSync(join(projectRoot, rel), 'utf8');
}

describe('user-photo-thumb', () => {
    it('exponiert Cache/Hydrate-API und nutzt Graph-Foto', () => {
        const js = read('src/shared/user-photo-thumb.js');
        expect(js).toContain('ms365UserPhotoThumb');
        expect(js).toContain('data-ms365-user-photo');
        expect(js).toContain('fetchUserPhotoBlob');
        expect(js).toContain('hydrate');
    });

    it('Graph stellt fetchUserPhotoBlob bereit', () => {
        const js = read('src/shared/graph-unified-groups.js');
        expect(js).toContain('fetchUserPhotoBlob');
        expect(js).toContain('/photos/');
    });

    it('Organigramm-Seiten laden User- und Gruppenfotos', () => {
        for (const html of [
            'tools/schul-organigramm.html',
            'tools/verwaltung.html',
            'tenant.html'
        ]) {
            const src = read(html);
            expect(src, html).toContain('user-photo-thumb.js');
            expect(src, html).toContain('group-photo-thumb.js');
        }
    });
});
