import { describe, it, expect } from 'vitest';
import {
    buildFsFromFileList,
    listVirtualChildren,
    inventoryVirtual,
    folderSamplesFromSelection
} from '../src/tools/nextcloud-teams-migrate/nextcloud-teams-migrate-local.js';

function fakeFile(rel, size) {
    const name = rel.split('/').pop();
    const f = new File([new Uint8Array(size || 1)], name);
    Object.defineProperty(f, 'webkitRelativePath', { value: rel });
    return f;
}

describe('local virtual fs', () => {
    it('builds tree from folder files', () => {
        const vfs = buildFsFromFileList([
            fakeFile('Physik/01 Fachkonferenz/a.pdf', 10),
            fakeFile('Physik/Material/b.txt', 5)
        ]);
        const root = listVirtualChildren(vfs, '');
        expect(root.some((x) => x.name === 'Physik' && x.isFolder)).toBe(true);
        const physik = listVirtualChildren(vfs, 'Physik');
        expect(physik.map((x) => x.name).sort()).toEqual(['01 Fachkonferenz', 'Material']);
        const inv = inventoryVirtual(vfs, 'Physik');
        expect(inv.fileCount).toBe(2);
        expect(inv.totalBytes).toBe(15);
    });

    it('folder samples only selected folders, not children', () => {
        const vfs = buildFsFromFileList([
            fakeFile('1. Fachkonferenz/21_22/a.pdf', 1),
            fakeFile('1. Fachkonferenz/22_23/b.pdf', 1)
        ]);
        const samples = folderSamplesFromSelection(vfs, ['1. Fachkonferenz']);
        expect(samples).toEqual([
            { name: '1. Fachkonferenz', isFolder: true, relPath: '1. Fachkonferenz' }
        ]);
        expect(samples.some((s) => s.name === '21_22')).toBe(false);
    });
});
