import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

function loadRenamePreview() {
    const full = join(projectRoot, 'src/shared/graph-rename-preview.js');
    const code = readFileSync(full, 'utf8');
    const sandbox = { console };
    sandbox.window = sandbox;
    createContext(sandbox);
    runInContext(code, sandbox, { filename: full });
    return sandbox;
}

describe('graph-rename-preview', () => {
    let ctx;

    beforeEach(() => {
        ctx = loadRenamePreview();
    });

    it('computeNewDisplayNamePlusOne increments stage', () => {
        const fn = ctx.ms365GraphRenamePreview.computeNewDisplayNamePlusOne;
        expect(fn('Klasse 1A', 'Klasse')).toBe('Klasse 2A');
        expect(fn('Klasse 10HAK', 'Klasse')).toBe('Klasse 11HAK');
        expect(fn('1A', 'Klasse')).toBe(null);
    });

    it('incrementLeadingGrade erhöht 1HMA und 10HAK', () => {
        const inc = ctx.ms365GraphRenamePreview.incrementLeadingGrade;
        expect(inc('1HMA')).toBe('2HMA');
        expect(inc('10HAK')).toBe('11HAK');
        expect(inc('Klasse 1A')).toBe(null);
    });

    it('suggestDisplayNamePlusOne nutzt Präfix oder führende Stufe', () => {
        const sug = ctx.ms365GraphRenamePreview.suggestDisplayNamePlusOne;
        expect(sug('Klasse 1A', 'Klasse')).toBe('Klasse 2A');
        expect(sug('1HMA', 'Klasse')).toBe('2HMA');
        expect(sug('Projektteam', 'Klasse')).toBe(null);
    });
});
