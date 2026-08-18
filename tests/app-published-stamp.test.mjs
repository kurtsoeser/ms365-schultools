import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createAppBuildInfo, writeAppBuildInfo } from '../scripts/app-build-info.mjs';
import {
    STAMP_ELEMENT_ID,
    createPublishedStampElement,
    formatPublishedStamp,
    parsePublishedAt
} from '../src/shared/app-published-stamp-logic.js';

const projectRoot = join(dirname(fileURLToPath(import.meta.url)), '..');

function read(rel) {
    return readFileSync(join(projectRoot, rel), 'utf8');
}

describe('Veröffentlichungsstempel', () => {
    it('liest publishedAt aus der Build-Info', () => {
        expect(parsePublishedAt({ publishedAt: ' 2026-08-18T10:31:00.000Z ' })).toBe(
            '2026-08-18T10:31:00.000Z'
        );
        expect(parsePublishedAt(null)).toBe('');
        expect(parsePublishedAt({})).toBe('');
    });

    it('formatiert Datum und Uhrzeit auf Wiener Zeit', () => {
        const label = formatPublishedStamp('2026-08-18T10:31:00.000Z');
        expect(label).toMatch(/^Stand: 18\.08\.2026/);
        expect(label).toMatch(/12:31/);
        expect(formatPublishedStamp('')).toBe('');
        expect(formatPublishedStamp('kein-datum')).toBe('');
    });

    it('erzeugt ein kleines Fußzeilen-Element', () => {
        const el = createPublishedStampElement('Stand: 18.08.2026, 12:31', {
            createElement(tag) {
                const node = {
                    tagName: String(tag).toUpperCase(),
                    id: '',
                    className: '',
                    textContent: '',
                    title: ''
                };
                return node;
            }
        });
        expect(el.id).toBe(STAMP_ELEMENT_ID);
        expect(el.className).toBe('app-published-stamp');
        expect(el.textContent).toBe('Stand: 18.08.2026, 12:31');
        expect(createPublishedStampElement('  ', { createElement() { return {}; } })).toBe(null);
    });

    it('schreibt app-build.json in den Dist-Ordner', async () => {
        const dir = await mkdtemp(join(tmpdir(), 'ms365-build-info-'));
        try {
            const info = createAppBuildInfo(new Date('2026-08-18T10:31:00.000Z'));
            const dest = await writeAppBuildInfo(dir, info);
            const raw = await readFile(dest, 'utf8');
            expect(JSON.parse(raw)).toEqual({ publishedAt: '2026-08-18T10:31:00.000Z' });
        } finally {
            await rm(dir, { recursive: true, force: true });
        }
    });

    it('wird über die PIN-Schranke auf allen Seiten eingebunden', () => {
        const pinGate = read('src/shared/pin-gate.js');
        expect(pinGate).toContain('app-published-stamp.js');
        expect(pinGate).toContain('data-ms365-published-stamp');
        expect(read('scripts/copy-static.mjs')).toContain('writeAppBuildInfo');
        expect(read('app.css')).toContain('.app-published-stamp');
    });
});
