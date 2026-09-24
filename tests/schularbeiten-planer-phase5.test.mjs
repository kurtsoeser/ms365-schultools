import { describe, it, expect } from 'vitest';
import { saMarker, buildSchulterminFields } from '../src/tools/schularbeiten-planer/schularbeiten-planer-sync.js';
import { buildEmbedSnippet, fachMetaForCode } from '../src/tools/schularbeiten-planer/schularbeiten-planer-state.js';
import {
    parseDemoImportJson,
    packToLocalPlanerState
} from '../src/tools/schularbeiten-planer/schularbeiten-planer-demo-seed.js';
import { getDemoSeedPackage } from '../src/tools/schularbeiten-planer/schularbeiten-planer-demo-data.js';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

describe('schularbeiten-planer-sync', () => {
    it('saMarker und Schultermine-Felder', () => {
        expect(saMarker('sa-abc')).toBe('[SA:sa-abc]');
        const fields = buildSchulterminFields(
            {
                schularbeitId: 'sa-abc',
                fachCode: 'D',
                klasseCode: '3AK',
                thema: 'Erörterung',
                datum: '2026-11-10',
                lehrerCode: 'BAU',
                dauerMinuten: 100
            },
            { fach: 'Deutsch', klasse: '3AK' }
        );
        expect(fields.Kategorie).toBe('Prüfung');
        expect(fields.AllDay).toBe(true);
        expect(fields.Info).toContain('[SA:sa-abc]');
        expect(fields.Title).toContain('Deutsch');
        expect(fields.SyncStatus).toBe('pending');
    });
});

describe('phase5 helpers', () => {
    it('buildEmbedSnippet enthält Link', () => {
        const s = buildEmbedSnippet('https://example.test/tools/schularbeiten-planer.html');
        expect(s).toContain('schularbeiten-planer.html');
        expect(s).toContain('iframe');
    });

    it('fachMetaForCode', () => {
        expect(fachMetaForCode([{ fachCode: 'D', farbe: '#111' }], 'D')).toMatchObject({ farbe: '#111' });
        expect(fachMetaForCode([], 'D')).toBe(null);
    });
});

describe('demo JSON import', () => {
    it('parst docs/demo-data JSON und mappt auf Planer-State', () => {
        const raw = readFileSync(resolve('docs/demo-data/schularbeiten-2026-27.json'), 'utf8');
        const pack = parseDemoImportJson(raw);
        expect(pack.schularbeiten.length).toBeGreaterThanOrEqual(50);
        const local = packToLocalPlanerState(pack);
        expect(local.items.length).toBe(pack.schularbeiten.length);
        expect(local.items[0].thema).toBeTruthy();
        expect(local.windows.length).toBe(pack.terminfenster.length);
        expect(local.rules.maxProTag).toBe(1);
        expect(local.fachMeta.length).toBe(pack.fachMeta.length);
    });

    it('akzeptiert getDemoSeedPackage()', () => {
        const local = packToLocalPlanerState(getDemoSeedPackage());
        expect(local.items.some((i) => i.status === 'fixiert')).toBe(true);
    });
});
