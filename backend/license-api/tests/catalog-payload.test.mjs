import { describe, it, expect } from 'vitest';
import { buildStoredCatalog, emptyCatalog } from '../src/lib/catalog-payload.js';
import { encodeDrivePath } from '../src/lib/sharepoint-catalog.js';

describe('catalog-payload', () => {
    it('speichert nur erlaubte Felder und lehnt doppelte IDs ab', () => {
        const stored = buildStoredCatalog(
            {
                kind: 'egal',
                schoolForms: ['HAKB', 'hakb', 'AHS'],
                templates: [
                    {
                        id: 'tpl-1',
                        name: 'MAM',
                        schoolForm: 'HAKB',
                        subjectCode: 'MAM',
                        schulstufe: '10',
                        semester: 'WS',
                        description: 'Grundlagen',
                        channels: [{ id: 'c1', displayName: '01 - Terme' }, { displayName: '' }],
                        secret: 'nein'
                    }
                ]
            },
            { strict: true, updatedBy: 'kurt@kurtsoeser.at', updatedAt: '2026-09-22T00:00:00.000Z' }
        );
        expect(stored.kind).toBe('ms365-kursteam-templates');
        expect(stored.version).toBe(3);
        expect(stored.updatedBy).toBe('kurt@kurtsoeser.at');
        expect(stored.schoolForms).toEqual(['HAKB', 'AHS']);
        expect(stored.templates).toHaveLength(1);
        expect(stored.templates[0].secret).toBeUndefined();
        expect(stored.templates[0].channels).toEqual([{ id: 'c1', displayName: '01 - Terme' }]);

        expect(() =>
            buildStoredCatalog(
                {
                    templates: [
                        { id: 'a', name: 'Eins', channels: [] },
                        { id: 'a', name: 'Zwei', channels: [] }
                    ]
                },
                { strict: true }
            )
        ).toThrow(/Doppelte Vorlagen-ID/);
    });

    it('liest unvollständige Dateien tolerant und liefert einen leeren Katalog', () => {
        const loose = buildStoredCatalog(
            { templates: [{ name: '' }, { id: 'ok', name: 'Bleibt', channels: [] }] },
            { strict: false, updatedAt: '2026-01-01T00:00:00.000Z' }
        );
        expect(loose.templates.map((t) => t.id)).toEqual(['ok']);
        const empty = emptyCatalog(
            {
                catalogLibraryName: 'MS365-Katalog',
                catalogKursteamPath: 'vorlagen/kursteam-kanaele.json',
                siteWebUrl: 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools'
            },
            { missing: true, message: 'noch leer' }
        );
        expect(empty.missing).toBe(true);
        expect(empty.templates).toEqual([]);
        expect(empty.library).toBe('MS365-Katalog');
    });

    it('kodiert den Dateipfad je Segment', () => {
        expect(encodeDrivePath('vorlagen/kursteam-kanaele.json')).toBe('vorlagen/kursteam-kanaele.json');
        expect(encodeDrivePath('/vorlagen/mein blatt.json')).toBe('vorlagen/mein%20blatt.json');
    });
});
