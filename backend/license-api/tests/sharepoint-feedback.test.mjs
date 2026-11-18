import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const {
    mapFeedbackItem,
    parseAnhangUrls,
    normalizeCategory,
    normalizeStatus,
    normalizeAttachments,
    feedbackColumnDefs,
    CATEGORIES,
    STATUSES
} = require('../src/lib/sharepoint-feedback.js');

describe('sharepoint-feedback helpers', () => {
    it('hat Kern-Spalten und Status-/Kategorie-Werte', () => {
        const names = feedbackColumnDefs().map((c) => c.name);
        expect(names).toContain('Beschreibung');
        expect(names).toContain('MelderEmail');
        expect(names).toContain('AnhangUrls');
        expect(CATEGORIES).toContain('Bug');
        expect(STATUSES).toContain('Neu');
    });

    it('normalisiert Kategorie und Status', () => {
        expect(normalizeCategory('bug')).toBe('Bug');
        expect(normalizeCategory('xyz')).toBe('Feedback');
        expect(normalizeStatus('in bearbeitung')).toBe('In Bearbeitung');
        expect(normalizeStatus('')).toBe('Neu');
    });

    it('parst Anhang-URLs aus JSON und Zeilen', () => {
        expect(parseAnhangUrls('[{"name":"a.png","url":"https://x/a.png"}]')).toEqual([
            { name: 'a.png', url: 'https://x/a.png', path: undefined }
        ]);
        expect(parseAnhangUrls('https://x/a.png\nhttps://x/b.png')).toEqual([
            { name: 'Anhang', url: 'https://x/a.png' },
            { name: 'Anhang', url: 'https://x/b.png' }
        ]);
    });

    it('mapped Listeneintrag', () => {
        const mapped = mapFeedbackItem({
            id: '42',
            fields: {
                Title: 'Login hängt',
                Beschreibung: 'Nach PIN nichts',
                Kategorie: 'Bug',
                Status: 'Neu',
                MelderEmail: 'a@schule.at',
                MelderName: 'Anna',
                TenantId: 'tid',
                Schulname: 'HTL Demo',
                Domain: 'schule.at',
                SeitenUrl: 'https://app.ms365.schule/',
                UserAgent: 'Mozilla',
                AnhangUrls: '[]',
                Created: '2026-03-20T10:00:00Z'
            }
        });
        expect(mapped.id).toBe('42');
        expect(mapped.title).toBe('Login hängt');
        expect(mapped.reporterEmail).toBe('a@schule.at');
        expect(mapped.schoolName).toBe('HTL Demo');
        expect(mapped.attachments).toEqual([]);
    });

    it('lehnt zu viele oder zu große Anhänge ab', () => {
        expect(() =>
            normalizeAttachments([
                { name: '1.png', contentType: 'image/png', base64: 'QQ==' },
                { name: '2.png', contentType: 'image/png', base64: 'QQ==' },
                { name: '3.png', contentType: 'image/png', base64: 'QQ==' },
                { name: '4.png', contentType: 'image/png', base64: 'QQ==' }
            ])
        ).toThrow(/Maximal/);

        expect(() =>
            normalizeAttachments([{ name: 'x.pdf', contentType: 'application/pdf', base64: 'QQ==' }])
        ).toThrow(/Bild/);
    });
});
