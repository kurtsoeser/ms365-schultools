import { describe, expect, it } from 'vitest';
import {
    normalizeNote,
    plainTextToHtml,
    sanitizeReleaseHtml,
    getNewReleaseNotes,
    toPublishedJson
} from '../src/shared/release-notes-store.js';

describe('release-notes-store', () => {
    it('normalisiert bodyHtml und kind', () => {
        const n = normalizeNote(
            {
                title: 'Test',
                body: 'Zeile 1\n\nZeile 2',
                kind: 'feat',
                images: [{ src: 'https://example.com/a.png', alt: 'A' }]
            },
            0
        );
        expect(n.kind).toBe('feature');
        expect(n.bodyHtml).toContain('<p>');
        expect(n.images).toHaveLength(1);
    });

    it('sanitized gefährliches HTML weg', () => {
        const html = sanitizeReleaseHtml('<p>Hi<script>alert(1)</script></p><img src=x onerror=alert(1)>');
        expect(html.toLowerCase()).not.toContain('script');
        expect(html.toLowerCase()).not.toContain('onerror');
        expect(html).toContain('Hi');
    });

    it('plainTextToHtml escaped', () => {
        expect(plainTextToHtml('a < b')).toContain('&lt;');
    });

    it('getNewReleaseNotes filtert nach lastSeen', () => {
        const notes = [
            { id: '1', at: '2026-01-02T00:00:00.000Z', title: 'neu', bodyHtml: '<p>x</p>' },
            { id: '2', at: '2025-01-01T00:00:00.000Z', title: 'alt', bodyHtml: '<p>y</p>' }
        ].map((n, i) => normalizeNote(n, i));
        const neu = getNewReleaseNotes({ notes, lastSeenAtIso: '2026-01-01T00:00:00.000Z' });
        expect(neu.map((n) => n.id)).toEqual(['1']);
    });

    it('toPublishedJson ist gültiges JSON-Array', () => {
        const raw = toPublishedJson([normalizeNote({ title: 'A', bodyHtml: '<p>B</p>' }, 0)]);
        const parsed = JSON.parse(raw);
        expect(Array.isArray(parsed)).toBe(true);
        expect(parsed[0].title).toBe('A');
    });
});
