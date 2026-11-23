import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const { normalizeReleaseNotes, notesForView } = require('../src/lib/release-notes-payload.js');

describe('release-notes-payload', () => {
    it('behält nur alle auf der öffentlichen Sicht', () => {
        const notes = normalizeReleaseNotes([
            { id: 'a', title: 'Für alle', audience: 'all', bodyHtml: '<p>Hi</p>' },
            { id: 'i', title: 'IT', audience: 'Schul-IT', bodyHtml: '<p>intern</p>' },
            { id: 'e', title: 'Lehrer', audience: 'enduser', bodyHtml: '<p>x</p>' }
        ]);
        expect(notes.map((n) => n.audience)).toEqual(['all', 'it', 'enduser']);
        expect(notesForView(notes, 'public').map((n) => n.id)).toEqual(['a']);
        expect(notesForView(notes, 'full')).toHaveLength(3);
    });

    it('entfernt Script-Tags', () => {
        const notes = normalizeReleaseNotes([
            { title: 'Sicher', bodyHtml: '<p>Ok</p><script>alert(1)</script>' }
        ]);
        expect(notes[0].bodyHtml.toLowerCase()).not.toContain('script');
        expect(notes[0].bodyHtml).toContain('Ok');
    });
});
