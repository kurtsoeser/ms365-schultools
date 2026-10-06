import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import {
    peekWebuntisImportPayload,
    clearWebuntisImportPayload,
    consumeWebuntisImportPayload,
    resolveReturnUrl,
    WEBUNTIS_IMPORT_PAYLOAD_KEY
} from '../src/shared/webuntis-stammdaten-import-handoff.js';
import { buildWebuntisTenantReview } from '../src/shared/stammdaten-import-review.js';

function createMemorySession() {
    const map = new Map();
    return {
        getItem(k) {
            return map.has(k) ? map.get(k) : null;
        },
        setItem(k, v) {
            map.set(String(k), String(v));
        },
        removeItem(k) {
            map.delete(k);
        },
        clear() {
            map.clear();
        }
    };
}

describe('webuntis-stammdaten-import-handoff', () => {
    let prevSession;

    beforeEach(() => {
        prevSession = globalThis.sessionStorage;
        globalThis.sessionStorage = createMemorySession();
    });
    afterEach(() => {
        globalThis.sessionStorage = prevSession;
    });

    it('peek lässt Payload bis consume/clear bestehen', () => {
        sessionStorage.setItem(WEBUNTIS_IMPORT_PAYLOAD_KEY, JSON.stringify({ teachersLines: 'A;B;c@x.de' }));
        expect(peekWebuntisImportPayload()).toEqual({ teachersLines: 'A;B;c@x.de' });
        expect(peekWebuntisImportPayload()).toEqual({ teachersLines: 'A;B;c@x.de' });
        const c = consumeWebuntisImportPayload();
        expect(c.teachersLines).toBe('A;B;c@x.de');
        expect(peekWebuntisImportPayload()).toBeNull();
    });

    it('clear entfernt Payload', () => {
        sessionStorage.setItem(WEBUNTIS_IMPORT_PAYLOAD_KEY, JSON.stringify({ counts: { students: 1 } }));
        clearWebuntisImportPayload();
        expect(peekWebuntisImportPayload()).toBeNull();
    });

    it('resolveReturnUrl: Playbook als Rücksprung', () => {
        expect(resolveReturnUrl('playbook-daten-import')).toBe('playbook-daten-import-verknuepfen.html');
        expect(resolveReturnUrl('tenant')).toBe('../tenant.html');
    });
});

describe('buildWebuntisTenantReview', () => {
    it('fasst Stats und Payload-Zähler zusammen', () => {
        const merged = {
            stats: {
                teachers: { added: 1, updated: 0, unchanged: 2 },
                subjects: { added: 0, updated: 0, unchanged: 0 },
                classes: { added: 0, updated: 0, unchanged: 0 }
            },
            studentDiff: null
        };
        const review = buildWebuntisTenantReview(merged, { counts: { teachers: 1, students: 0, subjects: 0, classes: 0 } }, {});
        expect(review.title).toContain('1 Lehrer');
        expect(review.bullets.length).toBeGreaterThan(0);
        expect(review.hasConflicts).toBe(false);
    });
});
