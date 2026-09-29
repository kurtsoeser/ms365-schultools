import { describe, expect, it, beforeEach, vi } from 'vitest';
import { loadPlaybookState, savePlaybookState } from '../src/shared/playbook-store.js';
import { createBulkProgress } from '../src/shared/bulk-progress.js';
import { guardApplyAgainstTruncation } from '../src/shared/truncation-banner.js';

describe('playbook-store', () => {
    beforeEach(() => {
        const map = new Map();
        vi.stubGlobal('localStorage', {
            getItem: (k) => (map.has(k) ? map.get(k) : null),
            setItem: (k, v) => map.set(k, String(v)),
            removeItem: (k) => map.delete(k),
            clear: () => map.clear()
        });
    });

    it('speichert und lädt Haken', () => {
        savePlaybookState('ms365-test-pb-v1', { a: true, b: false });
        expect(loadPlaybookState('ms365-test-pb-v1')).toEqual({ a: true, b: false });
    });

    it('liefert {} bei kaputtem JSON', () => {
        localStorage.setItem('ms365-test-pb-v1', '{');
        expect(loadPlaybookState('ms365-test-pb-v1')).toEqual({});
    });
});

describe('bulk-progress', () => {
    it('liefert No-Op-API ohne Root', () => {
        const api = createBulkProgress(null);
        expect(api.isCancelled()).toBe(false);
        api.show('x', 3);
        api.set(1, 3);
        api.addError('e');
        expect(api.exportErrorsCsv()).toBe('');
    });
});

describe('truncation-banner guard', () => {
    it('sperrt Apply bei Truncation', () => {
        const btn = { disabled: false, title: '', removeAttribute(name) {
            if (name === 'title') this.title = '';
        } };
        expect(guardApplyAgainstTruncation(true, btn)).toBe(false);
        expect(btn.disabled).toBe(true);
        expect(guardApplyAgainstTruncation(false, btn)).toBe(true);
        expect(btn.disabled).toBe(false);
    });
});
