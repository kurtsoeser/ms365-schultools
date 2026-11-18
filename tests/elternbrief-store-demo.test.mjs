import { describe, expect, it, beforeEach, afterEach } from 'vitest';
import { ELTERNBRIEF_DEMO_SEED_TAG } from '../src/tools/elternbrief/elternbrief-demo.js';
import { STORAGE_KEY } from '../src/tools/elternbrief/elternbrief-model.js';
import {
    clearDemoSeeds,
    loadDemoData,
    loadState,
    saveState
} from '../src/tools/elternbrief/elternbrief-store.js';

describe('elternbrief-store demo local-first', () => {
    const store = new Map();
    const prev = globalThis.localStorage;

    beforeEach(() => {
        store.clear();
        globalThis.localStorage = {
            getItem: (k) => (store.has(k) ? store.get(k) : null),
            setItem: (k, v) => store.set(k, String(v)),
            removeItem: (k) => store.delete(k)
        };
        saveState({
            briefe: [
                {
                    id: 'real-1',
                    betreff: 'Echt',
                    inhalt: 'Hallo',
                    status: 'entwurf',
                    erstellerEmail: 'lehrer@schule.at',
                    quelle: 'ui'
                }
            ],
            empfaenger: [],
            updatedAt: new Date().toISOString()
        });
    });

    afterEach(() => {
        globalThis.localStorage = prev;
    });

    it('lädt Demo ohne echte Briefe zu löschen und reset nur Seed', () => {
        loadDemoData({ email: 'lehrer@schule.at', name: 'Lehrkraft' });
        const after = loadState();
        expect(after.briefe.some((b) => b.id === 'real-1')).toBe(true);
        expect(after.briefe.some((b) => b.quelle === 'demo')).toBe(true);
        expect(after.briefe.some((b) => b.seedTag === ELTERNBRIEF_DEMO_SEED_TAG)).toBe(true);

        clearDemoSeeds();
        const cleaned = loadState();
        expect(cleaned.briefe).toHaveLength(1);
        expect(cleaned.briefe[0].id).toBe('real-1');
        expect(cleaned.empfaenger).toHaveLength(0);
        expect(store.has(STORAGE_KEY)).toBe(true);
    });
});
