import { describe, expect, it, beforeEach, afterEach } from 'vitest';
import { EINSAMMLER_DEMO_SEED_TAG } from '../src/tools/einsammler/einsammler-demo.js';
import { STORAGE_KEY } from '../src/tools/einsammler/einsammler-model.js';
import {
    clearDemoSeeds,
    loadDemoData,
    loadState,
    saveState
} from '../src/tools/einsammler/einsammler-store.js';

describe('einsammler-store demo local-first', () => {
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
            einsammler: [
                {
                    id: 'real-1',
                    titel: 'Echt',
                    modus: 'einwilligung',
                    typ: 'sonstiges',
                    status: 'entwurf',
                    erstellerEmail: 'lehrer@schule.at',
                    quelle: 'ui'
                }
            ],
            antworten: []
        });
    });

    afterEach(() => {
        globalThis.localStorage = prev;
    });

    it('lädt Demo ohne echte Vorgänge zu löschen und reset nur Seed', () => {
        loadDemoData({ email: 'lehrer@schule.at', name: 'Lehrkraft' });
        const after = loadState();
        expect(after.einsammler.some((e) => e.id === 'real-1')).toBe(true);
        expect(after.einsammler.some((e) => e.quelle === 'demo')).toBe(true);
        expect(after.einsammler.some((e) => e.seedTag === EINSAMMLER_DEMO_SEED_TAG)).toBe(true);

        clearDemoSeeds();
        const cleaned = loadState();
        expect(cleaned.einsammler).toHaveLength(1);
        expect(cleaned.einsammler[0].id).toBe('real-1');
        expect(cleaned.antworten).toHaveLength(0);
        expect(store.has(STORAGE_KEY)).toBe(true);
    });
});
