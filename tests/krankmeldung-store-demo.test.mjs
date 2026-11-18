import { describe, expect, it, beforeEach, afterEach } from 'vitest';
import { KRANKMELDUNG_DEMO_SEED_TAG } from '../src/tools/krankmeldung/krankmeldung-demo.js';
import { STORAGE_KEY } from '../src/tools/krankmeldung/krankmeldung-model.js';
import {
    clearDemoSeeds,
    loadDemoData,
    loadState,
    saveState
} from '../src/tools/krankmeldung/krankmeldung-store.js';

describe('krankmeldung-store demo local-first', () => {
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
            meldungen: [
                {
                    id: 'real-1',
                    schuelerName: 'Echt',
                    klasseCode: '1A',
                    von: '2026-11-18',
                    bis: '2026-11-18',
                    kategorie: 'krankheit',
                    status: 'neu',
                    quelle: 'public'
                }
            ]
        });
    });

    afterEach(() => {
        globalThis.localStorage = prev;
    });

    it('lädt Demo ohne echte Meldungen zu löschen und reset nur Seed', () => {
        loadDemoData();
        const after = loadState();
        expect(after.meldungen.some((m) => m.id === 'real-1')).toBe(true);
        expect(after.meldungen.some((m) => m.quelle === 'demo')).toBe(true);
        expect(after.meldungen.some((m) => m.seedTag === KRANKMELDUNG_DEMO_SEED_TAG)).toBe(true);

        clearDemoSeeds();
        const cleaned = loadState();
        expect(cleaned.meldungen).toHaveLength(1);
        expect(cleaned.meldungen[0].id).toBe('real-1');
        expect(store.has(STORAGE_KEY)).toBe(true);
    });
});
