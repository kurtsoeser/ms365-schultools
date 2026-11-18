import { describe, expect, it } from 'vitest';
import {
    DEMO_SEED_TAG,
    buildWahlfachDemoState,
    isWzDemoSeedState
} from '../src/tools/wahlfach-zuteilung/wahlfach-zuteilung-demo.js';

describe('wahlfach-zuteilung-demo', () => {
    it('liefert 6 Angebote und 15 Anmeldungen mit Seed-Tag', () => {
        const s = buildWahlfachDemoState();
        expect(s.runde).toBeTruthy();
        expect(s.angebote).toHaveLength(6);
        expect(s.anmeldungen).toHaveLength(15);
        expect(s.zuteilungen.length).toBeGreaterThanOrEqual(8);
        expect(s.anmeldungen.some((a) => a.status === 'eingereicht')).toBe(true);
        expect(s.anmeldungen.some((a) => a.status === 'zugeordnet')).toBe(true);
        expect(s.anmeldungen.every((a) => a.seedTag === DEMO_SEED_TAG)).toBe(true);
        expect(s.anmeldungen.every((a) => a.schuelerEmail)).toBe(true);
        expect(isWzDemoSeedState(s)).toBe(true);
        expect(isWzDemoSeedState({ runde: null, angebote: [], anmeldungen: [] })).toBe(false);
    });
});
