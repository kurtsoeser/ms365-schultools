import { describe, expect, it } from 'vitest';
import {
    evaluateSchuljahresstartPrerequisites,
    evaluateStepUnlock,
    schuljahresstartPlaybookProgress,
    nextSchuljahresstartGateHint,
    tenantCountsFromSettings,
    SCHULJAHRSTART_STEP_IDS
} from '../src/shared/playbook-schuljahresstart-gates.js';

describe('playbook-schuljahresstart-gates', () => {
    it('sperrt Kursteams ohne Fächer', () => {
        const gates = evaluateSchuljahresstartPrerequisites({
            year: '2025/26',
            classes: 3,
            students: 10,
            teachers: 2,
            subjects: 0,
            domain: 'schule.at'
        });
        expect(gates.kursteams.ok).toBe(false);
        const unlock = evaluateStepUnlock('kursteams', 3, {}, gates, SCHULJAHRSTART_STEP_IDS);
        expect(unlock.unlocked).toBe(false);
    });

    it('öffnet Schritt nach abgehaktem Vorgänger', () => {
        const gates = evaluateSchuljahresstartPrerequisites({
            year: '2025/26',
            classes: 0,
            students: 0,
            teachers: 0,
            subjects: 0,
            domain: 'schule.at'
        });
        const state = { stammdaten: true };
        const unlock = evaluateStepUnlock('klassen', 1, state, gates, SCHULJAHRSTART_STEP_IDS);
        expect(unlock.unlocked).toBe(true);
    });

    it('zählt Playbook-Fortschritt', () => {
        const p = schuljahresstartPlaybookProgress({ stammdaten: true, klassen: true }, SCHULJAHRSTART_STEP_IDS);
        expect(p).toEqual({ done: 2, total: 8 });
    });

    it('liefert Gate-Hinweis für ersten offenen gesperrten Schritt', () => {
        const counts = tenantCountsFromSettings({
            domain: 'x.at',
            classes: [],
            students: [],
            teachers: [],
            subjects: []
        });
        const hint = nextSchuljahresstartGateHint({}, counts);
        expect(hint).toMatch(/Klasse|Schuljahr/i);
    });
});
