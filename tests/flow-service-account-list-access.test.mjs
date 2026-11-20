import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import {
    effectiveFlowServiceAccountFromCfg,
    loadFlowServiceAccountFromStorage,
    resolveFlowServiceAccount,
    normalizeFlowServiceEmails,
    FLOW_SETUP_STORAGE_KEYS
} from '../src/shared/flow-service-account-list-access.js';

function memoryStorage(map = {}) {
    return {
        getItem(k) {
            return Object.prototype.hasOwnProperty.call(map, k) ? map[k] : null;
        },
        setItem(k, v) {
            map[k] = String(v);
        }
    };
}

describe('flow-service-account-list-access', () => {
    it('effectiveFlowServiceAccountFromCfg bevorzugt flowServiceAccount', () => {
        expect(
            effectiveFlowServiceAccountFromCfg({
                flowServiceAccount: 'automate@schule.at',
                emailMailbox: 'other@schule.at'
            })
        ).toBe('automate@schule.at');
        expect(effectiveFlowServiceAccountFromCfg({ emailMailbox: 'mail@schule.at' })).toBe(
            'mail@schule.at'
        );
        expect(effectiveFlowServiceAccountFromCfg({})).toBe('');
    });

    it('normalizeFlowServiceEmails dedupliziert', () => {
        expect(normalizeFlowServiceEmails(['A@x.at', 'a@x.at', '', 'b@x.at'])).toEqual([
            'a@x.at',
            'b@x.at'
        ]);
    });

    it('lädt Technik-Konto aus Setup-Storage', () => {
        const store = memoryStorage({
            [FLOW_SETUP_STORAGE_KEYS.akt]: JSON.stringify({
                flowServiceAccount: 'automate@ms365.schule'
            })
        });
        expect(loadFlowServiceAccountFromStorage(FLOW_SETUP_STORAGE_KEYS.akt, store)).toBe(
            'automate@ms365.schule'
        );
        expect(loadFlowServiceAccountFromStorage(FLOW_SETUP_STORAGE_KEYS.akt, memoryStorage())).toBe(
            ''
        );
    });

    it('resolveFlowServiceAccount: explizit schlägt Storage', () => {
        expect(resolveFlowServiceAccount('lfr', 'technik@schule.at')).toBe('technik@schule.at');
    });
});

describe('resolveFlowServiceAccount mit localStorage', () => {
    const prev = globalThis.localStorage;
    beforeEach(() => {
        const map = {};
        globalThis.localStorage = memoryStorage(map);
        localStorage.setItem(
            FLOW_SETUP_STORAGE_KEYS.lfr,
            JSON.stringify({ flowServiceAccount: 'auto@demo.at' })
        );
    });
    afterEach(() => {
        if (prev === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prev;
    });

    it('liest LFR-Setup wenn kein Explicit', () => {
        expect(resolveFlowServiceAccount('lfr')).toBe('auto@demo.at');
    });
});
