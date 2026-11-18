import { describe, it, expect } from 'vitest';
import {
    STAMMDATEN_LOCAL_FIRST,
    STAMMDATEN_IT_MANUAL_BACKUP,
    isStammdatenLocalFirst,
    isSpoManualBackupAllowed,
    isSpoNetworkSyncAllowed,
    isIntranetExportAllowed,
    localFirstStatusMessage,
    localFirstShortLabel
} from '../src/shared/stammdaten-local-first.js';
import {
    STAMMDATEN_CANONICAL_KEY,
    STAMMDATEN_SIDE_STATE_KEYS,
    isStammdatenSideStateKey,
    sideStatePolicyNote
} from '../src/shared/stammdaten-side-state-keys.js';
import { V1_MIRROR_WRITE_ENABLED, isV1MirrorWriteEnabled } from '../src/shared/tenant-storage-mirror.js';

describe('stammdaten-local-first', () => {
    it('Lokal-First bleibt aktiv: Auto-Sync und Intranet aus', () => {
        expect(STAMMDATEN_LOCAL_FIRST).toBe(true);
        expect(isStammdatenLocalFirst()).toBe(true);
        expect(isSpoNetworkSyncAllowed()).toBe(false);
        expect(isIntranetExportAllowed()).toBe(false);
        expect(localFirstStatusMessage()).toMatch(/Lokal-First/i);
        expect(localFirstShortLabel()).toMatch(/Lokal-First/i);
    });

    it('Phase 4: manuelle IT-Bibliothek ist freigeschaltet', () => {
        expect(STAMMDATEN_IT_MANUAL_BACKUP).toBe(true);
        expect(isSpoManualBackupAllowed()).toBe(true);
        expect(localFirstStatusMessage()).toMatch(/manuell/i);
        expect(localFirstStatusMessage()).toMatch(/kein Auto-Abgleich/i);
    });

    it('v1-Spiegel bleibt aus', () => {
        expect(V1_MIRROR_WRITE_ENABLED).toBe(false);
        expect(isV1MirrorWriteEnabled()).toBe(false);
    });

    it('Nebenstates sind katalogisiert und nicht kanonisch', () => {
        expect(STAMMDATEN_CANONICAL_KEY).toBe('ms365-schooltool-data-v2');
        expect(STAMMDATEN_SIDE_STATE_KEYS.length).toBeGreaterThan(5);
        expect(isStammdatenSideStateKey('webuntis-teams-creator-state-v1')).toBe(true);
        expect(isStammdatenSideStateKey('ms365-arge-state-v2')).toBe(true);
        expect(isStammdatenSideStateKey(STAMMDATEN_CANONICAL_KEY)).toBe(false);
        expect(sideStatePolicyNote()).toMatch(/Wahrheit/);
    });
});
