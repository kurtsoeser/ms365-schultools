import { describe, expect, it, beforeEach, vi } from 'vitest';
import {
    writeTenantSettingsV1Mirror,
    isV1MirrorWriteEnabled,
    disableV1MirrorWrites,
    TENANT_SETTINGS_V1_KEY,
    TENANT_SETTINGS_V2_KEY,
    V1_MIRROR_WRITE_ENABLED,
    buildLegacyMirrorExportMeta
} from '../src/shared/tenant-storage-mirror.js';

describe('tenant-storage-mirror', () => {
    beforeEach(() => {
        const map = new Map();
        vi.stubGlobal('localStorage', {
            getItem: (k) => (map.has(k) ? map.get(k) : null),
            setItem: (k, v) => map.set(k, String(v)),
            removeItem: (k) => map.delete(k),
            clear: () => map.clear()
        });
    });

    it('schreibt v1-Spiegel nicht mehr (Epic A)', () => {
        expect(V1_MIRROR_WRITE_ENABLED).toBe(false);
        expect(isV1MirrorWriteEnabled()).toBe(false);
        expect(writeTenantSettingsV1Mirror({ domain: 'schule.at', classes: [] })).toBe(false);
        expect(localStorage.getItem(TENANT_SETTINGS_V1_KEY)).toBeNull();
    });

    it('disableV1MirrorWrites bleibt gesetzt', () => {
        disableV1MirrorWrites();
        expect(isV1MirrorWriteEnabled()).toBe(false);
    });

    it('buildLegacyMirrorExportMeta markiert legacy-read-only', () => {
        const meta = buildLegacyMirrorExportMeta({
            [TENANT_SETTINGS_V1_KEY]: '{}',
            [TENANT_SETTINGS_V2_KEY]: '{}'
        });
        expect(meta.role).toBe('legacy-read-only');
        expect(meta.writeEnabled).toBe(false);
        expect(meta.canonicalKey).toBe(TENANT_SETTINGS_V2_KEY);
    });
});
