import { describe, it, expect } from 'vitest';
import {
    CONFIG_BUNDLE_PARTS,
    CONFIG_MANIFEST_FILE,
    buildConfigBundleFromBackup,
    isConfigBundleManifest,
    parseConfigBundlePart,
    mergeConfigPartsLocalStorage,
    fingerprintLocalStorageSlice,
    configPartRelativePath
} from '../src/shared/stammdaten-sharepoint-config-bundle.js';

describe('stammdaten-sharepoint-config-bundle', () => {
    it('baut Manifest und vier Teil-Dateien', () => {
        const payload = {
            exportedAt: '2026-10-04T12:00:00.000Z',
            schoolName: 'Demo',
            domain: 'demo.at',
            contentFingerprint: 'abc:5',
            localStorage: {
                'ms365-schooltool-data-v2': '{"a":1}',
                'ms365-dashboard-tool-access-v1': '{}',
                'ms365-schularbeiten-perms-v1': '{}',
                'ms365-unknown-local-only': 'x'
            }
        };
        const bundle = buildConfigBundleFromBackup(payload, {
            tenantId: 'tid-1',
            monolithPath: 'Backups/ms365-stammdaten-aktuell.json'
        });
        expect(bundle.files).toHaveLength(CONFIG_BUNDLE_PARTS.length);
        expect(CONFIG_BUNDLE_PARTS.some((p) => p.id === 'permissions-freistellung')).toBe(true);
        expect(CONFIG_BUNDLE_PARTS.some((p) => p.fileName === 'permissions-freistellung.json')).toBe(true);
        expect(bundle.manifestPath).toBe('config/manifest.json');
        expect(isConfigBundleManifest(bundle.manifest)).toBe(true);
        expect(bundle.manifest.monolithPath).toBe('Backups/ms365-stammdaten-aktuell.json');
        expect(bundle.manifest.contentFingerprint).toBe('abc:5');
        const core = bundle.files.find((f) => f.id === 'stammdaten-core');
        expect(core.body.localStorage['ms365-schooltool-data-v2']).toBe('{"a":1}');
        expect(core.body.localStorage['ms365-unknown-local-only']).toBeUndefined();
    });

    it('parst Teil-Datei und merged Patches', () => {
        const slice = { 'ms365-tenant-settings-v1': '{}' };
        const body = {
            bundleKind: 'ms365-spo-config-bundle-v1',
            bundlePart: 'stammdaten-core',
            localStorage: slice,
            fingerprint: fingerprintLocalStorageSlice(slice)
        };
        const parsed = parseConfigBundlePart(body, 'stammdaten-core');
        expect(parsed).not.toBeNull();
        expect(parsed.localStorage['ms365-tenant-settings-v1']).toBe('{}');
        const merged = mergeConfigPartsLocalStorage([
            { localStorage: { a: '1' } },
            { localStorage: { b: '2', a: '9' } }
        ]);
        expect(merged).toEqual({ a: '9', b: '2' });
    });

    it('encodiert Config-Pfade', () => {
        expect(configPartRelativePath('config', CONFIG_MANIFEST_FILE)).toBe('config/manifest.json');
    });
});
