import { describe, it, expect } from 'vitest';
import {
    CONFIG_BUNDLE_PARTS,
    CONFIG_MANIFEST_FILE,
    SCHUL_ORGANIGRAMM_CONFIG_PART_ID,
    SCHUL_ORGANIGRAMM_CONFIG_FILE,
    buildConfigBundleFromBackup,
    extractSchulOrganigrammFromStore,
    mergeSchulOrganigrammIntoSchooltoolContainer,
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

    it('legt Schul-Organigramm als eigene Config-Datei ab', () => {
        const v2 = {
            core: {
                schoolName: 'Demo',
                schulOrganigramm: {
                    version: 2,
                    nodes: [
                        { id: 'sec_a', type: 'area', label: 'SL' },
                        { id: 'role_1', type: 'role', label: 'Direktor' }
                    ]
                }
            }
        };
        const extracted = extractSchulOrganigrammFromStore({
            'ms365-schooltool-data-v2': v2
        });
        expect(extracted.nodeCount).toBe(2);
        const bundle = buildConfigBundleFromBackup({
            schoolName: 'Demo',
            localStorage: { 'ms365-schooltool-data-v2': v2 }
        });
        const orgFile = bundle.files.find((f) => f.id === SCHUL_ORGANIGRAMM_CONFIG_PART_ID);
        expect(orgFile).toBeTruthy();
        expect(orgFile.path).toBe('config/' + SCHUL_ORGANIGRAMM_CONFIG_FILE);
        expect(orgFile.body.nodeCount).toBe(2);
        expect(orgFile.body.schulOrganigramm.nodes).toHaveLength(2);
        const parsed = parseConfigBundlePart(orgFile.body, SCHUL_ORGANIGRAMM_CONFIG_PART_ID);
        expect(parsed.schulOrganigramm.nodes[1].label).toBe('Direktor');
        const merged = mergeSchulOrganigrammIntoSchooltoolContainer(
            { core: { schoolName: 'X', schulOrganigramm: { nodes: [] } }, years: {} },
            parsed.schulOrganigramm
        );
        expect(merged.core.schulOrganigramm.nodes).toHaveLength(2);
        expect(merged.core.schoolName).toBe('X');
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
