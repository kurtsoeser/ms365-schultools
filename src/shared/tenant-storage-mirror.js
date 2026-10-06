/**
 * v1-Spiegel (ms365-tenant-settings-v1): ab Epic A aus – kanonisch nur v2.
 * Siehe docs/stammdaten-v1-sunset.md
 */
export const TENANT_SETTINGS_V1_KEY = 'ms365-tenant-settings-v1';
export const TENANT_SETTINGS_V2_KEY = 'ms365-schooltool-data-v2';

export const V1_MIRROR_SUNSET_LABEL = '2026-H2';

/** Automatisches Schreiben des v1-Spiegels ist deaktiviert (hart). */
export const V1_MIRROR_WRITE_ENABLED = false;

const DISABLE_KEY = 'ms365-tenant-v1-mirror-disable';

export function isV1MirrorWriteEnabled() {
    return V1_MIRROR_WRITE_ENABLED === true;
}

/**
 * @param {object} normalized Ergebnis von ms365TenantSettingsSave / normalizeSettings
 */
export function writeTenantSettingsV1Mirror(normalized) {
    if (!isV1MirrorWriteEnabled()) return false;
    if (!normalized || typeof normalized !== 'object') return false;
    try {
        localStorage.setItem(TENANT_SETTINGS_V1_KEY, JSON.stringify(normalized));
        return true;
    } catch {
        return false;
    }
}

export function disableV1MirrorWrites() {
    try {
        localStorage.setItem(DISABLE_KEY, '1');
    } catch {
        /* ignore */
    }
}

/**
 * @param {Record<string, unknown>} localMap collectLocalStorage-Ergebnis
 */
export function buildLegacyMirrorExportMeta(localMap) {
    const local = localMap && typeof localMap === 'object' ? localMap : {};
    const hasV1 = Object.prototype.hasOwnProperty.call(local, TENANT_SETTINGS_V1_KEY);
    const hasV2 = Object.prototype.hasOwnProperty.call(local, TENANT_SETTINGS_V2_KEY);
    return {
        key: TENANT_SETTINGS_V1_KEY,
        role: 'legacy-read-only',
        writeEnabled: false,
        canonicalKey: TENANT_SETTINGS_V2_KEY,
        sunset: V1_MIRROR_SUNSET_LABEL,
        present: hasV1,
        canonicalPresent: hasV2,
        note:
            'Kanonisch ist nur ' +
            TENANT_SETTINGS_V2_KEY +
            '. Der Schlüssel ' +
            TENANT_SETTINGS_V1_KEY +
            ' wird nicht mehr automatisch geschrieben (Epic A).'
    };
}

if (typeof window !== 'undefined') {
    disableV1MirrorWrites();
    window.ms365TenantStorageMirror = {
        TENANT_SETTINGS_V1_KEY,
        TENANT_SETTINGS_V2_KEY,
        V1_MIRROR_SUNSET_LABEL,
        V1_MIRROR_WRITE_ENABLED,
        isV1MirrorWriteEnabled,
        writeTenantSettingsV1Mirror,
        disableV1MirrorWrites,
        buildLegacyMirrorExportMeta
    };
}
