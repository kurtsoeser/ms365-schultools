/**
 * Zentraler Kursteam-Katalog (SharePoint des Betreibers, über die License-API).
 * Schul-Tenants lesen ihn mit ihrem eigenen Login – ohne Recht auf kurtrocks.
 */

export function catalogApiConfigured() {
    const cfg = window.MS365_LICENSE_API || {};
    return !!String(cfg.baseUrl || '').trim();
}

export function isLoggedIn() {
    return typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
}

export function isOperatorUser() {
    return !!(
        window.ms365OperatorAccess &&
        typeof window.ms365OperatorAccess.isCurrentUserOperator === 'function' &&
        window.ms365OperatorAccess.isCurrentUserOperator()
    );
}

async function catalogToken() {
    if (window.ms365LicenseApi && typeof window.ms365LicenseApi.acquireLicenseToken === 'function') {
        return window.ms365LicenseApi.acquireLicenseToken();
    }
    const scopes = ['https://graph.microsoft.com/User.Read'];
    if (typeof window.ms365AuthAcquireIdToken === 'function') {
        return window.ms365AuthAcquireIdToken(scopes);
    }
    if (typeof window.ms365AuthAcquireIdTokenPopup === 'function') {
        return window.ms365AuthAcquireIdTokenPopup(scopes);
    }
    throw new Error('Anmeldung nicht verfügbar.');
}

/**
 * @returns {Promise<{
 *   ok: boolean,
 *   missing?: boolean,
 *   templates: object[],
 *   schoolForms: string[],
 *   updatedAt?: string|null,
 *   updatedBy?: string,
 *   library?: string,
 *   path?: string,
 *   siteWebUrl?: string,
 *   webUrl?: string,
 *   message?: string
 * }>}
 */
export async function fetchCentralCatalog() {
    if (!catalogApiConfigured()) {
        return {
            ok: false,
            templates: [],
            schoolForms: [],
            message: 'License-API ist nicht konfiguriert. Es gilt nur die lokale Bibliothek.'
        };
    }
    if (!isLoggedIn()) {
        return {
            ok: false,
            templates: [],
            schoolForms: [],
            message: 'Anmelden, um die zentrale Bibliothek zu laden.'
        };
    }
    const api = window.ms365LicenseApi;
    if (!api || typeof api.fetchKursteamCatalog !== 'function') {
        return {
            ok: false,
            templates: [],
            schoolForms: [],
            message: 'Katalog-Client fehlt.'
        };
    }
    const token = await catalogToken();
    const data = await api.fetchKursteamCatalog(token);
    return {
        ok: true,
        missing: !!data.missing,
        templates: Array.isArray(data.templates) ? data.templates : [],
        schoolForms: Array.isArray(data.schoolForms) ? data.schoolForms : [],
        updatedAt: data.updatedAt || null,
        updatedBy: data.updatedBy || '',
        library: data.library || '',
        path: data.path || '',
        siteWebUrl: data.siteWebUrl || '',
        webUrl: data.webUrl || '',
        message: data.message || ''
    };
}

/**
 * @param {object} payload
 */
export async function publishCentralCatalog(payload) {
    if (!catalogApiConfigured()) {
        throw new Error('License-API ist nicht konfiguriert.');
    }
    if (!isOperatorUser()) {
        throw new Error('Nur das Betreiber-Konto kann die Zentrale veröffentlichen.');
    }
    const api = window.ms365LicenseApi;
    if (!api || typeof api.publishKursteamCatalog !== 'function') {
        throw new Error('Katalog-Client fehlt.');
    }
    const token = await catalogToken();
    return api.publishKursteamCatalog(token, payload);
}

/**
 * @param {string} [path]
 */
export async function fetchMaterials(path) {
    if (!catalogApiConfigured()) {
        throw new Error('License-API ist nicht konfiguriert.');
    }
    if (!isLoggedIn()) {
        throw new Error('Anmelden, um Materialien zu laden.');
    }
    const api = window.ms365LicenseApi;
    if (!api || typeof api.fetchCatalogMaterials !== 'function') {
        throw new Error('Katalog-Client fehlt.');
    }
    const token = await catalogToken();
    return api.fetchCatalogMaterials(token, path || '');
}

/**
 * @param {string} path
 */
export async function downloadMaterialFile(path) {
    if (!catalogApiConfigured()) {
        throw new Error('License-API ist nicht konfiguriert.');
    }
    if (!isLoggedIn()) {
        throw new Error('Anmelden, um Materialien zu laden.');
    }
    const api = window.ms365LicenseApi;
    if (!api || typeof api.fetchCatalogMaterialFile !== 'function') {
        throw new Error('Katalog-Client fehlt.');
    }
    const token = await catalogToken();
    return api.fetchCatalogMaterialFile(token, path);
}

/**
 * @param {object} body
 */
export async function createMaterialFolder(body) {
    if (!isOperatorUser()) throw new Error('Nur das Betreiber-Konto kann Ordner anlegen.');
    const api = window.ms365LicenseApi;
    if (!api || typeof api.createCatalogMaterialFolder !== 'function') {
        throw new Error('Katalog-Client fehlt.');
    }
    const token = await catalogToken();
    return api.createCatalogMaterialFolder(token, body || {});
}

/**
 * @param {string} path
 * @param {Uint8Array|ArrayBuffer|Blob} bytes
 * @param {string} [contentType]
 */
export async function uploadMaterialFile(path, bytes, contentType) {
    if (!isOperatorUser()) throw new Error('Nur das Betreiber-Konto kann Dateien hochladen.');
    const api = window.ms365LicenseApi;
    if (!api || typeof api.uploadCatalogMaterialFile !== 'function') {
        throw new Error('Katalog-Client fehlt.');
    }
    const token = await catalogToken();
    return api.uploadCatalogMaterialFile(token, path, bytes, contentType);
}

/**
 * @param {string} path
 */
export async function deleteMaterialItem(path) {
    if (!isOperatorUser()) throw new Error('Nur das Betreiber-Konto kann löschen.');
    const api = window.ms365LicenseApi;
    if (!api || typeof api.deleteCatalogMaterial !== 'function') {
        throw new Error('Katalog-Client fehlt.');
    }
    const token = await catalogToken();
    return api.deleteCatalogMaterial(token, path);
}
