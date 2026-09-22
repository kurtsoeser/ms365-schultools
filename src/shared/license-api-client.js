/**
 * Client für License-API inkl. Admin-CRUD (Phase 5).
 */
(function (global) {
    'use strict';

    var DEFAULT_SCOPE = 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/License.Access';

    function apiConfig() {
        return global.MS365_LICENSE_API || {};
    }

    function baseUrl() {
        return String(apiConfig().baseUrl || '').replace(/\/+$/, '');
    }

    function licenseScope() {
        return String(apiConfig().scope || DEFAULT_SCOPE).trim() || DEFAULT_SCOPE;
    }

    function authHeaders(accessToken) {
        const headers = {
            Accept: 'application/json',
            Authorization: 'Bearer ' + String(accessToken || '').trim()
        };
        const key = String(apiConfig().functionKey || '').trim();
        if (key) headers['x-functions-key'] = key;
        return headers;
    }

    async function acquireLicenseToken() {
        var scopes = [licenseScope()];
        if (typeof global.ms365AuthAcquireToken === 'function') {
            return global.ms365AuthAcquireToken(scopes);
        }
        if (typeof global.ms365AuthAcquireTokenPopup === 'function') {
            return global.ms365AuthAcquireTokenPopup(scopes);
        }
        throw new Error('Bitte zuerst mit MS365 anmelden.');
    }

    async function parseResponse(res) {
        const text = await res.text();
        let data = null;
        if (text) {
            try {
                data = JSON.parse(text);
            } catch {
                data = { message: text };
            }
        }
        if (!res.ok) {
            const msg = (data && (data.error || data.message)) || 'HTTP ' + res.status;
            const err = new Error(String(msg));
            err.status = res.status;
            err.payload = data;
            throw err;
        }
        return data || {};
    }

    async function fetchLicenseMe(accessToken) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        const res = await fetch(base + '/me', {
            method: 'GET',
            headers: authHeaders(accessToken)
        });
        return parseResponse(res);
    }

    async function adminListSchools(accessToken) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        const res = await fetch(base + '/admin/schools', {
            method: 'GET',
            headers: authHeaders(accessToken)
        });
        return parseResponse(res);
    }

    async function adminCreateSchool(accessToken, body) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        const res = await fetch(base + '/admin/schools', {
            method: 'POST',
            headers: Object.assign({ 'Content-Type': 'application/json' }, authHeaders(accessToken)),
            body: JSON.stringify(body || {})
        });
        return parseResponse(res);
    }

    async function adminUpdateSchool(accessToken, id, body) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        const res = await fetch(base + '/admin/schools/' + encodeURIComponent(id), {
            method: 'PATCH',
            headers: Object.assign({ 'Content-Type': 'application/json' }, authHeaders(accessToken)),
            body: JSON.stringify(body || {})
        });
        return parseResponse(res);
    }

    async function adminDeleteSchool(accessToken, id) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        const res = await fetch(base + '/admin/schools/' + encodeURIComponent(id), {
            method: 'DELETE',
            headers: authHeaders(accessToken)
        });
        return parseResponse(res);
    }

    async function adminListColumns(accessToken) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        const res = await fetch(base + '/admin/columns', {
            method: 'GET',
            headers: authHeaders(accessToken)
        });
        return parseResponse(res);
    }

    async function adminCreateColumn(accessToken, body) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        const res = await fetch(base + '/admin/columns', {
            method: 'POST',
            headers: Object.assign({ 'Content-Type': 'application/json' }, authHeaders(accessToken)),
            body: JSON.stringify(body || {})
        });
        return parseResponse(res);
    }

    async function fetchKursteamCatalog(accessToken) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        const res = await fetch(base + '/catalog/kursteam-templates', {
            method: 'GET',
            headers: authHeaders(accessToken)
        });
        return parseResponse(res);
    }

    async function publishKursteamCatalog(accessToken, body) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        const res = await fetch(base + '/catalog/kursteam-templates', {
            method: 'PUT',
            headers: Object.assign({ 'Content-Type': 'application/json' }, authHeaders(accessToken)),
            body: JSON.stringify(body || {})
        });
        return parseResponse(res);
    }

    async function fetchCatalogMaterials(accessToken, path) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        const q = path ? '?path=' + encodeURIComponent(path) : '';
        const res = await fetch(base + '/catalog/materials' + q, {
            method: 'GET',
            headers: authHeaders(accessToken)
        });
        return parseResponse(res);
    }

    async function fetchCatalogMaterialFile(accessToken, path) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        if (!path) throw new Error('Material-Pfad fehlt.');
        const res = await fetch(
            base + '/catalog/materials/file?path=' + encodeURIComponent(path),
            {
                method: 'GET',
                headers: authHeaders(accessToken)
            }
        );
        if (!res.ok) {
            let msg = 'HTTP ' + res.status;
            try {
                const data = await res.json();
                if (data && (data.error || data.message)) msg = String(data.error || data.message);
            } catch {
                /* ignore */
            }
            const err = new Error(msg);
            err.status = res.status;
            throw err;
        }
        const blob = await res.blob();
        const disposition = res.headers.get('Content-Disposition') || '';
        let name = '';
        const star = disposition.match(/filename\*=UTF-8''([^;]+)/i);
        const plain = disposition.match(/filename="([^"]+)"/i);
        if (star) {
            try {
                name = decodeURIComponent(star[1]);
            } catch {
                name = star[1];
            }
        } else if (plain) {
            name = plain[1];
        }
        if (!name) name = String(path).split('/').pop() || 'datei';
        return {
            name,
            path,
            contentType: blob.type || 'application/octet-stream',
            blob,
            bytes: new Uint8Array(await blob.arrayBuffer())
        };
    }

    async function createCatalogMaterialFolder(accessToken, body) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        const res = await fetch(base + '/catalog/materials/folder', {
            method: 'POST',
            headers: Object.assign({ 'Content-Type': 'application/json' }, authHeaders(accessToken)),
            body: JSON.stringify(body || {})
        });
        return parseResponse(res);
    }

    async function uploadCatalogMaterialFile(accessToken, path, bytes, contentType) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        if (!path) throw new Error('Material-Pfad fehlt.');
        const res = await fetch(
            base + '/catalog/materials/file?path=' + encodeURIComponent(path),
            {
                method: 'PUT',
                headers: Object.assign(
                    { 'Content-Type': contentType || 'application/octet-stream' },
                    authHeaders(accessToken)
                ),
                body: bytes
            }
        );
        return parseResponse(res);
    }

    async function deleteCatalogMaterial(accessToken, path) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        if (!path) throw new Error('Material-Pfad fehlt.');
        const res = await fetch(
            base + '/catalog/materials?path=' + encodeURIComponent(path),
            {
                method: 'DELETE',
                headers: authHeaders(accessToken)
            }
        );
        return parseResponse(res);
    }

    async function adminDeleteColumn(accessToken, name) {
        const base = baseUrl();
        if (!base) throw new Error('MS365_LICENSE_API.baseUrl ist nicht gesetzt.');
        const res = await fetch(base + '/admin/columns/' + encodeURIComponent(name), {
            method: 'DELETE',
            headers: authHeaders(accessToken)
        });
        return parseResponse(res);
    }

    global.ms365LicenseApi = {
        acquireLicenseToken: acquireLicenseToken,
        fetchLicenseMe: fetchLicenseMe,
        adminListSchools: adminListSchools,
        adminCreateSchool: adminCreateSchool,
        adminUpdateSchool: adminUpdateSchool,
        adminDeleteSchool: adminDeleteSchool,
        adminListColumns: adminListColumns,
        adminCreateColumn: adminCreateColumn,
        adminDeleteColumn: adminDeleteColumn,
        fetchKursteamCatalog: fetchKursteamCatalog,
        publishKursteamCatalog: publishKursteamCatalog,
        fetchCatalogMaterials: fetchCatalogMaterials,
        fetchCatalogMaterialFile: fetchCatalogMaterialFile,
        createCatalogMaterialFolder: createCatalogMaterialFolder,
        uploadCatalogMaterialFile: uploadCatalogMaterialFile,
        deleteCatalogMaterial: deleteCatalogMaterial
    };
})(typeof window !== 'undefined' ? window : globalThis);
