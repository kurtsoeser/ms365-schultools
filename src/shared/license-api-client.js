/**
 * Client für License-API inkl. Admin-CRUD (Phase 5).
 */
(function (global) {
    'use strict';

    function apiConfig() {
        return global.MS365_LICENSE_API || {};
    }

    function baseUrl() {
        return String(apiConfig().baseUrl || '').replace(/\/+$/, '');
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
        fetchLicenseMe: fetchLicenseMe,
        adminListSchools: adminListSchools,
        adminCreateSchool: adminCreateSchool,
        adminUpdateSchool: adminUpdateSchool,
        adminDeleteSchool: adminDeleteSchool,
        adminListColumns: adminListColumns,
        adminCreateColumn: adminCreateColumn,
        adminDeleteColumn: adminDeleteColumn,
        fetchKursteamCatalog: fetchKursteamCatalog,
        publishKursteamCatalog: publishKursteamCatalog
    };
})(typeof window !== 'undefined' ? window : globalThis);
