'use strict';

const { getConfig } = require('./config');

/**
 * License-Header aus der Anfrage lesen.
 * @param {import('@azure/functions').HttpRequest} request
 */
function licenseTokenFromRequest(request) {
    const h =
        (request.headers &&
            (request.headers.get('x-ms365-license-authorization') ||
                request.headers.get('X-MS365-License-Authorization'))) ||
        '';
    const m = String(h).match(/^Bearer\s+(.+)$/i);
    return m ? m[1].trim() : '';
}

/**
 * Serverseitig: Tenant muss laut License-API freigeschaltet sein.
 * @param {{ tid: string }} caller
 * @param {string} licenseAccessToken
 */
async function assertTenantLicenseAllowed(caller, licenseAccessToken) {
    const cfg = getConfig();
    const base = String(cfg.licenseApiBaseUrl || '')
        .trim()
        .replace(/\/+$/, '');
    if (!base) {
        const err = new Error('LICENSE_API_BASE_URL ist im Kursteams-Backend nicht gesetzt.');
        err.status = 503;
        throw err;
    }
    const raw = String(licenseAccessToken || '').trim();
    if (!raw) {
        const err = new Error('Lizenz-Token fehlt.');
        err.status = 401;
        throw err;
    }

    let res;
    try {
        res = await fetch(base + '/me', {
            method: 'GET',
            headers: {
                Accept: 'application/json',
                Authorization: 'Bearer ' + raw
            }
        });
    } catch (e) {
        const err = new Error('Lizenzprüfung nicht erreichbar.');
        err.status = 503;
        err.cause = e;
        throw err;
    }

    let data = null;
    const text = await res.text();
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = null;
        }
    }

    if (!res.ok) {
        const err = new Error(
            res.status === 401 || res.status === 403
                ? 'Lizenzprüfung abgelehnt.'
                : 'Lizenzprüfung fehlgeschlagen.'
        );
        err.status = res.status === 401 || res.status === 403 ? res.status : 403;
        throw err;
    }

    const tid = String((data && data.tenantId) || '')
        .trim()
        .toLowerCase();
    const callerTid = String(caller && caller.tid ? caller.tid : '')
        .trim()
        .toLowerCase();
    if (!tid || !callerTid || tid !== callerTid) {
        const err = new Error('Lizenz-Token passt nicht zum Mandanten.');
        err.status = 403;
        throw err;
    }
    if (!data || data.allowed !== true) {
        const err = new Error(
            (data && data.message) || 'Dieser Mandant ist nicht freigeschaltet.'
        );
        err.status = 403;
        throw err;
    }
}

module.exports = { assertTenantLicenseAllowed, licenseTokenFromRequest };
