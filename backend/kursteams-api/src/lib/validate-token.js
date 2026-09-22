'use strict';

const { createRemoteJWKSet, jwtVerify } = require('jose');

/** @type {Map<string, ReturnType<typeof createRemoteJWKSet>>} */
const jwksCache = new Map();

const JWKS_URLS = [
    'https://login.microsoftonline.com/common/discovery/v2.0/keys',
    'https://login.microsoftonline.com/common/discovery/keys'
];

const GUID_RE =
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function getJwks(url) {
    const key = String(url);
    if (!jwksCache.has(key)) {
        jwksCache.set(key, createRemoteJWKSet(new URL(key)));
    }
    return jwksCache.get(key);
}

/**
 * @param {string} iss
 * @param {string} tid
 */
function issuerMatchesTenant(iss, tid) {
    const issuer = String(iss || '');
    const tenant = String(tid || '').toLowerCase();
    if (!issuer || !GUID_RE.test(tenant)) return false;
    const patterns = [
        new RegExp('^https://login\\.microsoftonline\\.com/' + tenant + '/v2\\.0/?$', 'i'),
        new RegExp('^https://sts\\.windows\\.net/' + tenant + '/?$', 'i'),
        new RegExp('^https://login\\.microsoftonline\\.com/' + tenant + '/?$', 'i')
    ];
    return patterns.some((re) => re.test(issuer));
}

/**
 * @param {Record<string, unknown>} payload
 */
function tokenHasKursteamsScope(payload) {
    const raw = payload && (payload.scp || payload.scope) ? payload.scp || payload.scope : '';
    return String(raw)
        .split(/[\s,]+/)
        .map((s) => s.trim().toLowerCase())
        .includes('kursteams.create');
}

/**
 * @param {Record<string, unknown>} payload
 */
function claimsFromPayload(payload) {
    const tid = String(payload.tid || '')
        .trim()
        .toLowerCase();
    const oid = String(payload.oid || '')
        .trim()
        .toLowerCase();
    const upn = String(payload.preferred_username || payload.upn || '').trim();
    return { tid, oid, upn };
}

/**
 * @param {string} raw
 * @param {string[]} audiences
 */
async function verifyWithJwks(raw, audiences) {
    if (!Array.isArray(audiences) || !audiences.length) {
        const err = new Error('Token-Audience ist nicht konfiguriert.');
        err.status = 500;
        throw err;
    }
    let lastErr = null;
    for (let i = 0; i < JWKS_URLS.length; i++) {
        try {
            const verified = await jwtVerify(raw, getJwks(JWKS_URLS[i]), {
                audience: audiences,
                clockTolerance: 60
            });
            return verified.payload;
        } catch (e) {
            lastErr = e;
        }
    }
    const err = new Error('Token ungültig oder abgelaufen.');
    err.status = 401;
    err.cause = lastErr;
    throw err;
}

/**
 * Delegiertes Access-Token der Kursteams-API prüfen.
 * Audience ist nur die Backend-App, Graph-Token werden abgelehnt.
 * @param {string} token
 * @param {string[]} audiences
 * @returns {Promise<{ tid: string, oid: string, upn: string }>}
 */
async function validateCallerToken(token, audiences) {
    const raw = String(token || '').trim();
    if (!raw) {
        const err = new Error('Anmeldung fehlt.');
        err.status = 401;
        throw err;
    }

    const payload = await verifyWithJwks(raw, audiences);
    if (String(payload.idtyp || '').toLowerCase() === 'app') {
        const err = new Error('Token ungültig oder abgelaufen.');
        err.status = 401;
        throw err;
    }
    if (!tokenHasKursteamsScope(payload)) {
        const err = new Error('Token ungültig oder abgelaufen.');
        err.status = 401;
        throw err;
    }

    const claims = claimsFromPayload(payload);
    if (!GUID_RE.test(claims.tid) || !GUID_RE.test(claims.oid)) {
        const err = new Error('Token ungültig oder abgelaufen.');
        err.status = 401;
        throw err;
    }
    if (!issuerMatchesTenant(payload.iss, claims.tid)) {
        const err = new Error('Token ungültig oder abgelaufen.');
        err.status = 401;
        throw err;
    }
    return claims;
}

module.exports = {
    validateCallerToken,
    issuerMatchesTenant,
    tokenHasKursteamsScope,
    claimsFromPayload,
    GUID_RE
};
