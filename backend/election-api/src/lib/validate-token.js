'use strict';

const { createRemoteJWKSet, jwtVerify, decodeJwt } = require('jose');
const { getConfig } = require('./config');

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
 * @param {string[]} requiredScopes
 */
function tokenHasRequiredScope(payload, requiredScopes) {
    const raw = payload && (payload.scp || payload.scope) ? payload.scp || payload.scope : '';
    const have = String(raw)
        .split(/[\s,]+/)
        .map((s) => s.trim().toLowerCase())
        .filter(Boolean);
    if (!requiredScopes.length) return have.length > 0;
    return requiredScopes.some((need) => have.includes(String(need).toLowerCase()));
}

/**
 * @param {string} raw
 */
function peekTokenMeta(raw) {
    try {
        const p = decodeJwt(raw);
        const aud = Array.isArray(p.aud) ? p.aud.join(',') : String(p.aud || '');
        return {
            aud: aud.slice(0, 120),
            scp: String(p.scp || p.scope || '').slice(0, 120),
            tid: String(p.tid || '').slice(0, 40),
            iss: String(p.iss || '').slice(0, 80)
        };
    } catch {
        return null;
    }
}

/**
 * @param {Record<string, unknown>} payload
 */
function claimsFromPayload(payload) {
    return {
        tid: String(payload.tid || '')
            .trim()
            .toLowerCase(),
        oid: String(payload.oid || '')
            .trim()
            .toLowerCase(),
        upn: String(payload.preferred_username || payload.upn || '').trim(),
        name: String(payload.name || '').trim(),
        payload
    };
}

/**
 * @param {string} raw
 * @param {string[]} audiences
 */
async function verifyWithJwks(raw, audiences) {
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
    err.code = 'invalid_token';
    err.meta = peekTokenMeta(raw);
    err.cause = lastErr;
    throw err;
}

/**
 * @param {string} token
 */
async function validateCallerToken(token) {
    const cfg = getConfig();
    const raw = String(token || '').trim();
    if (!raw) {
        const err = new Error('Anmeldung fehlt.');
        err.status = 401;
        err.code = 'missing_token';
        throw err;
    }

    const payload = await verifyWithJwks(raw, cfg.tokenAudiences);
    if (String(payload.idtyp || '').toLowerCase() === 'app') {
        const err = new Error('App-Only-Token ist hier nicht erlaubt.');
        err.status = 401;
        err.code = 'app_token';
        throw err;
    }
    if (!tokenHasRequiredScope(payload, cfg.requiredScopes)) {
        const err = new Error(
            'Token ohne Election.Access (oder License.Access). Bitte neu anmelden und die API-Berechtigung zulassen.'
        );
        err.status = 401;
        err.code = 'missing_scope';
        err.meta = peekTokenMeta(raw);
        throw err;
    }

    const claims = claimsFromPayload(payload);
    if (!GUID_RE.test(claims.tid) || !GUID_RE.test(claims.oid)) {
        const err = new Error('Token ungültig (fehlende Mandanten-/Benutzer-Claims).');
        err.status = 401;
        err.code = 'invalid_claims';
        throw err;
    }
    if (!issuerMatchesTenant(payload.iss, claims.tid)) {
        const err = new Error('Token ungültig (Issuer passt nicht zum Mandanten).');
        err.status = 401;
        err.code = 'invalid_issuer';
        throw err;
    }
    return claims;
}

module.exports = {
    validateCallerToken,
    tokenHasRequiredScope,
    claimsFromPayload,
    GUID_RE
};
