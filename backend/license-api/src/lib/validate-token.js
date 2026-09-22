'use strict';

const { createRemoteJWKSet, jwtVerify, decodeJwt } = require('jose');

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
function tokenHasLicenseScope(payload) {
    const raw = payload && (payload.scp || payload.scope) ? payload.scp || payload.scope : '';
    return String(raw)
        .split(/[\s,]+/)
        .map((s) => s.trim().toLowerCase())
        .includes('license.access');
}

/**
 * Sichere Diagnose-Claims (kein Token-Inhalt außer Metadaten).
 * @param {string} raw
 */
function peekTokenMeta(raw) {
    try {
        const p = decodeJwt(raw);
        const aud = Array.isArray(p.aud) ? p.aud.join(',') : String(p.aud || '');
        const scp = String(p.scp || p.scope || '');
        return {
            aud: aud.slice(0, 120),
            scp: scp.slice(0, 120),
            tid: String(p.tid || '').slice(0, 40),
            iss: String(p.iss || '').slice(0, 80),
            idtyp: String(p.idtyp || ''),
            hasLicenseScope: tokenHasLicenseScope(p)
        };
    } catch {
        return null;
    }
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
    const name = String(payload.name || '').trim();
    return { tid, oid, upn, name, payload };
}

/**
 * @param {string} raw
 * @param {string[]} audiences
 */
async function verifyWithJwks(raw, audiences) {
    if (!Array.isArray(audiences) || !audiences.length) {
        const err = new Error('Token-Audience ist nicht konfiguriert.');
        err.status = 500;
        err.code = 'audience_config';
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

    const meta = peekTokenMeta(raw);
    let message = 'Token ungültig oder abgelaufen.';
    let code = 'invalid_token';
    const joseCode = lastErr && (lastErr.code || lastErr.claim);
    if (String(joseCode || '').toLowerCase().includes('aud') || /\"aud\"/i.test(String(lastErr && lastErr.message))) {
        message =
            'Falsches Token (Audience). Es wurde vermutlich ein Graph-Token statt License.Access gesendet. Bitte abmelden, Cache leeren, neu anmelden und Consent für License.Access erlauben.';
        code = 'wrong_audience';
    } else if (String(joseCode || '') === 'ERR_JWT_EXPIRED' || /exp/i.test(String(joseCode || ''))) {
        message = 'Token abgelaufen. Bitte neu anmelden.';
        code = 'token_expired';
    } else if (meta && meta.aud && /graph\.microsoft|00000003-0000-0000-c000-000000000000/i.test(meta.aud)) {
        message =
            'Falsches Token: Graph-Token statt License.Access. Beim Laden der Vorlagen muss der Scope License.Access bestätigt werden.';
        code = 'graph_token';
    }

    const err = new Error(message);
    err.status = 401;
    err.code = code;
    err.meta = meta;
    err.cause = lastErr;
    throw err;
}

/**
 * Delegiertes Access-Token der License-API prüfen.
 * Audience ist die License-Backend-App; Graph-Tokens werden abgelehnt.
 * @param {string} token
 * @param {string[]} audiences
 * @returns {Promise<{ tid: string, oid: string, upn: string, name: string, payload: Record<string, unknown> }>}
 */
async function validateCallerToken(token, audiences) {
    const raw = String(token || '').trim();
    if (!raw) {
        const err = new Error('Anmeldung fehlt.');
        err.status = 401;
        err.code = 'missing_token';
        throw err;
    }

    const payload = await verifyWithJwks(raw, audiences);
    if (String(payload.idtyp || '').toLowerCase() === 'app') {
        const err = new Error('App-Only-Token ist hier nicht erlaubt. Bitte mit Benutzerkonto anmelden.');
        err.status = 401;
        err.code = 'app_token';
        throw err;
    }
    if (!tokenHasLicenseScope(payload)) {
        const err = new Error(
            'Token ohne Scope License.Access. Abmelden, neu anmelden und die Berechtigung „Lizenz-API und Katalog lesen“ zulassen.'
        );
        err.status = 401;
        err.code = 'missing_license_scope';
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
    issuerMatchesTenant,
    tokenHasLicenseScope,
    claimsFromPayload,
    peekTokenMeta,
    GUID_RE
};
