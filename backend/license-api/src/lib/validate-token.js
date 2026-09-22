'use strict';

const { createRemoteJWKSet, jwtVerify, decodeJwt } = require('jose');

/** @type {Map<string, ReturnType<typeof createRemoteJWKSet>>} */
const jwksCache = new Map();

function getJwks(url) {
    const key = String(url);
    if (!jwksCache.has(key)) {
        jwksCache.set(key, createRemoteJWKSet(new URL(key)));
    }
    return jwksCache.get(key);
}

const JWKS_URLS = [
    'https://login.microsoftonline.com/common/discovery/v2.0/keys',
    'https://login.microsoftonline.com/common/discovery/keys'
];

/**
 * @param {string} iss
 * @param {string} tid
 */
function issuerMatchesTenant(iss, tid) {
    const issuer = String(iss || '');
    const tenant = String(tid || '').toLowerCase();
    if (!issuer || !tenant) return false;
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
function claimsFromPayload(payload) {
    const tid = String(payload.tid || '').trim();
    const oid = String(payload.oid || payload.sub || '').trim();
    const upn = String(
        payload.preferred_username || payload.upn || payload.unique_name || ''
    ).trim();
    const name = String(payload.name || '').trim();
    return { tid, oid, upn, name, payload };
}

/**
 * @param {string} raw
 * @param {string[]} audiences
 */
async function verifyWithJwks(raw, audiences) {
    const aud = Array.isArray(audiences) && audiences.length ? audiences : undefined;
    let lastErr = null;
    for (let i = 0; i < JWKS_URLS.length; i++) {
        try {
            const verified = await jwtVerify(raw, getJwks(JWKS_URLS[i]), {
                audience: aud,
                clockTolerance: 60
            });
            return verified.payload;
        } catch (e) {
            lastErr = e;
            // Ohne Audience-Check erneut versuchen (Graph-Token-aud-Varianten).
            try {
                const verified = await jwtVerify(raw, getJwks(JWKS_URLS[i]), {
                    clockTolerance: 60
                });
                const payload = verified.payload;
                if (aud && aud.length) {
                    const tokenAud = payload.aud;
                    const list = Array.isArray(tokenAud) ? tokenAud : [tokenAud];
                    const ok = list.some((a) => aud.includes(String(a)));
                    if (!ok) {
                        lastErr = new Error('Token-Audience nicht erlaubt.');
                        continue;
                    }
                }
                return payload;
            } catch (e2) {
                lastErr = e2;
            }
        }
    }
    throw lastErr || new Error('JWT-Signaturprüfung fehlgeschlagen.');
}

/**
 * Access-Token indirekt prüfen: Graph akzeptiert es → Claims aus dem JWT lesen.
 * @param {string} raw
 */
async function validateViaGraph(raw) {
    const res = await fetch(
        'https://graph.microsoft.com/v1.0/me?$select=id,displayName,userPrincipalName',
        {
            headers: {
                Authorization: 'Bearer ' + raw,
                Accept: 'application/json'
            }
        }
    );
    if (!res.ok) {
        const text = await res.text();
        const err = new Error(
            'Graph-Token ungültig (HTTP ' + res.status + '): ' + (text || res.statusText)
        );
        err.status = 401;
        throw err;
    }
    const me = await res.json();
    let payload = {};
    try {
        payload = decodeJwt(raw);
    } catch {
        payload = {};
    }

    let tid = String(payload.tid || '').trim();
    if (!tid) {
        const orgRes = await fetch('https://graph.microsoft.com/v1.0/organization?$select=id', {
            headers: {
                Authorization: 'Bearer ' + raw,
                Accept: 'application/json'
            }
        });
        if (orgRes.ok) {
            const org = await orgRes.json();
            const first = org && org.value && org.value[0];
            tid = String((first && first.id) || '').trim();
        }
    }
    if (!tid) {
        const err = new Error('Tenant-ID konnte aus dem Token nicht gelesen werden.');
        err.status = 401;
        throw err;
    }

    return {
        tid,
        oid: String(payload.oid || me.id || '').trim(),
        upn: String(
            payload.preferred_username ||
                payload.upn ||
                me.userPrincipalName ||
                ''
        ).trim(),
        name: String(payload.name || me.displayName || '').trim(),
        payload
    };
}

/**
 * Delegiertes Entra-/Graph-/ID-Token des Schul-Users prüfen.
 * @param {string} token
 * @param {string[]} audiences
 * @returns {Promise<{ tid: string, oid: string, upn: string, name: string, payload: Record<string, unknown> }>}
 */
async function validateCallerToken(token, audiences) {
    const raw = String(token || '').trim();
    if (!raw) {
        const err = new Error('Authorization Bearer-Token fehlt.');
        err.status = 401;
        throw err;
    }

    let payload = null;
    let verifyErr = null;
    try {
        payload = await verifyWithJwks(raw, audiences);
    } catch (e) {
        verifyErr = e;
    }

    if (payload) {
        const claims = claimsFromPayload(payload);
        if (!claims.tid) {
            const err = new Error('Token enthält keine Tenant-ID (tid).');
            err.status = 401;
            throw err;
        }
        if (!issuerMatchesTenant(payload.iss, claims.tid)) {
            const err = new Error('Token-Issuer passt nicht zur Tenant-ID.');
            err.status = 401;
            throw err;
        }
        return claims;
    }

    // Graph-Access-Tokens scheitern oft an lokaler Signaturprüfung → Fallback.
    try {
        return await validateViaGraph(raw);
    } catch (e) {
        const err = new Error(
            'Token ungültig oder abgelaufen: ' +
                ((verifyErr && verifyErr.message) || '') +
                (verifyErr && e.message ? ' / ' : '') +
                (e.message || String(e))
        );
        err.status = 401;
        throw err;
    }
}

module.exports = {
    validateCallerToken,
    issuerMatchesTenant,
    validateViaGraph
};
