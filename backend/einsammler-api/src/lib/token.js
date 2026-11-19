'use strict';

const crypto = require('node:crypto');

function fromBase64Url(raw) {
    const s = String(raw || '')
        .trim()
        .replace(/-/g, '+')
        .replace(/_/g, '/');
    const pad = s.length % 4 === 0 ? '' : '='.repeat(4 - (s.length % 4));
    return Buffer.from(s + pad, 'base64').toString('utf8');
}

function linkSecret() {
    return String(process.env.EINSAMMLER_LINK_SECRET || '').trim();
}

/**
 * Remote-Elternkanal nur mit konfiguriertem HMAC-Geheimnis (w2/a5).
 * @returns {boolean}
 */
function publicApiEnabled() {
    return !!linkSecret();
}

/**
 * @param {string} a
 * @param {string} b
 */
function timingSafeEqualStr(a, b) {
    const aa = Buffer.from(String(a || ''), 'utf8');
    const bb = Buffer.from(String(b || ''), 'utf8');
    if (aa.length !== bb.length) return false;
    return crypto.timingSafeEqual(aa, bb);
}

/**
 * @param {string} token
 * @returns {{ ok: boolean, error?: string, payload?: object }}
 */
function decodePublicToken(token) {
    const secret = linkSecret();
    if (!secret) {
        return {
            ok: false,
            error: 'Remote-API nicht freigeschaltet (EINSAMMLER_LINK_SECRET fehlt).'
        };
    }
    const raw = String(token || '').trim();
    const dot = raw.lastIndexOf('.');
    if (dot <= 0 || dot === raw.length - 1) {
        return { ok: false, error: 'Link ohne Signatur.' };
    }
    const body = raw.slice(0, dot);
    const sig = raw.slice(dot + 1);
    const expected = crypto.createHmac('sha256', secret).update(body, 'utf8').digest('base64url');
    if (!timingSafeEqualStr(sig, expected)) {
        return { ok: false, error: 'Link-Signatur ungültig.' };
    }
    try {
        const data = JSON.parse(fromBase64Url(body));
        if (!data || Number(data.v) !== 1 || !data.eid || !data.aid || !data.titel) {
            return { ok: false, error: 'Link ungültig.' };
        }
        if (data.exp && Date.now() > Number(data.exp)) {
            return { ok: false, error: 'Link abgelaufen.' };
        }
        return { ok: true, payload: data };
    } catch {
        return { ok: false, error: 'Link konnte nicht gelesen werden.' };
    }
}

module.exports = { decodePublicToken, publicApiEnabled, linkSecret };
