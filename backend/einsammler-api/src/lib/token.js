'use strict';

function fromBase64Url(raw) {
    const s = String(raw || '')
        .trim()
        .replace(/-/g, '+')
        .replace(/_/g, '/');
    const pad = s.length % 4 === 0 ? '' : '='.repeat(4 - (s.length % 4));
    return Buffer.from(s + pad, 'base64').toString('utf8');
}

/**
 * @param {string} token
 * @returns {{ ok: boolean, error?: string, payload?: object }}
 */
function decodePublicToken(token) {
    try {
        const data = JSON.parse(fromBase64Url(token));
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

module.exports = { decodePublicToken };
