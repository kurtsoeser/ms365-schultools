'use strict';

const crypto = require('node:crypto');

function pepper() {
    return String(process.env.STAMMDATEN_PIN_PEPPER || '');
}

function sessionSecret() {
    const configured = String(process.env.STAMMDATEN_SESSION_SECRET || pepper() || '').trim();
    if (configured) return configured;
    const isAzure = !!(process.env.WEBSITE_INSTANCE_ID || process.env.WEBSITE_SITE_NAME);
    if (isAzure) {
        throw new Error(
            'STAMMDATEN_SESSION_SECRET (oder STAMMDATEN_PIN_PEPPER) fehlt in den Azure App Settings.'
        );
    }
    return 'dev-only';
}

async function hashPin(pin, customPepper) {
    const material = `${customPepper != null ? customPepper : pepper()}:${String(pin || '').trim()}`;
    return crypto.createHash('sha256').update(material, 'utf8').digest('hex');
}

function timingSafeEqualHex(a, b) {
    const aa = Buffer.from(String(a || ''), 'utf8');
    const bb = Buffer.from(String(b || ''), 'utf8');
    if (aa.length !== bb.length) return false;
    return crypto.timingSafeEqual(aa, bb);
}

async function verifyPin(pin, expectedHash) {
    if (!expectedHash) return false;
    const actual = await hashPin(pin);
    return timingSafeEqualHex(actual, expectedHash);
}

function createSessionToken(guid) {
    const exp = Date.now() + 60 * 60 * 1000;
    const payload = Buffer.from(JSON.stringify({ g: guid, exp }), 'utf8').toString('base64url');
    const sig = crypto.createHmac('sha256', sessionSecret()).update(payload).digest('base64url');
    return `${payload}.${sig}`;
}

function readSessionToken(token) {
    const raw = String(token || '');
    const [payload, sig] = raw.split('.');
    if (!payload || !sig) return null;
    const expected = crypto.createHmac('sha256', sessionSecret()).update(payload).digest('base64url');
    const a = Buffer.from(sig);
    const b = Buffer.from(expected);
    if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
    try {
        const data = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
        if (!data?.g || !data?.exp || Date.now() > Number(data.exp)) return null;
        return { guid: String(data.g), exp: Number(data.exp) };
    } catch {
        return null;
    }
}

module.exports = {
    hashPin,
    verifyPin,
    createSessionToken,
    readSessionToken
};
