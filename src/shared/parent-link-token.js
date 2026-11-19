/**
 * Signierte Eltern-Magic-Links (HMAC-SHA256) + Fragment-URLs.
 * Token-Format: base64url(JSON).base64url(HMAC)
 * Ohne gültige Signatur lehnt die Remote-API ab (R2 / a5).
 */

/**
 * @param {unknown} value
 * @returns {string}
 */
export function toBase64Url(value) {
    const json = typeof value === 'string' ? value : JSON.stringify(value);
    if (typeof Buffer !== 'undefined') {
        return Buffer.from(json, 'utf8').toString('base64url');
    }
    const bytes = new TextEncoder().encode(json);
    let bin = '';
    for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

/**
 * @param {string} raw
 * @returns {string}
 */
export function fromBase64Url(raw) {
    const s = String(raw || '')
        .trim()
        .replace(/-/g, '+')
        .replace(/_/g, '/');
    const pad = s.length % 4 === 0 ? '' : '='.repeat(4 - (s.length % 4));
    if (typeof Buffer !== 'undefined') {
        return Buffer.from(s + pad, 'base64').toString('utf8');
    }
    const bin = atob(s + pad);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return new TextDecoder().decode(bytes);
}

/**
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function bytesToBase64Url(bytes) {
    if (typeof Buffer !== 'undefined') {
        return Buffer.from(bytes).toString('base64url');
    }
    let bin = '';
    for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

/**
 * @param {Uint8Array} a
 * @param {Uint8Array} b
 */
function concatBytes(a, b) {
    const out = new Uint8Array(a.length + b.length);
    out.set(a, 0);
    out.set(b, a.length);
    return out;
}

/**
 * @param {Uint8Array} data
 * @returns {Uint8Array}
 */
function sha256Bytes(data) {
    const K = new Uint32Array([
        0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4,
        0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe,
        0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f,
        0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7,
        0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc,
        0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b,
        0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116,
        0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
        0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7,
        0xc67178f2
    ]);
    const bitLen = data.length * 8;
    let total = data.length + 1 + 8;
    const mod = total % 64;
    if (mod !== 0) total += 64 - mod;
    const buf = new Uint8Array(total);
    buf.set(data);
    buf[data.length] = 0x80;
    const view = new DataView(buf.buffer);
    view.setUint32(total - 4, bitLen >>> 0, false);
    view.setUint32(total - 8, Math.floor(bitLen / 0x100000000), false);

    let h0 = 0x6a09e667;
    let h1 = 0xbb67ae85;
    let h2 = 0x3c6ef372;
    let h3 = 0xa54ff53a;
    let h4 = 0x510e527f;
    let h5 = 0x9b05688c;
    let h6 = 0x1f83d9ab;
    let h7 = 0x5be0cd19;

    const w = new Uint32Array(64);
    for (let i = 0; i < total; i += 64) {
        for (let j = 0; j < 16; j++) w[j] = view.getUint32(i + j * 4, false);
        for (let j = 16; j < 64; j++) {
            const v1 = w[j - 15];
            const s0 = ((v1 >>> 7) | (v1 << 25)) ^ ((v1 >>> 18) | (v1 << 14)) ^ (v1 >>> 3);
            const v2 = w[j - 2];
            const s1 = ((v2 >>> 17) | (v2 << 15)) ^ ((v2 >>> 19) | (v2 << 13)) ^ (v2 >>> 10);
            w[j] = (w[j - 16] + s0 + w[j - 7] + s1) >>> 0;
        }
        let a = h0;
        let b = h1;
        let c = h2;
        let d = h3;
        let e = h4;
        let f = h5;
        let g = h6;
        let h = h7;
        for (let j = 0; j < 64; j++) {
            const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
            const ch = (e & f) ^ (~e & g);
            const temp1 = (h + S1 + ch + K[j] + w[j]) >>> 0;
            const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
            const maj = (a & b) ^ (a & c) ^ (b & c);
            const temp2 = (S0 + maj) >>> 0;
            h = g;
            g = f;
            f = e;
            e = (d + temp1) >>> 0;
            d = c;
            c = b;
            b = a;
            a = (temp1 + temp2) >>> 0;
        }
        h0 = (h0 + a) >>> 0;
        h1 = (h1 + b) >>> 0;
        h2 = (h2 + c) >>> 0;
        h3 = (h3 + d) >>> 0;
        h4 = (h4 + e) >>> 0;
        h5 = (h5 + f) >>> 0;
        h6 = (h6 + g) >>> 0;
        h7 = (h7 + h) >>> 0;
    }
    const out = new Uint8Array(32);
    const ov = new DataView(out.buffer);
    ov.setUint32(0, h0, false);
    ov.setUint32(4, h1, false);
    ov.setUint32(8, h2, false);
    ov.setUint32(12, h3, false);
    ov.setUint32(16, h4, false);
    ov.setUint32(20, h5, false);
    ov.setUint32(24, h6, false);
    ov.setUint32(28, h7, false);
    return out;
}

/**
 * @param {string} secret
 * @param {string} message
 * @returns {string}
 */
export function hmacSha256Base64Url(secret, message) {
    const enc = new TextEncoder();
    const keyBytes = enc.encode(String(secret || ''));
    const msgBytes = enc.encode(String(message || ''));
    const block = 64;
    const key = keyBytes.length > block ? sha256Bytes(keyBytes) : keyBytes;
    const k = new Uint8Array(block);
    k.set(key);
    const oPad = new Uint8Array(block);
    const iPad = new Uint8Array(block);
    for (let i = 0; i < block; i++) {
        oPad[i] = k[i] ^ 0x5c;
        iPad[i] = k[i] ^ 0x36;
    }
    const inner = sha256Bytes(concatBytes(iPad, msgBytes));
    return bytesToBase64Url(sha256Bytes(concatBytes(oPad, inner)));
}

/**
 * @param {string} a
 * @param {string} b
 */
export function timingSafeEqualStr(a, b) {
    const aa = String(a || '');
    const bb = String(b || '');
    if (aa.length !== bb.length) return false;
    let diff = 0;
    for (let i = 0; i < aa.length; i++) diff |= aa.charCodeAt(i) ^ bb.charCodeAt(i);
    return diff === 0;
}

/**
 * @param {string} token
 * @returns {{ body: string, sig: string }}
 */
export function splitSignedToken(token) {
    const raw = String(token || '').trim();
    const i = raw.lastIndexOf('.');
    if (i <= 0 || i === raw.length - 1) return { body: raw, sig: '' };
    return { body: raw.slice(0, i), sig: raw.slice(i + 1) };
}

/**
 * @param {object | string} payload
 * @param {string} secret
 * @returns {string}
 */
export function signTokenBody(payload, secret) {
    const body = toBase64Url(payload);
    const sec = String(secret || '').trim();
    if (!sec) return body;
    return body + '.' + hmacSha256Base64Url(sec, body);
}

/**
 * @param {string} token
 * @param {string} secret
 * @param {{ requireSignature?: boolean }} [opts]
 * @returns {{ ok: boolean, error?: string, body?: string, payloadText?: string }}
 */
export function verifySignedToken(token, secret, opts = {}) {
    const requireSignature = opts.requireSignature !== false;
    const { body, sig } = splitSignedToken(token);
    if (!body) return { ok: false, error: 'Link ungültig.' };
    const sec = String(secret || '').trim();
    if (requireSignature) {
        if (!sec) return { ok: false, error: 'Link-Geheimnis fehlt – Remote-API nicht freigeschaltet.' };
        if (!sig) return { ok: false, error: 'Link ohne Signatur.' };
        const expected = hmacSha256Base64Url(sec, body);
        if (!timingSafeEqualStr(sig, expected)) {
            return { ok: false, error: 'Link-Signatur ungültig.' };
        }
    } else if (sec && sig) {
        const expected = hmacSha256Base64Url(sec, body);
        if (!timingSafeEqualStr(sig, expected)) {
            return { ok: false, error: 'Link-Signatur ungültig.' };
        }
    }
    try {
        return { ok: true, body, payloadText: fromBase64Url(body) };
    } catch {
        return { ok: false, error: 'Link konnte nicht gelesen werden.' };
    }
}

/**
 * @param {string} [baseHref]
 * @param {string} token
 * @returns {string} URL with token in the fragment (#p=…)
 */
export function buildFragmentPublicUrl(baseHref, token) {
    const base = String(baseHref || '').trim() || 'index.html';
    const hashIdx = base.indexOf('#');
    const withoutHash = hashIdx >= 0 ? base.slice(0, hashIdx) : base;
    return withoutHash + '#p=' + encodeURIComponent(token);
}

/**
 * Liest Token aus Fragment (#p=) oder Legacy-Query (?p=).
 * @param {{ search?: string, hash?: string } | string | null | undefined} [loc]
 * @returns {string}
 */
export function readParentLinkToken(loc) {
    let search = '';
    let hash = '';
    if (typeof loc === 'string') {
        try {
            const u = new URL(loc, 'https://local.invalid');
            search = u.search;
            hash = u.hash;
        } catch {
            const h = loc.indexOf('#');
            const q = loc.indexOf('?');
            if (h >= 0) hash = loc.slice(h);
            if (q >= 0) search = loc.slice(q, h >= 0 ? h : undefined);
        }
    } else if (loc && typeof loc === 'object') {
        search = String(loc.search || '');
        hash = String(loc.hash || '');
    } else if (typeof window !== 'undefined' && window.location) {
        search = window.location.search || '';
        hash = window.location.hash || '';
    }
    if (hash) {
        const raw = hash.charAt(0) === '#' ? hash.slice(1) : hash;
        const hp = new URLSearchParams(raw);
        const fromHash = String(hp.get('p') || hp.get('token') || '').trim();
        if (fromHash) return fromHash;
    }
    const qp = new URLSearchParams(search.charAt(0) === '?' ? search.slice(1) : search);
    return String(qp.get('p') || qp.get('token') || '').trim();
}

/**
 * Zufälliges Link-Geheimnis.
 * @returns {string}
 */
export function generateLinkSecret() {
    const bytes = new Uint8Array(32);
    if (typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function') {
        crypto.getRandomValues(bytes);
    } else {
        for (let i = 0; i < 32; i++) bytes[i] = Math.floor(Math.random() * 256);
    }
    return bytesToBase64Url(bytes);
}
