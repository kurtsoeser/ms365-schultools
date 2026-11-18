/**
 * Magic-Link für Elternbrief lesen + Lesebestätigung.
 */

/**
 * @param {unknown} value
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
 * @param {{ brief: object, empfaenger: object, schoolName?: string, expDays?: number, hasAnhang?: boolean }} args
 */
export function buildPublicPayload(args) {
    const b = args.brief || {};
    const e = args.empfaenger || {};
    const days = Math.max(1, Number(args.expDays) || 45);
    return {
        v: 1,
        kind: 'elternbrief',
        bid: String(b.id || '').trim(),
        eid: String(e.id || '').trim(),
        betreff: String(b.betreff || '').trim(),
        inhalt: String(b.inhalt || '').trim(),
        schuelerName: String(e.schuelerName || '').trim(),
        klasseCode: String(e.klasseCode || '').trim(),
        lesebestaetigung: b.lesebestaetigung !== false,
        hasAnhang: !!(args.hasAnhang || b.anhangName || b.anhangKey),
        anhangName: String(b.anhangName || '').trim(),
        absenderName: String(b.absenderName || b.erstellerName || '').trim(),
        schoolName: String(args.schoolName || '').trim(),
        statusKopf: String(b.status || '').trim().toLowerCase(),
        exp: Date.now() + days * 86400000
    };
}

/**
 * @param {object} payload
 */
export function validatePublicPayload(payload) {
    if (!payload || payload.kind !== 'elternbrief') return { ok: false, error: 'Ungültiger Link.' };
    if (!payload.bid || !payload.eid || !payload.betreff) {
        return { ok: false, error: 'Link unvollständig.' };
    }
    if (payload.exp && Date.now() > Number(payload.exp)) {
        return { ok: false, error: 'Link abgelaufen.' };
    }
    const st = String(payload.statusKopf || '').toLowerCase();
    if (st && st !== 'gesendet' && st !== 'bereit') {
        return { ok: false, error: 'Dieser Elternbrief ist nicht mehr offen.' };
    }
    return { ok: true, payload };
}

export function encodePublicToken(payload) {
    return toBase64Url(payload);
}

export function decodePublicToken(token) {
    try {
        const data = JSON.parse(fromBase64Url(String(token || '').trim()));
        return validatePublicPayload(data);
    } catch {
        return { ok: false, error: 'Link konnte nicht gelesen werden.' };
    }
}

/**
 * @param {string} baseHref
 * @param {string} token
 */
export function buildPublicUrl(baseHref, token) {
    const base = String(baseHref || 'elternbrief-lesen.html').trim() || 'elternbrief-lesen.html';
    const join = base.includes('?') ? '&' : '?';
    return base + join + 'p=' + encodeURIComponent(token);
}

/**
 * Antwortcode Lesebestätigung.
 * @param {{ bid: string, eid: string }} draft
 */
export function buildLeseSlip(draft) {
    return {
        v: 1,
        kind: 'elternbrief-gelesen',
        bid: String(draft.bid || '').trim(),
        eid: String(draft.eid || '').trim(),
        gelesenAm: new Date().toISOString()
    };
}

export function encodeLeseSlip(slip) {
    return toBase64Url(slip);
}

export function decodeLeseSlip(raw) {
    try {
        const data = JSON.parse(fromBase64Url(String(raw || '').trim()));
        if (!data || data.kind !== 'elternbrief-gelesen' || !data.bid || !data.eid) {
            return { ok: false, error: 'Kein gültiger Bestätigungscode.' };
        }
        return { ok: true, slip: data };
    } catch {
        return { ok: false, error: 'Code ungültig.' };
    }
}
