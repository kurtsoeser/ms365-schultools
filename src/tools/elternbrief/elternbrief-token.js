/**
 * Magic-Link für Elternbrief lesen + Lesebestätigung.
 * Signiert mit HMAC (linkSecret); Token im Fragment (#p=).
 */

import {
    buildFragmentPublicUrl,
    fromBase64Url,
    generateLinkSecret,
    readParentLinkToken,
    signTokenBody,
    toBase64Url,
    verifySignedToken
} from '../../shared/parent-link-token.js';
import { loadElternbriefSettings, saveElternbriefSettings } from './elternbrief-settings.js';

export {
    buildFragmentPublicUrl as buildPublicUrl,
    fromBase64Url,
    readParentLinkToken,
    toBase64Url
};

/**
 * @returns {string}
 */
export function ensureElternbriefLinkSecret() {
    const cur = loadElternbriefSettings();
    if (cur.linkSecret) return cur.linkSecret;
    const linkSecret = generateLinkSecret();
    try {
        saveElternbriefSettings({ linkSecret });
    } catch {
        /* z. B. Node-Tests ohne localStorage */
    }
    return linkSecret;
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

/**
 * @param {object} payload
 * @param {string} [secret]
 */
export function encodePublicToken(payload, secret) {
    const sec = secret != null ? String(secret) : ensureElternbriefLinkSecret();
    return signTokenBody(payload, sec);
}

/**
 * @param {string} token
 * @param {{ secret?: string, requireSignature?: boolean }} [opts]
 */
export function decodePublicToken(token, opts = {}) {
    const requireSignature = opts.requireSignature === true;
    const secret =
        opts.secret != null ? String(opts.secret) : loadElternbriefSettings().linkSecret || '';
    const verified = verifySignedToken(token, secret, { requireSignature });
    if (!verified.ok) return { ok: false, error: verified.error };
    try {
        const data = JSON.parse(verified.payloadText || '');
        return validatePublicPayload(data);
    } catch {
        return { ok: false, error: 'Link konnte nicht gelesen werden.' };
    }
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
