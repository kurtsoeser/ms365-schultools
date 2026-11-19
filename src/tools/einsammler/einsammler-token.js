/**
 * Magic-Link Payload für Eltern-Antwortseite (capability URL).
 * Signiert mit HMAC (linkSecret); URL trägt Token im Fragment (#p=).
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
import { loadEinsammlerSettings, saveEinsammlerSettings } from './einsammler-settings.js';

export {
    buildFragmentPublicUrl as buildPublicUrl,
    fromBase64Url,
    readParentLinkToken,
    toBase64Url
};

/**
 * Stellt sicher, dass ein Link-Geheimnis in den Einstellungen liegt.
 * @returns {string}
 */
export function ensureEinsammlerLinkSecret() {
    const cur = loadEinsammlerSettings();
    if (cur.linkSecret) return cur.linkSecret;
    const linkSecret = generateLinkSecret();
    try {
        saveEinsammlerSettings({ linkSecret });
    } catch {
        /* z. B. Node-Tests ohne localStorage */
    }
    return linkSecret;
}

/**
 * Öffentliche Ansicht für einen Empfänger (in den Link eingebettet).
 * @param {{ einsammler: object, antwort: object, schoolName?: string, expDays?: number }} args
 */
export function buildPublicPayload(args) {
    const head = args.einsammler || {};
    const ant = args.antwort || {};
    const days = Math.max(1, Number(args.expDays) || 45);
    const exp = Date.now() + days * 86400000;
    return {
        v: 1,
        eid: String(head.id || '').trim(),
        aid: String(ant.id || '').trim(),
        titel: String(head.titel || '').trim(),
        textInhalt: String(head.textInhalt || '').trim(),
        modus: String(head.modus || 'einwilligung').trim().toLowerCase(),
        typ: String(head.typ || 'sonstiges').trim().toLowerCase(),
        faelligAm: String(head.faelligAm || '').slice(0, 10),
        statusKopf: String(head.status || '').trim().toLowerCase(),
        uploadErlaubt: !!(
            head.uploadErlaubt ||
            head.uploadPflicht ||
            String(head.modus || '').toLowerCase() === 'abgabe'
        ),
        uploadPflicht: !!(
            head.uploadPflicht || String(head.modus || '').toLowerCase() === 'abgabe'
        ),
        schuelerName: String(ant.schuelerName || '').trim(),
        klasseCode: String(ant.klasseCode || '').trim(),
        schoolName: String(args.schoolName || '').trim(),
        exp
    };
}

/**
 * @param {object} payload
 * @returns {{ ok: boolean, error?: string, payload?: object }}
 */
export function validatePublicPayload(payload) {
    if (!payload || typeof payload !== 'object') return { ok: false, error: 'Link ungültig.' };
    if (Number(payload.v) !== 1) return { ok: false, error: 'Unbekannte Link-Version.' };
    if (!payload.eid || !payload.aid) return { ok: false, error: 'Link unvollständig.' };
    if (!payload.titel) return { ok: false, error: 'Kein Titel im Link.' };
    if (payload.exp && Date.now() > Number(payload.exp)) {
        return { ok: false, error: 'Dieser Link ist abgelaufen.' };
    }
    const kopf = String(payload.statusKopf || '').toLowerCase();
    if (kopf && kopf !== 'aktiv' && kopf !== 'entwurf') {
        return { ok: false, error: 'Dieser Einsammler ist nicht mehr offen.' };
    }
    return { ok: true, payload };
}

/**
 * @param {object} payload
 * @param {string} [secret]
 */
export function encodePublicToken(payload, secret) {
    const sec = secret != null ? String(secret) : ensureEinsammlerLinkSecret();
    return signTokenBody(payload, sec);
}

/**
 * @param {string} token
 * @param {{ secret?: string, requireSignature?: boolean }} [opts]
 */
export function decodePublicToken(token, opts = {}) {
    const requireSignature = opts.requireSignature === true;
    const secret =
        opts.secret != null ? String(opts.secret) : loadEinsammlerSettings().linkSecret || '';
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
 * Antwort-Beleg zum Einfügen in der Admin-UI (cross-device ohne API).
 * @param {{ eid: string, aid: string, status: string, kommentar?: string, schuelerName?: string }} draft
 */
export function buildResponseSlip(draft) {
    return {
        v: 1,
        kind: 'einsammler-antwort',
        eid: String(draft.eid || '').trim(),
        aid: String(draft.aid || '').trim(),
        status: String(draft.status || '').trim().toLowerCase(),
        kommentar: String(draft.kommentar || '').trim(),
        schuelerName: String(draft.schuelerName || '').trim(),
        antwortAm: new Date().toISOString(),
        anhangName: String(draft.anhangName || '').trim(),
        anhangMime: String(draft.anhangMime || '').trim(),
        anhangSize: Number(draft.anhangSize) || 0,
        anhangKey: String(draft.anhangKey || draft.aid || '').trim()
    };
}

/**
 * @param {object} slip
 */
export function encodeResponseSlip(slip) {
    return toBase64Url(slip);
}

/**
 * @param {string} raw
 * @returns {{ ok: boolean, error?: string, slip?: object }}
 */
export function decodeResponseSlip(raw) {
    try {
        const data = JSON.parse(fromBase64Url(String(raw || '').trim()));
        if (!data || data.kind !== 'einsammler-antwort' || !data.aid || !data.eid || !data.status) {
            return { ok: false, error: 'Kein gültiger Antwortcode.' };
        }
        return { ok: true, slip: data };
    } catch {
        return { ok: false, error: 'Antwortcode ungültig.' };
    }
}
