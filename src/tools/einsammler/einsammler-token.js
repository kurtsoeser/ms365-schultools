/**
 * Magic-Link Payload für Eltern-Antwortseite (capability URL).
 * Lokal: base64url JSON (Link = Zugang). Remote: gleiche Payload + optional HMAC.
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
 */
export function encodePublicToken(payload) {
    return toBase64Url(payload);
}

/**
 * @param {string} token
 */
export function decodePublicToken(token) {
    try {
        const raw = fromBase64Url(String(token || '').trim());
        const data = JSON.parse(raw);
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

/**
 * @param {string} [baseHref] absolute or relative path to antwort page
 * @param {string} token
 */
export function buildPublicUrl(baseHref, token) {
    const base = String(baseHref || 'einsammler-antwort.html').trim() || 'einsammler-antwort.html';
    const join = base.includes('?') ? '&' : '?';
    return base + join + 'p=' + encodeURIComponent(token);
}
