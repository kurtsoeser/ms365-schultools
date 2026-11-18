/**
 * Reine Domänenlogik Einsammler.
 */

import {
    ANTWORT_STATUS,
    EINSAMMLER_STATUS,
    MODUS,
    normalizeAntwort,
    normalizeEinsammler,
    toIsoDateOnly
} from './einsammler-model.js';

/**
 * @param {object} draft
 * @returns {{ ok: boolean, errors: string[] }}
 */
export function validateEinsammler(draft) {
    const errors = [];
    const item = normalizeEinsammler(draft);
    if (!item.titel) errors.push('Titel fehlt.');
    if (
        !item.modus ||
        (item.modus !== MODUS.EINWILLIGUNG &&
            item.modus !== MODUS.LESEBESTAETIGUNG &&
            item.modus !== MODUS.ABGABE)
    ) {
        errors.push('Modus ungültig.');
    }
    if (!item.typ) errors.push('Typ fehlt.');
    return { ok: errors.length === 0, errors };
}

/**
 * @param {object} draft
 * @returns {{ ok: boolean, errors: string[] }}
 */
export function validateAntwort(draft) {
    const errors = [];
    const item = normalizeAntwort(draft);
    if (!item.einsammlerId) errors.push('Einsammler-ID fehlt.');
    if (!item.schuelerName) errors.push('Schülername fehlt.');
    return { ok: errors.length === 0, errors };
}

/**
 * Empfängerzeilen aus mehrzeiligem Text (Name[;Klasse][;E-Mail]).
 * @param {string} text
 * @param {string} [defaultKlasse]
 * @returns {Array<{ schuelerName: string, klasseCode: string, elternEmail: string }>}
 */
export function parseEmpfaengerText(text, defaultKlasse) {
    const lines = String(text || '')
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean);
    const defK = String(defaultKlasse || '').trim();
    /** @type {Array<{ schuelerName: string, klasseCode: string, elternEmail: string }>} */
    const out = [];
    for (const line of lines) {
        const parts = line.split(/[;|\t]/).map((p) => p.trim());
        const name = parts[0] || '';
        if (!name) continue;
        let klasse = defK;
        let email = '';
        if (parts.length >= 3) {
            klasse = parts[1] || defK;
            email = parts[2] || '';
        } else if (parts.length === 2) {
            if (parts[1].includes('@')) email = parts[1];
            else klasse = parts[1] || defK;
        }
        out.push({ schuelerName: name, klasseCode: klasse, elternEmail: email });
    }
    return out;
}

/**
 * @param {object} einsammler
 * @param {Array<object>} antworten
 * @returns {{ total: number, offen: number, erledigt: number, zugestimmt: number, abgelehnt: number, gelesen: number, abgelaufen: number, pct: number }}
 */
export function progressStats(einsammler, antworten) {
    const id = String((einsammler && (einsammler.id || einsammler.EinsammlerId)) || '').trim();
    const rows = (antworten || []).filter((a) => a && String(a.einsammlerId || '') === id);
    const total = rows.length;
    let offen = 0;
    let zugestimmt = 0;
    let abgelehnt = 0;
    let gelesen = 0;
    let abgegeben = 0;
    let abgelaufen = 0;
    for (const r of rows) {
        const st = String(r.status || '').toLowerCase();
        if (st === ANTWORT_STATUS.OFFEN) offen++;
        else if (st === ANTWORT_STATUS.ZUGESTIMMT) zugestimmt++;
        else if (st === ANTWORT_STATUS.ABGELEHNT) abgelehnt++;
        else if (st === ANTWORT_STATUS.GELESEN) gelesen++;
        else if (st === ANTWORT_STATUS.ABGEGEBEN) abgegeben++;
        else if (st === ANTWORT_STATUS.ABGELAUFEN) abgelaufen++;
        else offen++;
    }
    const erledigt = total - offen;
    const pct = total ? Math.round((erledigt / total) * 100) : 0;
    return { total, offen, erledigt, zugestimmt, abgelehnt, gelesen, abgegeben, abgelaufen, pct };
}

/**
 * Offene Antworten nach Fälligkeit als abgelaufen markieren (Rückgabe neuer Status).
 * @param {object} einsammler
 * @param {object} antwort
 * @param {Date|string} [now]
 */
export function resolveAntwortStatus(einsammler, antwort, now = new Date()) {
    const st = String((antwort && antwort.status) || ANTWORT_STATUS.OFFEN).toLowerCase();
    if (st !== ANTWORT_STATUS.OFFEN) return st;
    const due = toIsoDateOnly(einsammler && einsammler.faelligAm);
    if (!due) return st;
    const today =
        now instanceof Date && !Number.isNaN(now.getTime())
            ? toIsoDateOnly(now)
            : toIsoDateOnly(now);
    if (today && today > due) return ANTWORT_STATUS.ABGELAUFEN;
    return st;
}

/**
 * Erlaubte manuelle Statuswechsel je Modus.
 * @param {string} modus
 * @returns {string[]}
 */
export function allowedAntwortStatusesForModus(modus) {
    const m = String(modus || '').toLowerCase();
    if (m === MODUS.LESEBESTAETIGUNG) {
        return [ANTWORT_STATUS.OFFEN, ANTWORT_STATUS.GELESEN, ANTWORT_STATUS.ABGELAUFEN];
    }
    if (m === MODUS.ABGABE) {
        return [ANTWORT_STATUS.OFFEN, ANTWORT_STATUS.ABGEGEBEN, ANTWORT_STATUS.ABGELAUFEN];
    }
    return [
        ANTWORT_STATUS.OFFEN,
        ANTWORT_STATUS.ZUGESTIMMT,
        ANTWORT_STATUS.ABGELEHNT,
        ANTWORT_STATUS.ABGELAUFEN
    ];
}

/**
 * Braucht dieser Einsammler einen Eltern-Upload?
 * @param {object} head
 */
export function needsUpload(head) {
    const h = head || {};
    if (String(h.modus || '').toLowerCase() === MODUS.ABGABE) return true;
    return !!(h.uploadErlaubt || h.uploadPflicht);
}

/**
 * @param {string} status
 * @returns {boolean}
 */
export function isEinsammlerEditable(status) {
    const s = String(status || '').toLowerCase();
    return s === EINSAMMLER_STATUS.ENTWURF || s === EINSAMMLER_STATUS.AKTIV;
}
