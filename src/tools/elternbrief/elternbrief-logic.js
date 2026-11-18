/**
 * Domänenlogik Elternbrief.
 */

import {
    ABSENDER_MODUS,
    LESE_STATUS,
    STATUS,
    normalizeBrief,
    normalizeEmpfaenger,
    toIsoDateOnly
} from './elternbrief-model.js';

/**
 * @param {object} draft
 */
export function validateBrief(draft) {
    const errors = [];
    const b = normalizeBrief(draft);
    if (!b.betreff) errors.push('Betreff fehlt.');
    if (!b.inhalt || b.inhalt.trim().length < 5) errors.push('Inhalt fehlt.');
    if (!Object.values(ABSENDER_MODUS).includes(b.absenderModus)) {
        errors.push('Absender-Modus ungültig.');
    }
    return { ok: errors.length === 0, errors };
}

/**
 * @param {object} brief
 * @param {object[]} empfaenger
 */
export function progressStats(brief, empfaenger) {
    const id = String((brief && brief.id) || '').trim();
    const rows = (empfaenger || []).filter((e) => e && e.briefId === id);
    const total = rows.length;
    let offen = 0;
    let gelesen = 0;
    let abgelaufen = 0;
    for (const r of rows) {
        const st = resolveLeseStatus(brief, r);
        if (st === LESE_STATUS.OFFEN) offen++;
        else if (st === LESE_STATUS.GELESEN) gelesen++;
        else if (st === LESE_STATUS.ABGELAUFEN) abgelaufen++;
        else offen++;
    }
    const erledigt = gelesen;
    const pct = total ? Math.round((erledigt / total) * 100) : 0;
    return { total, offen, gelesen, abgelaufen, erledigt, pct };
}

/**
 * @param {object} brief
 * @param {object} emp
 * @param {Date|string} [now]
 */
export function resolveLeseStatus(brief, emp, now = new Date()) {
    const st = String((emp && emp.status) || LESE_STATUS.OFFEN).toLowerCase();
    if (st === LESE_STATUS.GELESEN) return LESE_STATUS.GELESEN;
    const due = toIsoDateOnly(brief && brief.faelligAm);
    if (!due) return st === LESE_STATUS.ABGELAUFEN ? LESE_STATUS.ABGELAUFEN : LESE_STATUS.OFFEN;
    const today =
        now instanceof Date && !Number.isNaN(now.getTime())
            ? toIsoDateOnly(now)
            : toIsoDateOnly(now);
    if (today && today > due && st === LESE_STATUS.OFFEN) return LESE_STATUS.ABGELAUFEN;
    return st;
}

/**
 * mailto für einen Empfänger (Link im Body).
 * @param {object} brief
 * @param {object} emp
 * @param {string} publicUrl
 */
export function buildMailto(brief, emp, publicUrl) {
    const to = String((emp && emp.elternEmail) || '').trim();
    if (!to) return '';
    const b = normalizeBrief(brief || {});
    const subject = b.betreff || 'Elternbrief';
    const body =
        'Guten Tag,\n\nbitte öffnen Sie den persönlichen Link zum Elternbrief „' +
        subject +
        '“:\n\n' +
        publicUrl +
        '\n\nMit freundlichen Grüßen\n' +
        (b.absenderName || b.erstellerName || '');
    return (
        'mailto:' +
        to +
        '?subject=' +
        encodeURIComponent(subject) +
        '&body=' +
        encodeURIComponent(body)
    );
}

/**
 * @param {string} text
 * @returns {string[]}
 */
export function parseKlasseCodes(text) {
    return String(text || '')
        .split(/[;,\n]+/)
        .map((s) => s.trim())
        .filter(Boolean);
}

/**
 * Empfängerzeilen Name;Klasse;E-Mail
 * @param {string} text
 */
export function parseEmpfaengerText(text) {
    return String(text || '')
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean)
        .map((line) => {
            const parts = line.split(/[;|\t]/).map((p) => p.trim());
            return normalizeEmpfaenger({
                schuelerName: parts[0] || '',
                klasseCode: parts[1] || '',
                elternEmail: parts[2] || (parts[1] && parts[1].includes('@') ? parts[1] : '')
            });
        })
        .filter((e) => e.schuelerName || e.elternEmail);
}

/**
 * Brief aktiv für Eltern-Link?
 * @param {object} brief
 */
export function briefLinkAktiv(brief) {
    const st = String((brief && brief.status) || '').toLowerCase();
    return st === STATUS.GESENDET || st === STATUS.BEREIT;
}
