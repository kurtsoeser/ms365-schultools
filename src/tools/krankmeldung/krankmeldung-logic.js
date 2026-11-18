/**
 * Reine Domänenlogik Krankmeldung.
 */

import {
    KATEGORIE,
    MAX_FUTURE_DAYS,
    MAX_PAST_DAYS,
    MAX_RANGE_DAYS,
    STATUS,
    normalizeMeldung,
    toIsoDateOnly
} from './krankmeldung-model.js';

/**
 * @param {string} a
 * @param {string} b
 */
export function daysBetween(a, b) {
    const x = toIsoDateOnly(a);
    const y = toIsoDateOnly(b);
    if (!x || !y) return 0;
    const t0 = Date.parse(x + 'T12:00:00');
    const t1 = Date.parse(y + 'T12:00:00');
    if (Number.isNaN(t0) || Number.isNaN(t1)) return 0;
    return Math.round((t1 - t0) / 86400000);
}

/**
 * @param {object} draft
 * @param {{ today?: string, requireAnhang?: boolean }} [opts]
 * @returns {{ ok: boolean, errors: string[], warnings: string[] }}
 */
export function validateMeldung(draft, opts = {}) {
    const errors = [];
    const warnings = [];
    const item = normalizeMeldung(draft);
    const today = toIsoDateOnly(opts.today || new Date());
    if (!item.schuelerName) errors.push('Schülername fehlt.');
    if (!item.klasseCode) errors.push('Klasse fehlt.');
    if (!item.von) errors.push('Von-Datum fehlt.');
    if (!item.bis) errors.push('Bis-Datum fehlt.');
    if (item.von && item.bis && item.bis < item.von) errors.push('Bis-Datum liegt vor Von-Datum.');
    const kat = item.kategorie;
    if (!Object.values(KATEGORIE).includes(kat)) errors.push('Kategorie ungültig.');

    if (item.von && item.bis) {
        const span = daysBetween(item.von, item.bis);
        if (span > MAX_RANGE_DAYS) {
            errors.push(
                'Zeitraum zu lang (max. ' + MAX_RANGE_DAYS + ' Tage). Bitte bei längerer Abwesenheit die Schule kontaktieren.'
            );
        } else if (span >= 7) {
            warnings.push('Längerer Zeitraum (' + (span + 1) + ' Tage) – Sekretariat prüft ggf. nach.');
        }
    }
    if (item.von && today) {
        const fromToday = daysBetween(today, item.von);
        if (fromToday < -MAX_PAST_DAYS) {
            errors.push(
                'Von-Datum liegt zu weit in der Vergangenheit (max. ' + MAX_PAST_DAYS + ' Tage zurück).'
            );
        }
        if (fromToday > MAX_FUTURE_DAYS) {
            errors.push(
                'Von-Datum liegt zu weit in der Zukunft (max. ' + MAX_FUTURE_DAYS + ' Tage voraus).'
            );
        }
    }
    if (opts.requireAnhang && !item.anhangName && !item.anhangKey && !item.anhangUrl) {
        errors.push('Bitte einen Anhang (z. B. Bestätigung) hochladen.');
    }
    return { ok: errors.length === 0, errors, warnings };
}

/**
 * Überlappung gleicher Name+Klasse mit bestehendem Zeitraum.
 * @param {object} draft
 * @param {object[]} existing
 * @param {string} [ignoreId]
 */
export function findDuplicate(draft, existing, ignoreId) {
    const item = normalizeMeldung(draft);
    const ignore = String(ignoreId || '').trim();
    return (existing || []).find((m) => {
        if (!m || m.id === ignore) return false;
        if (m.status === STATUS.ARCHIVIERT) return false;
        if (m.schuelerName.toLowerCase() !== item.schuelerName.toLowerCase()) return false;
        if (m.klasseCode.toLowerCase() !== item.klasseCode.toLowerCase()) return false;
        // Datumsüberlappung
        return item.von <= m.bis && item.bis >= m.von;
    });
}

/**
 * @param {object[]} meldungen
 * @param {{
 *   day?: string,
 *   klasse?: string,
 *   status?: string,
 *   weekOf?: string,
 *   kategorie?: string,
 *   q?: string,
 *   klassenAllow?: string[]
 * }} filters
 */
export function filterMeldungen(meldungen, filters = {}) {
    const day = toIsoDateOnly(filters.day);
    const weekOf = toIsoDateOnly(filters.weekOf);
    const klasse = String(filters.klasse || '')
        .trim()
        .toLowerCase();
    const status = String(filters.status || '')
        .trim()
        .toLowerCase();
    const kategorie = String(filters.kategorie || '')
        .trim()
        .toLowerCase();
    const q = String(filters.q || '')
        .trim()
        .toLowerCase();
    const allow = Array.isArray(filters.klassenAllow)
        ? filters.klassenAllow.map((c) => String(c || '').trim().toLowerCase()).filter(Boolean)
        : null;
    const allowSet = allow && allow.length ? new Set(allow) : null;

    let weekStart = '';
    let weekEnd = '';
    if (weekOf) {
        const d = new Date(weekOf + 'T12:00:00');
        const dow = d.getDay(); // 0 So
        const mondayOffset = dow === 0 ? -6 : 1 - dow;
        const mon = new Date(d);
        mon.setDate(d.getDate() + mondayOffset);
        const sun = new Date(mon);
        sun.setDate(mon.getDate() + 6);
        weekStart = toIsoDateOnly(mon);
        weekEnd = toIsoDateOnly(sun);
    }

    return (meldungen || []).filter((m) => {
        if (!m) return false;
        if (klasse && String(m.klasseCode || '').toLowerCase() !== klasse) return false;
        if (allowSet && !allowSet.has(String(m.klasseCode || '').toLowerCase())) return false;
        if (status && String(m.status || '').toLowerCase() !== status) return false;
        if (kategorie && String(m.kategorie || '').toLowerCase() !== kategorie) return false;
        if (q) {
            const hay =
                String(m.schuelerName || '').toLowerCase() +
                ' ' +
                String(m.klasseCode || '').toLowerCase() +
                ' ' +
                String(m.notiz || '').toLowerCase() +
                ' ' +
                String(m.gemeldetVonName || '').toLowerCase();
            if (!hay.includes(q)) return false;
        }
        if (day) {
            if (!(m.von <= day && m.bis >= day)) return false;
        }
        if (weekStart && weekEnd) {
            if (!(m.von <= weekEnd && m.bis >= weekStart)) return false;
        }
        return true;
    });
}

/**
 * CSV für Inbox-Export.
 * @param {object[]} meldungen
 */
export function meldungenToCsv(meldungen) {
    const header = [
        'Id',
        'SchuelerName',
        'KlasseCode',
        'Von',
        'Bis',
        'Kategorie',
        'Status',
        'GemeldetAm',
        'GemeldetVon',
        'Notiz',
        'AnhangName'
    ];
    const lines = [header.join(';')];
    for (const m of meldungen || []) {
        if (!m) continue;
        const row = [
            m.id,
            m.schuelerName,
            m.klasseCode,
            m.von,
            m.bis,
            m.kategorie,
            m.status,
            m.gemeldetAm,
            m.gemeldetVonName || m.gemeldetVonEmail || '',
            m.notiz,
            m.anhangName || ''
        ].map((v) => {
            const s = String(v ?? '');
            return /[;"\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
        });
        lines.push(row.join(';'));
    }
    return lines.join('\n');
}

/**
 * @param {object[]} meldungen
 * @param {string} [day]
 */
export function inboxStats(meldungen, day) {
    const today = day || toIsoDateOnly(new Date());
    const list = meldungen || [];
    const neu = list.filter((m) => m.status === STATUS.NEU).length;
    const heute = filterMeldungen(list, { day: today }).length;
    const offenHeute = filterMeldungen(list, { day: today, status: STATUS.NEU }).length;
    return { neu, heute, offenHeute, total: list.length };
}
