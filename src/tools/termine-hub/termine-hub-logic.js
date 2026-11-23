/**
 * Domänenlogik Kalender und Termine.
 */

import {
    SICHTBARKEIT,
    SYNC_STATUS,
    TERMIN_TYP,
    normalizeTermin,
    toIsoDateOnly,
    toIsoDateTime
} from './termine-hub-model.js';

/**
 * @param {object} draft
 */
export function validateTermin(draft) {
    const errors = [];
    const t = normalizeTermin(draft);
    if (!t.title) errors.push('Titel fehlt.');
    if (!t.start) errors.push('Beginn fehlt.');
    if (!t.end) errors.push('Ende fehlt.');
    if (t.start && t.end && new Date(t.start) > new Date(t.end)) {
        errors.push('Ende liegt vor Beginn.');
    }
    if (!Object.values(SICHTBARKEIT).includes(t.sichtbarkeit)) {
        errors.push('Sichtbarkeit ungültig.');
    }
    if (!Object.values(TERMIN_TYP).includes(t.terminTyp)) {
        errors.push('Termin-Typ ungültig.');
    }
    return { ok: errors.length === 0, errors, item: t };
}

/**
 * @param {Date|string|number} [from]
 * @param {number} [days]
 * @returns {{ from: Date, to: Date }}
 */
export function nextDaysWindow(from = new Date(), days = 14) {
    const start = new Date(from);
    start.setHours(0, 0, 0, 0);
    const end = new Date(start);
    end.setDate(end.getDate() + Number(days || 14));
    end.setHours(23, 59, 59, 999);
    return { from: start, to: end };
}

/**
 * Termine in den nächsten N Tagen (inkl. laufender ganztägiger).
 * @param {object[]} termine
 * @param {{ days?: number, now?: Date }} [opts]
 */
export function listUpcoming(termine, opts = {}) {
    const days = opts.days == null ? 14 : Number(opts.days);
    const { from, to } = nextDaysWindow(opts.now || new Date(), days);
    return sortTermine(termine).filter((t) => {
        if (!t || !t.start) return false;
        const s = new Date(t.start);
        const e = new Date(t.end || t.start);
        if (Number.isNaN(s.getTime())) return false;
        return s <= to && e >= from;
    });
}

/**
 * @param {object[]} termine
 * @param {{ year: number, month: number }} ym month 0-based
 */
export function listForMonth(termine, ym) {
    const year = Number(ym.year);
    const month = Number(ym.month);
    const from = new Date(year, month, 1, 0, 0, 0, 0);
    const to = new Date(year, month + 1, 0, 23, 59, 59, 999);
    return sortTermine(termine).filter((t) => {
        if (!t || !t.start) return false;
        const s = new Date(t.start);
        const e = new Date(t.end || t.start);
        if (Number.isNaN(s.getTime())) return false;
        return s <= to && e >= from;
    });
}

/**
 * @param {object[]} termine
 */
export function sortTermine(termine) {
    return [...(termine || [])].sort((a, b) => {
        const as = String(a && a.start ? a.start : '');
        const bs = String(b && b.start ? b.start : '');
        return as.localeCompare(bs);
    });
}

/**
 * @param {object[]} termine
 */
export function hubStats(termine) {
    const upcoming = listUpcoming(termine, { days: 14 });
    const pending = (termine || []).filter((t) => t && t.syncStatus === SYNC_STATUS.PENDING).length;
    const oeffentlich = upcoming.filter((t) => t.sichtbarkeit === SICHTBARKEIT.OEFFENTLICH).length;
    const intern = upcoming.filter((t) => t.sichtbarkeit === SICHTBARKEIT.INTERN).length;
    return {
        total: (termine || []).length,
        upcoming: upcoming.length,
        oeffentlich,
        intern,
        pending
    };
}

/**
 * @param {object} termin
 * @returns {string}
 */
export function formatTerminRange(termin) {
    const t = normalizeTermin(termin || {});
    if (!t.start) return '–';
    if (t.allDay) {
        const a = toIsoDateOnly(t.start);
        const b = toIsoDateOnly(t.end || t.start);
        return a === b ? formatDeDate(a) + ' (ganztägig)' : formatDeDate(a) + ' – ' + formatDeDate(b) + ' (ganztägig)';
    }
    return formatDeDateTime(t.start) + ' – ' + formatDeDateTime(t.end || t.start);
}

/**
 * @param {string} iso
 */
export function formatDeDate(iso) {
    const s = toIsoDateOnly(iso);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return s || '–';
    const [y, m, d] = s.split('-');
    return d + '.' + m + '.' + y;
}

/**
 * @param {string} iso
 */
export function formatDeDateTime(iso) {
    const raw = toIsoDateTime(iso);
    if (!raw) return '–';
    const d = new Date(raw);
    if (Number.isNaN(d.getTime())) return formatDeDate(raw);
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    const hh = String(d.getHours()).padStart(2, '0');
    const mi = String(d.getMinutes()).padStart(2, '0');
    return dd + '.' + mm + '.' + yyyy + ', ' + hh + ':' + mi;
}

/**
 * datetime-local Wert aus ISO.
 * @param {string} iso
 */
export function toLocalInputValue(iso) {
    const raw = toIsoDateTime(iso);
    if (!raw) return '';
    const d = new Date(raw);
    if (Number.isNaN(d.getTime())) return '';
    const pad = (n) => String(n).padStart(2, '0');
    return (
        d.getFullYear() +
        '-' +
        pad(d.getMonth() + 1) +
        '-' +
        pad(d.getDate()) +
        'T' +
        pad(d.getHours()) +
        ':' +
        pad(d.getMinutes())
    );
}

/**
 * @param {string} localValue datetime-local
 */
export function fromLocalInputValue(localValue) {
    const s = String(localValue || '').trim();
    if (!s) return '';
    const d = new Date(s);
    return Number.isNaN(d.getTime()) ? '' : d.toISOString();
}

/**
 * Monatsname DE.
 * @param {number} year
 * @param {number} month 0-based
 */
export function monthLabelDe(year, month) {
    try {
        return new Intl.DateTimeFormat('de-AT', { month: 'long', year: 'numeric' }).format(
            new Date(year, month, 1)
        );
    } catch {
        return month + 1 + '/' + year;
    }
}

/**
 * Montag 00:00 der Woche (lokal).
 * @param {Date} [from]
 */
export function startOfWeekMonday(from = new Date()) {
    const d = new Date(from);
    d.setHours(0, 0, 0, 0);
    const day = (d.getDay() + 6) % 7;
    d.setDate(d.getDate() - day);
    return d;
}

/**
 * 7 Tage Mo–So ab Wochenbeginn.
 * @param {Date} [anchor]
 * @returns {Date[]}
 */
export function weekDays(anchor = new Date()) {
    const start = startOfWeekMonday(anchor);
    const days = [];
    for (let i = 0; i < 7; i++) {
        const d = new Date(start);
        d.setDate(start.getDate() + i);
        days.push(d);
    }
    return days;
}

/**
 * Monatsraster (Mo-Start, 6 Wochen).
 * @param {number} year
 * @param {number} month 0-based
 * @returns {Date[]}
 */
export function monthGridDays(year, month) {
    const first = new Date(year, month, 1);
    const start = startOfWeekMonday(first);
    const days = [];
    for (let i = 0; i < 42; i++) {
        const d = new Date(start);
        d.setDate(start.getDate() + i);
        days.push(d);
    }
    return days;
}

/**
 * Leertext für Agenda/Woche/Monat – verhindert „lädt nicht“-Eindruck.
 * @param {'agenda'|'week'|'month'|string} mode
 */
export function emptyPeriodLabel(mode) {
    const m = String(mode || '').trim();
    if (m === 'week') return 'Keine Termine in dieser Woche.';
    if (m === 'month') return 'Keine Termine in diesem Monat.';
    return 'Keine Termine in diesem Zeitraum.';
}

/**
 * @param {Date} weekStart
 */
export function weekRangeLabelDe(weekStart) {
    const days = weekDays(weekStart);
    const a = days[0];
    const b = days[6];
    try {
        const sameMonth = a.getMonth() === b.getMonth() && a.getFullYear() === b.getFullYear();
        if (sameMonth) {
            const month = new Intl.DateTimeFormat('de-AT', { month: 'long' }).format(a);
            return a.getDate() + '.–' + b.getDate() + '. ' + month + ' ' + a.getFullYear();
        }
        const left = new Intl.DateTimeFormat('de-AT', {
            day: 'numeric',
            month: 'short'
        }).format(a);
        const right = new Intl.DateTimeFormat('de-AT', {
            day: 'numeric',
            month: 'short',
            year: 'numeric'
        }).format(b);
        return left + ' – ' + right;
    } catch {
        const fmt = (d) =>
            String(d.getDate()).padStart(2, '0') +
            '.' +
            String(d.getMonth() + 1).padStart(2, '0') +
            '.' +
            d.getFullYear();
        return fmt(a) + ' – ' + fmt(b);
    }
}

/**
 * Datum auf 15-Minuten-Raster runden (lokal).
 * @param {Date} [from]
 * @param {{ minutes?: number, addHours?: number }} [opts]
 * @returns {Date}
 */
export function roundToTimeGrid(from = new Date(), opts = {}) {
    const minutes = opts.minutes == null ? 15 : Math.max(1, Number(opts.minutes) || 15);
    const d = new Date(from);
    d.setSeconds(0, 0);
    const total = d.getHours() * 60 + d.getMinutes();
    const rounded = Math.ceil(total / minutes) * minutes;
    d.setHours(Math.floor(rounded / 60), rounded % 60, 0, 0);
    if (opts.addHours) d.setHours(d.getHours() + Number(opts.addHours));
    return d;
}

/**
 * Agenda-Gruppen für kommende Tage.
 * @param {object[]} termine
 * @param {{ days?: number, from?: Date }} [opts]
 * @returns {Array<{ dayIso: string, label: string, items: object[] }>}
 */
export function groupAgendaByDay(termine, opts = {}) {
    const days = opts.days == null ? 14 : Math.max(1, Number(opts.days) || 14);
    const from = opts.from instanceof Date ? new Date(opts.from) : new Date();
    from.setHours(0, 0, 0, 0);
    const todayIso = toIsoDateOnly(from);
    const tomorrow = new Date(from);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowIso = toIsoDateOnly(tomorrow);
    /** @type {Map<string, object[]>} */
    const map = new Map();
    for (let i = 0; i < days; i++) {
        const d = new Date(from);
        d.setDate(d.getDate() + i);
        map.set(toIsoDateOnly(d), []);
    }
    listUpcoming(termine, { days, now: from }).forEach((t) => {
        const day = toIsoDateOnly(t.start);
        if (!day || !map.has(day)) return;
        map.get(day).push(t);
    });
    const out = [];
    map.forEach((items, dayIso) => {
        if (!items.length) return;
        const label = formatAgendaDayHeading(dayIso, { todayIso, tomorrowIso });
        out.push({ dayIso, label, items: sortTermine(items) });
    });
    return out;
}

/**
 * @param {string} dayIso
 * @param {{ todayIso?: string, tomorrowIso?: string }} [opts]
 */
export function formatAgendaDayHeading(dayIso, opts = {}) {
    const day = toIsoDateOnly(dayIso);
    if (!day) return '';
    const d = new Date(day + 'T12:00:00');
    if (Number.isNaN(d.getTime())) return day;
    let prefix = '';
    if (opts.todayIso && day === opts.todayIso) prefix = 'Heute · ';
    else if (opts.tomorrowIso && day === opts.tomorrowIso) prefix = 'Morgen · ';
    try {
        const wd = new Intl.DateTimeFormat('de-AT', { weekday: 'short' })
            .format(d)
            .replace(/\.$/, '');
        const rest = new Intl.DateTimeFormat('de-AT', {
            day: 'numeric',
            month: 'long'
        }).format(d);
        return (prefix + wd + ', ' + rest).toUpperCase();
    } catch {
        return prefix + formatDeDate(day);
    }
}

/**
 * Termine, die einen Kalendertag berühren.
 * @param {object[]} termine
 * @param {string} dayIso YYYY-MM-DD
 */
export function listForDay(termine, dayIso) {
    const day = toIsoDateOnly(dayIso);
    if (!day) return [];
    const from = new Date(day + 'T00:00:00');
    const to = new Date(day + 'T23:59:59.999');
    return sortTermine(termine).filter((t) => {
        if (!t || !t.start) return false;
        const s = new Date(t.start);
        const e = new Date(t.end || t.start);
        if (Number.isNaN(s.getTime())) return false;
        return s <= to && e >= from;
    });
}

/** Standard-Stundenraster für Wochenansicht (lokal). */
export const WEEK_GRID_START_HOUR = 7;
export const WEEK_GRID_END_HOUR = 20;

/**
 * @param {string|Date} iso
 * @returns {{ h: number, m: number }|null}
 */
export function localTimeParts(iso) {
    const d = iso instanceof Date ? iso : new Date(iso);
    if (Number.isNaN(d.getTime())) return null;
    return { h: d.getHours(), m: d.getMinutes() };
}

/**
 * Kurze Uhrzeit HH:MM (lokal).
 * @param {string|Date} iso
 */
export function formatTimeHm(iso) {
    const p = localTimeParts(iso);
    if (!p) return '';
    return String(p.h).padStart(2, '0') + ':' + String(p.m).padStart(2, '0');
}

/**
 * Position eines Timed-Events im Tagesraster (Prozent von oben / Höhe).
 * @param {object} termin
 * @param {string} dayIso
 * @param {{ startHour?: number, endHour?: number }} [opts]
 * @returns {{ topPct: number, heightPct: number, startMin: number, endMin: number }|null}
 */
export function timedEventPlacement(termin, dayIso, opts = {}) {
    if (!termin || termin.allDay) return null;
    const day = toIsoDateOnly(dayIso);
    if (!day) return null;
    const startHour = opts.startHour == null ? WEEK_GRID_START_HOUR : Number(opts.startHour);
    const endHour = opts.endHour == null ? WEEK_GRID_END_HOUR : Number(opts.endHour);
    const spanMin = Math.max(1, (endHour - startHour) * 60);

    const dayStart = new Date(day + 'T00:00:00');
    const dayEnd = new Date(day + 'T23:59:59.999');
    let s = new Date(termin.start);
    let e = new Date(termin.end || termin.start);
    if (Number.isNaN(s.getTime())) return null;
    if (Number.isNaN(e.getTime()) || e <= s) {
        e = new Date(s.getTime() + 30 * 60 * 1000);
    }
    if (e < dayStart || s > dayEnd) return null;

    const clampStart = s < dayStart ? dayStart : s;
    const clampEnd = e > dayEnd ? dayEnd : e;
    const startMin =
        clampStart.getHours() * 60 + clampStart.getMinutes() - startHour * 60;
    const endMin = clampEnd.getHours() * 60 + clampEnd.getMinutes() - startHour * 60;
    const top = Math.max(0, Math.min(spanMin, startMin));
    let bottom = Math.max(0, Math.min(spanMin, endMin));
    if (bottom <= top) bottom = Math.min(spanMin, top + 20);
    return {
        topPct: (top / spanMin) * 100,
        heightPct: ((bottom - top) / spanMin) * 100,
        startMin: top,
        endMin: bottom
    };
}

/**
 * Einfaches Spalten-Layout für überlappende Timed-Events eines Tages.
 * @param {object[]} termine
 * @param {string} dayIso
 * @param {{ startHour?: number, endHour?: number }} [opts]
 * @returns {Array<{ termin: object, topPct: number, heightPct: number, col: number, cols: number }>}
 */
export function layoutTimedEventsForDay(termine, dayIso, opts = {}) {
    const timed = (termine || [])
        .map((t) => {
            const place = timedEventPlacement(t, dayIso, opts);
            if (!place) return null;
            return { termin: t, ...place };
        })
        .filter(Boolean)
        .sort((a, b) => a.startMin - b.startMin || b.endMin - a.endMin);

    /** @type {number[]} */
    const colEnds = [];
    for (const item of timed) {
        let col = 0;
        while (col < colEnds.length && colEnds[col] > item.startMin) col++;
        item.col = col;
        colEnds[col] = item.endMin;
    }
    const cols = Math.max(1, colEnds.length);
    for (const item of timed) item.cols = cols;
    return timed;
}
