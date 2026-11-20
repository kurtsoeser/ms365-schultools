/**
 * Lehrer-Freistellung → ms365-calendar-shell (Prefix lfr-cal).
 */
import {
    renderMs365CalendarShell,
    startOfWeekMonday,
    toIsoDateOnly
} from '../../shared/ms365-calendar-shell.js';
import { formatDisplayTextWithDeDates, normalizeStatus } from './lfr-logic.js';

function esc(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function lfrStatusTone(status) {
    const s = normalizeStatus(status).toLowerCase();
    if (s === 'genehmigt') return 'genehmigt';
    if (s === 'abgelehnt') return 'abgelehnt';
    return 'beantragt';
}

function clockPart(iso, fallback) {
    const s = String(iso || '');
    const m = /T(\d{2}):(\d{2})/.exec(s);
    if (!m) return fallback || '00:00:00';
    return m[1] + ':' + m[2] + ':00';
}

function hasClockTime(iso) {
    return /T\d{2}:\d{2}/.test(String(iso || ''));
}

/**
 * @param {object} it
 * @returns {object|null}
 */
export function lfrItemToCalEvent(it) {
    if (!it) return null;
    const startDay = toIsoDateOnly(it.beginn) || '';
    if (!startDay) return null;
    const endDay = toIsoDateOnly(it.ende) || startDay;
    const id = String(it.antragId || it.itemId || '');
    const timed =
        startDay === endDay && hasClockTime(it.beginn) && hasClockTime(it.ende || it.beginn);
    const st = normalizeStatus(it.status);
    const title =
        String(it.lehrerName || '').trim() ||
        formatDisplayTextWithDeDates(String(it.titel || '').trim()) ||
        'Freistellung';
    return {
        id,
        title,
        start: timed ? startDay + 'T' + clockPart(it.beginn) : startDay + 'T00:00:00',
        end: timed
            ? endDay + 'T' + clockPart(it.ende || it.beginn, '23:59:59')
            : endDay + 'T23:59:59',
        allDay: !timed,
        tone: lfrStatusTone(st),
        subtitle: st,
        openAttr: 'data-lfr-detail="' + esc(id) + '"'
    };
}

/**
 * @param {object[]} items
 */
export function buildLfrCalendarEvents(items) {
    return (items || [])
        .filter((it) => normalizeStatus(it && it.status) !== 'Abgelehnt')
        .map(lfrItemToCalEvent)
        .filter(Boolean);
}

/**
 * @param {{
 *   mode?: string,
 *   year?: number,
 *   month?: number,
 *   weekAnchor?: Date|string|number,
 *   events?: object[],
 *   hint?: string|false,
 *   actionsHtml?: string
 * }} opts
 */
export function renderLfrCalendarShell(opts) {
    const o = opts || {};
    const mode = String(o.mode || 'month');
    const year = o.year != null ? Number(o.year) : new Date().getFullYear();
    const month = o.month != null ? Number(o.month) : new Date().getMonth() + 1;
    let weekAnchor =
        o.weekAnchor instanceof Date
            ? o.weekAnchor
            : o.weekAnchor
              ? new Date(o.weekAnchor)
              : new Date();
    if (Number.isNaN(weekAnchor.getTime())) weekAnchor = new Date();
    weekAnchor = startOfWeekMonday(weekAnchor);

    return renderMs365CalendarShell({
        mode,
        events: Array.isArray(o.events) ? o.events : [],
        weekAnchor,
        year,
        month0: month - 1,
        prefix: 'lfr-cal',
        hint:
            o.hint !== undefined
                ? o.hint
                : 'Freistellung antippen für Details · Agenda, Woche und Monat.',
        actionsHtml: o.actionsHtml || ''
    });
}
