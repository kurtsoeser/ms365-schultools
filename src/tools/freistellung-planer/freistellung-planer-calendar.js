/**
 * Freistellung → ms365-calendar-shell Events (Prefix fr-cal).
 */
import {
    renderMs365CalendarShell,
    startOfWeekMonday,
    toIsoDateOnly
} from '../../shared/ms365-calendar-shell.js';
import { hasClockTime, toIsoDateTimeLocal } from './freistellung-planer-logic.js';
import { statusLabel } from './freistellung-planer-state.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * Anzeigename für Kalender-Chips: Klasse · Schüler oder Titel.
 * @param {object} it
 */
function frCalTitle(it) {
    const klasse = String((it && it.klasse) || '').trim();
    const name = String((it && (it.schuelerName || it.titel)) || '').trim();
    if (klasse && name) return klasse + ' · ' + name;
    return name || klasse || 'Freistellung';
}

/**
 * Tone für ms365-calendar-shell (CSS: genehmigt / beantragt / abgelehnt).
 * @param {unknown} status
 */
function frStatusTone(status) {
    const s = String(status || '').trim().toLowerCase();
    if (s === 'genehmigt') return 'genehmigt';
    if (s === 'abgelehnt') return 'abgelehnt';
    return 'beantragt';
}

/**
 * Lokale Wanduhr HH:mm:ss aus beginn/ende (datetime-local).
 * @param {unknown} iso
 * @param {string} [fallback]
 */
function clockPart(iso, fallback) {
    const local = toIsoDateTimeLocal(iso);
    if (!local || !local.includes('T')) return fallback || '00:00:00';
    const t = local.split('T')[1] || '';
    const m = /^(\d{2}):(\d{2})/.exec(t);
    if (!m) return fallback || '00:00:00';
    return m[1] + ':' + m[2] + ':00';
}

/**
 * Freistellung-Item → Kalender-Event für die Shared Shell.
 * @param {object} it
 * @returns {object|null}
 */
export function frItemToCalEvent(it) {
    if (!it) return null;
    const startDay = toIsoDateOnly(it.beginn) || '';
    if (!startDay) return null;
    const endDay = toIsoDateOnly(it.ende) || startDay;
    const id = String(it.itemId || '');
    const timed =
        startDay === endDay && hasClockTime(it.beginn) && hasClockTime(it.ende || it.beginn);
    const st = String(it.status || 'Ausstehend').trim();
    const title = frCalTitle(it);
    return {
        id,
        title,
        start: timed ? startDay + 'T' + clockPart(it.beginn) : startDay + 'T00:00:00',
        end: timed
            ? endDay + 'T' + clockPart(it.ende || it.beginn, '23:59:59')
            : endDay + 'T23:59:59',
        allDay: !timed,
        tone: frStatusTone(st),
        subtitle: statusLabel(st),
        openAttr: 'data-fr-detail="' + esc(id) + '"'
    };
}

/**
 * Items → Kalender-Events (Abgelehnt ausgeblendet).
 * @param {object[]} items
 * @returns {object[]}
 */
export function buildFrCalendarEvents(items) {
    return (items || [])
        .filter((it) => String(it && it.status || '').toLowerCase() !== 'abgelehnt')
        .map(frItemToCalEvent)
        .filter(Boolean);
}

/**
 * Shared Kalender-Shell mit Prefix `fr-cal`.
 * @param {{
 *   mode?: string,
 *   year?: number,
 *   month?: number,
 *   weekAnchor?: Date|string|number,
 *   events?: object[],
 *   hint?: string|false,
 *   actionsHtml?: string
 * }} opts
 * `month` ist 1–12 (wie state.calMonth); intern → month0.
 */
export function renderFrCalendarShell(opts) {
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
        prefix: 'fr-cal',
        hint:
            o.hint !== undefined
                ? o.hint
                : 'Freistellung antippen für Details · Agenda, Woche und Monat.',
        actionsHtml: o.actionsHtml || ''
    });
}
