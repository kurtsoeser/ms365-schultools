/**
 * Portal-Startseite: nächste Termine, ohne die Kalender-Oberfläche zu laden.
 */
import { DASHBOARD_TOOL_LINKS } from '../../shared/dashboard-audience-catalog.js';
import { formatDeDate, formatTimeHm, listUpcoming } from './termine-hub-logic.js';
import { SICHTBARKEIT, SICHTBARKEIT_LABEL, normalizeTermin } from './termine-hub-model.js';
import { SOURCE_IDS } from './termine-hub-settings.js';

export const TERMINE_BRIEFING_DAYS = 14;
export const TERMINE_BRIEFING_MAX = 5;

/**
 * @param {Date|string|number} [now]
 */
export function briefingTodayIso(now) {
    const d = now instanceof Date ? now : now ? new Date(now) : new Date();
    if (Number.isNaN(d.getTime())) return '';
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + day;
}

/**
 * @param {object} termin
 * @param {string} view
 */
function visibleToView(termin, view) {
    if (view !== 'schueler') return true;
    if (String(termin && termin.sourceId ? termin.sourceId : '') === SOURCE_IDS.ME) return true;
    return termin && termin.sichtbarkeit === SICHTBARKEIT.OEFFENTLICH;
}

/**
 * @param {object[]} termine
 * @param {string} view
 */
function prepared(termine, view) {
    return (Array.isArray(termine) ? termine : [])
        .map((row) => normalizeTermin(row))
        .filter((termin) => visibleToView(termin, view));
}

/**
 * @param {string} iso
 */
function shortDate(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
    if (!m) return '';
    return String(Number(m[3])) + '.' + String(Number(m[2])) + '.';
}

/**
 * @param {object[]} termine
 * @param {{ view?: string, now?: Date, max?: number }} [opts]
 */
export function buildTermineBriefing(termine, opts) {
    const o = opts || {};
    const view = o.view === 'schueler' ? 'schueler' : 'lehrer';
    const now = o.now instanceof Date ? o.now : new Date();
    const max = Number(o.max) > 0 ? Number(o.max) : TERMINE_BRIEFING_MAX;
    const visible = listUpcoming(prepared(termine, view), { days: TERMINE_BRIEFING_DAYS, now });
    const lines = visible.slice(0, max).map((t) => {
        const sicht = view === 'schueler' || t.sourceId === SOURCE_IDS.ME ? '' : SICHTBARKEIT_LABEL[t.sichtbarkeit] || '';
        return {
            kicker: t.allDay ? 'ganztägig' : formatTimeHm(t.start) || shortDate(t.start),
            label: t.title || 'Termin',
            meta: [formatDeDate(t.start), sicht].filter(Boolean).join(' · ')
        };
    });
    const n = lines.length;
    return {
        toolId: 'termine-hub',
        title: 'Nächste Termine',
        href: DASHBOARD_TOOL_LINKS['termine-hub'] || 'tools/termine-hub.html',
        lines,
        summary: n === 1 ? '1 Termin' : n ? n + ' Termine' : 'Nächste 14 Tage',
        attention: false,
        emptyText: n ? '' : 'Keine Termine in den nächsten 14 Tagen'
    };
}

/**
 * @param {object[]} termine
 * @param {{ view?: string, now?: Date, max?: number }} [opts]
 */
export function buildTermineHeute(termine, opts) {
    const o = opts || {};
    const view = o.view === 'schueler' ? 'schueler' : 'lehrer';
    const now = o.now instanceof Date ? o.now : new Date();
    const today = briefingTodayIso(now);
    const max = Number(o.max) > 0 ? Number(o.max) : 6;
    const visible = listUpcoming(prepared(termine, view), { days: 1, now }).filter((t) => {
        const start = String(t.start || '').slice(0, 10);
        const end = String(t.end || t.start || '').slice(0, 10);
        return start && start <= today && (end || start) >= today;
    });
    const lines = visible.slice(0, max).map((t) => ({
        kicker: t.allDay ? 'ganztägig' : formatTimeHm(t.start) || '',
        label: t.title || 'Termin',
        meta: t.ort || ''
    }));
    return {
        toolId: 'termine-hub',
        title: 'Heute',
        href: DASHBOARD_TOOL_LINKS['termine-hub'] || 'tools/termine-hub.html',
        lines,
        summary: lines.length === 1 ? '1 Termin' : lines.length ? lines.length + ' Termine' : 'Heute',
        attention: false,
        emptyText: lines.length ? '' : 'Heute nichts im Kalender'
    };
}

/**
 * @param {object[]} termine
 * @param {{ view?: string, now?: Date }} [opts]
 */
export function buildTermineWeek(termine, opts) {
    const o = opts || {};
    const view = o.view === 'schueler' ? 'schueler' : 'lehrer';
    const now = o.now instanceof Date ? o.now : new Date();
    const today = briefingTodayIso(now);
    const monday = new Date(now);
    const weekday = monday.getDay();
    monday.setDate(monday.getDate() + (weekday === 0 ? -6 : 1 - weekday));
    monday.setHours(12, 0, 0, 0);
    const names = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
    const pool = prepared(termine, view);
    const days = [];
    for (let i = 0; i < 5; i += 1) {
        const day = new Date(monday);
        day.setDate(monday.getDate() + i);
        const iso = briefingTodayIso(day);
        const items = pool
            .filter((t) => {
                const start = String(t.start || '').slice(0, 10);
                const end = String(t.end || t.start || '').slice(0, 10);
                return start && start <= iso && (end || start) >= iso;
            })
            .slice(0, 3)
            .map((t) => (t.allDay ? t.title : [formatTimeHm(t.start), t.title].filter(Boolean).join(' ')));
        days.push({
            label: names[day.getDay()] + ' ' + day.getDate() + '.',
            today: iso === today,
            items
        });
    }
    const count = days.reduce((sum, day) => sum + day.items.length, 0);
    return {
        toolId: 'termine-hub',
        title: 'Diese Woche',
        href: DASHBOARD_TOOL_LINKS['termine-hub'] || 'tools/termine-hub.html',
        presentation: 'week',
        days,
        lines: [],
        summary: count === 1 ? '1 Termin' : count ? count + ' Termine' : 'Diese Woche',
        attention: false,
        emptyText: count ? '' : 'Keine Termine in dieser Woche'
    };
}
