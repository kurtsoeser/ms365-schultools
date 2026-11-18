/**
 * Gemeinsame Kalender-Shell: Agenda | Woche | Monat.
 * Events: { id, title, start, end?, allDay?, tone?, subtitle?, openAttr? }
 */

export const MS365_CAL_MODES = Object.freeze([
    { id: 'agenda', label: 'Agenda' },
    { id: 'week', label: 'Woche' },
    { id: 'month', label: 'Monat' }
]);

export const WEEK_GRID_START_HOUR = 7;
export const WEEK_GRID_END_HOUR = 20;

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

export function toIsoDateOnly(value) {
    if (value instanceof Date) {
        if (Number.isNaN(value.getTime())) return '';
        const y = value.getFullYear();
        const m = String(value.getMonth() + 1).padStart(2, '0');
        const d = String(value.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + d;
    }
    const s = String(value || '').trim();
    if (!s) return '';
    const m = /^(\d{4}-\d{2}-\d{2})/.exec(s);
    if (m) return m[1];
    const d = new Date(s);
    if (Number.isNaN(d.getTime())) return '';
    return toIsoDateOnly(d);
}

export function monthLabelDe(year, month0) {
    try {
        return new Intl.DateTimeFormat('de-AT', { month: 'long', year: 'numeric' }).format(
            new Date(year, month0, 1)
        );
    } catch {
        return month0 + 1 + '/' + year;
    }
}

export function startOfWeekMonday(from = new Date()) {
    const d = new Date(from);
    d.setHours(0, 0, 0, 0);
    const day = (d.getDay() + 6) % 7;
    d.setDate(d.getDate() - day);
    return d;
}

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

export function monthGridDays(year, month0) {
    const first = new Date(year, month0, 1);
    const start = startOfWeekMonday(first);
    const days = [];
    for (let i = 0; i < 42; i++) {
        const d = new Date(start);
        d.setDate(start.getDate() + i);
        days.push(d);
    }
    return days;
}

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
        const left = new Intl.DateTimeFormat('de-AT', { day: 'numeric', month: 'short' }).format(a);
        const right = new Intl.DateTimeFormat('de-AT', {
            day: 'numeric',
            month: 'short',
            year: 'numeric'
        }).format(b);
        return left + ' – ' + right;
    } catch {
        return toIsoDateOnly(a) + ' – ' + toIsoDateOnly(b);
    }
}

export function formatTimeHm(iso) {
    const d = iso instanceof Date ? iso : new Date(iso);
    if (Number.isNaN(d.getTime())) return '';
    return String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
}

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
        const rest = new Intl.DateTimeFormat('de-AT', { day: 'numeric', month: 'long' }).format(d);
        return (prefix + wd + ', ' + rest).toUpperCase();
    } catch {
        return prefix + day;
    }
}

function sortEvents(list) {
    return (list || []).slice().sort((a, b) => {
        const sa = new Date(a && a.start).getTime() || 0;
        const sb = new Date(b && b.start).getTime() || 0;
        if (sa !== sb) return sa - sb;
        return String((a && a.title) || '').localeCompare(String((b && b.title) || ''), 'de');
    });
}

export function listForDay(events, dayIso) {
    const day = toIsoDateOnly(dayIso);
    if (!day) return [];
    const from = new Date(day + 'T00:00:00');
    const to = new Date(day + 'T23:59:59.999');
    return sortEvents(events).filter((t) => {
        if (!t || !t.start) return false;
        const s = new Date(t.start);
        const e = new Date(t.end || t.start);
        if (Number.isNaN(s.getTime())) return false;
        return s <= to && e >= from;
    });
}

export function groupAgendaByDay(events, opts = {}) {
    const days = opts.days == null ? 14 : Math.max(1, Number(opts.days) || 14);
    const from = opts.from instanceof Date ? new Date(opts.from) : new Date();
    from.setHours(0, 0, 0, 0);
    const todayIso = toIsoDateOnly(from);
    const tomorrow = new Date(from);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowIso = toIsoDateOnly(tomorrow);
    const end = new Date(from);
    end.setDate(end.getDate() + days);
    end.setHours(23, 59, 59, 999);

    /** @type {Map<string, object[]>} */
    const map = new Map();
    for (let i = 0; i < days; i++) {
        const d = new Date(from);
        d.setDate(d.getDate() + i);
        map.set(toIsoDateOnly(d), []);
    }
    sortEvents(events).forEach((t) => {
        if (!t || !t.start) return;
        const s = new Date(t.start);
        const e = new Date(t.end || t.start);
        if (Number.isNaN(s.getTime())) return;
        if (e < from || s > end) return;
        const day = toIsoDateOnly(s < from ? from : s);
        if (!day || !map.has(day)) return;
        map.get(day).push(t);
    });
    const out = [];
    map.forEach((items, dayIso) => {
        if (!items.length) return;
        out.push({
            dayIso,
            label: formatAgendaDayHeading(dayIso, { todayIso, tomorrowIso }),
            items: sortEvents(items)
        });
    });
    return out;
}

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
    if (Number.isNaN(e.getTime()) || e <= s) e = new Date(s.getTime() + 30 * 60 * 1000);
    if (e < dayStart || s > dayEnd) return null;
    const clampStart = s < dayStart ? dayStart : s;
    const clampEnd = e > dayEnd ? dayEnd : e;
    const startMin = clampStart.getHours() * 60 + clampStart.getMinutes() - startHour * 60;
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

export function layoutTimedEventsForDay(events, dayIso, opts = {}) {
    const timed = (events || [])
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

function toneClass(tone) {
    const t = String(tone || 'akt')
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9_-]/g, '');
    return 'ms365-ev--' + (t || 'akt');
}

function openAttr(ev) {
    if (ev && ev.openAttr) return String(ev.openAttr);
    const id = ev && ev.id != null ? String(ev.id) : '';
    return id ? 'data-ms365-cal-open="' + esc(id) + '"' : '';
}

function weekdayShort(day) {
    try {
        return new Intl.DateTimeFormat('de-AT', { weekday: 'short' }).format(day);
    } catch {
        return ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'][day.getDay()] || '';
    }
}

/**
 * @param {string} mode
 * @param {{ prefix?: string }} [opts] data-attr prefix, default ms365-cal
 */
export function renderCalModeSwitch(mode, opts) {
    const prefix = (opts && opts.prefix) || 'ms365-cal';
    const m = String(mode || 'month');
    return (
        '<div class="ms365-cal__seg" role="group" aria-label="Ansicht">' +
        MS365_CAL_MODES.map((x) => {
            return (
                '<button type="button" class="ms365-cal__seg-btn' +
                (m === x.id ? ' is-active' : '') +
                '" data-' +
                prefix +
                '-mode="' +
                x.id +
                '">' +
                x.label +
                '</button>'
            );
        }).join('') +
        '</div>'
    );
}

/**
 * @param {{ mode: string, title: string, prefix?: string, showNav?: boolean, actionsHtml?: string }} opts
 */
export function renderCalToolbar(opts) {
    const o = opts || {};
    const mode = String(o.mode || 'month');
    const prefix = o.prefix || 'ms365-cal';
    const showNav = o.showNav != null ? !!o.showNav : mode === 'week' || mode === 'month';
    const title = o.title || '';
    return (
        '<div class="ms365-cal__toolbar">' +
        '<div class="ms365-cal__title-block">' +
        (showNav
            ? '<button type="button" class="ms365-cal__textbtn" data-' +
              prefix +
              '-nav="0">Heute</button>' +
              '<button type="button" class="ms365-cal__iconbtn" data-' +
              prefix +
              '-nav="-1" aria-label="Zurück">‹</button>' +
              '<h2 class="ms365-cal__title">' +
              esc(title) +
              '</h2>' +
              '<button type="button" class="ms365-cal__iconbtn" data-' +
              prefix +
              '-nav="1" aria-label="Weiter">›</button>'
            : '<h2 class="ms365-cal__title">' + esc(title) + '</h2>') +
        '</div>' +
        '<div class="ms365-cal__actions">' +
        renderCalModeSwitch(mode, { prefix }) +
        (o.actionsHtml || '') +
        '</div></div>'
    );
}

function agendaTimeLabel(t) {
    if (t.allDay) return 'Ganztägig';
    const a = formatTimeHm(t.start);
    const b = formatTimeHm(t.end || t.start);
    return a && b && a !== b ? a + '–' + b : a || '–';
}

export function renderAgendaView(events, opts) {
    const groups = groupAgendaByDay(events, opts || {});
    if (!groups.length) {
        return '<p class="ms365-cal__hint">Keine Termine in den nächsten Tagen.</p>';
    }
    return (
        '<div class="ms365-agenda">' +
        groups
            .map((g) => {
                return (
                    '<section class="ms365-agenda__day">' +
                    '<h3 class="ms365-agenda__heading">' +
                    esc(g.label) +
                    '</h3>' +
                    '<ul class="ms365-agenda__list">' +
                    g.items
                        .map((t) => {
                            return (
                                '<li><button type="button" class="ms365-agenda__row ' +
                                toneClass(t.tone) +
                                '" ' +
                                openAttr(t) +
                                '>' +
                                '<span class="ms365-agenda__time">' +
                                esc(agendaTimeLabel(t)) +
                                '</span>' +
                                '<span class="ms365-agenda__body">' +
                                '<span class="ms365-agenda__title">' +
                                esc(t.title) +
                                '</span>' +
                                (t.subtitle
                                    ? '<span class="ms365-agenda__source">' + esc(t.subtitle) + '</span>'
                                    : '') +
                                '</span></button></li>'
                            );
                        })
                        .join('') +
                    '</ul></section>'
                );
            })
            .join('') +
        '</div>'
    );
}

export function renderWeekView(events, weekAnchor, opts) {
    const o = opts || {};
    const days = weekDays(weekAnchor);
    const startH = o.startHour == null ? WEEK_GRID_START_HOUR : Number(o.startHour);
    const endH = o.endHour == null ? WEEK_GRID_END_HOUR : Number(o.endHour);
    const hours = [];
    for (let h = startH; h < endH; h++) hours.push(h);
    const todayIso = toIsoDateOnly(new Date());
    const hourCount = Math.max(1, endH - startH);

    const head =
        '<div class="ms365-week__head"><div class="ms365-week__gutter"></div><div class="ms365-week__days">' +
        days
            .map((d) => {
                const iso = toIsoDateOnly(d);
                return (
                    '<div class="ms365-week__dayhead' +
                    (iso === todayIso ? ' is-today' : '') +
                    '"><span class="ms365-week__wd">' +
                    esc(weekdayShort(d)) +
                    '</span><span class="ms365-week__dn">' +
                    d.getDate() +
                    '</span></div>'
                );
            })
            .join('') +
        '</div></div>';

    const allDayRow =
        '<div class="ms365-week__allday"><div class="ms365-week__gutter"><span>Ganztägig</span></div><div class="ms365-week__days">' +
        days
            .map((d) => {
                const iso = toIsoDateOnly(d);
                const allDay = listForDay(events, iso).filter((t) => t.allDay);
                return (
                    '<div class="ms365-week__allday-col">' +
                    allDay
                        .slice(0, 4)
                        .map(
                            (t) =>
                                '<button type="button" class="ms365-week__allday-ev ' +
                                toneClass(t.tone) +
                                '" ' +
                                openAttr(t) +
                                '>' +
                                esc(t.title) +
                                '</button>'
                        )
                        .join('') +
                    (allDay.length > 4
                        ? '<span class="ms365-week__more">+' + (allDay.length - 4) + '</span>'
                        : '') +
                    '</div>'
                );
            })
            .join('') +
        '</div></div>';

    const gutter = hours
        .map(
            (h) =>
                '<div class="ms365-week__hourlabel" style="top:' +
                ((h - startH) / hourCount) * 100 +
                '%">' +
                String(h).padStart(2, '0') +
                ':00</div>'
        )
        .join('');

    const columns = days
        .map((d) => {
            const iso = toIsoDateOnly(d);
            const laid = layoutTimedEventsForDay(listForDay(events, iso), iso, {
                startHour: startH,
                endHour: endH
            });
            const blocks = laid
                .map((item) => {
                    const t = item.termin;
                    const width = 100 / item.cols;
                    const left = item.col * width;
                    const time = formatTimeHm(t.start);
                    return (
                        '<button type="button" class="ms365-week__event ' +
                        toneClass(t.tone) +
                        '" style="top:' +
                        item.topPct.toFixed(2) +
                        '%;height:' +
                        Math.max(item.heightPct, 3.2).toFixed(2) +
                        '%;left:calc(' +
                        left.toFixed(2) +
                        '% + 2px);width:calc(' +
                        width.toFixed(2) +
                        '% - 4px);" ' +
                        openAttr(t) +
                        ' title="' +
                        esc((time ? time + ' · ' : '') + t.title) +
                        '"><span class="ms365-week__event-time">' +
                        esc(time) +
                        '</span><span class="ms365-week__event-title">' +
                        esc(t.title) +
                        '</span></button>'
                    );
                })
                .join('');
            return (
                '<div class="ms365-week__col' +
                (iso === todayIso ? ' is-today' : '') +
                '">' +
                blocks +
                '</div>'
            );
        })
        .join('');

    return (
        '<div class="ms365-week">' +
        head +
        allDayRow +
        '<div class="ms365-week__body"><div class="ms365-week__gutter ms365-week__gutter--times">' +
        gutter +
        '</div><div class="ms365-week__cols">' +
        columns +
        '</div></div></div>'
    );
}

export function renderMonthView(events, year, month0, opts) {
    const days = monthGridDays(year, month0);
    const todayIso = toIsoDateOnly(new Date());
    const max = (opts && opts.maxPerDay) || 3;
    const heads = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So']
        .map((d) => '<div class="ms365-month__head">' + d + '</div>')
        .join('');
    const cells = days
        .map((d) => {
            const iso = toIsoDateOnly(d);
            const inMonth = d.getMonth() === month0;
            const items = listForDay(events, iso);
            const list = items
                .slice(0, max)
                .map((t) => {
                    return (
                        '<button type="button" class="ms365-month__ev ' +
                        toneClass(t.tone) +
                        '" ' +
                        openAttr(t) +
                        '><span>' +
                        esc(t.title) +
                        '</span></button>'
                    );
                })
                .join('');
            const more =
                items.length > max
                    ? '<div class="ms365-month__more">+' + (items.length - max) + ' weitere</div>'
                    : '';
            return (
                '<div class="ms365-month__cell' +
                (inMonth ? '' : ' is-out') +
                (iso === todayIso ? ' is-today' : '') +
                '"><div class="ms365-month__day">' +
                d.getDate() +
                '</div><div class="ms365-month__list">' +
                list +
                more +
                '</div></div>'
            );
        })
        .join('');
    return '<div class="ms365-month">' + heads + cells + '</div>';
}

/**
 * Vollständige Shell (Toolbar + Ansicht).
 * @param {{
 *   mode: 'agenda'|'week'|'month',
 *   events: object[],
 *   weekAnchor?: Date,
 *   year?: number,
 *   month0?: number,
 *   prefix?: string,
 *   actionsHtml?: string,
 *   hint?: string,
 *   extraHtml?: string
 * }} opts
 */
export function renderMs365CalendarShell(opts) {
    const o = opts || {};
    const mode = String(o.mode || 'month');
    const events = Array.isArray(o.events) ? o.events : [];
    const weekAnchor = o.weekAnchor instanceof Date ? o.weekAnchor : startOfWeekMonday(new Date());
    const year = o.year != null ? Number(o.year) : new Date().getFullYear();
    const month0 = o.month0 != null ? Number(o.month0) : new Date().getMonth();
    const title =
        mode === 'agenda'
            ? 'Nächste 14 Tage'
            : mode === 'week'
              ? weekRangeLabelDe(weekAnchor)
              : monthLabelDe(year, month0);
    const body =
        mode === 'agenda'
            ? renderAgendaView(events)
            : mode === 'week'
              ? renderWeekView(events, weekAnchor)
              : renderMonthView(events, year, month0);
    return (
        '<div class="ms365-cal">' +
        renderCalToolbar({
            mode,
            title,
            prefix: o.prefix || 'ms365-cal',
            actionsHtml: o.actionsHtml || ''
        }) +
        (o.extraHtml || '') +
        body +
        (o.hint !== false
            ? '<p class="ms365-cal__hint">' +
              esc(o.hint || 'Termin antippen für Details.') +
              '</p>'
            : '') +
        '</div>'
    );
}
