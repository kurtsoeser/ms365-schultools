/**
 * Export: iCal (.ics) für Schularbeiten.
 */
import { formatDeDate, labelMaps, statusLabel } from './schularbeiten-planer-state.js';
import {
    schularbeitCalendarSubject,
    schularbeitTerminZeitfenster,
    formatSchularbeitZeitspanne,
    SCHULARBEIT_CALENDAR_TZ,
    icsViennaTimezoneBlock
} from './schularbeiten-planer-logic.js';

/**
 * @param {object[]} items
 * @param {object} stammdaten
 * @param {string} [calendarName]
 */
export function buildIcs(items, stammdaten, calendarName) {
    const labels = labelMaps(stammdaten || {});
    const list = Array.isArray(items) ? items : [];
    const tz = SCHULARBEIT_CALENDAR_TZ;
    const lines = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//MS365schule//Schularbeiten-Planer//DE',
        'CALSCALE:GREGORIAN',
        'METHOD:PUBLISH',
        'X-WR-TIMEZONE:' + tz,
        'X-WR-CALNAME:' + escapeIcsText(calendarName || 'Schularbeiten'),
        icsViennaTimezoneBlock()
    ];

    list.forEach((sa) => {
        const day = String(sa.datum || '').replace(/-/g, '');
        if (!/^\d{8}$/.test(day)) return;
        const uid = (sa.schularbeitId || sa.itemId || day) + '@schularbeiten.ms365schule';
        const summary = schularbeitCalendarSubject(sa, {
            fach: labels.fach[sa.fachCode] || sa.fachCode,
            klasse: labels.klasse[sa.klasseCode] || sa.klasseCode
        });
        const slot = schularbeitTerminZeitfenster(sa);
        const zeit = formatSchularbeitZeitspanne(sa);
        const desc = [
            zeit ? 'Zeit: ' + zeit : '',
            'Status: ' + statusLabel(sa.status),
            'Dauer: ' + (sa.dauerMinuten || '') + ' Min.',
            sa.notiz ? 'Notiz: ' + sa.notiz : ''
        ]
            .filter(Boolean)
            .join('\\n');

        lines.push('BEGIN:VEVENT');
        lines.push('UID:' + uid);
        lines.push('DTSTAMP:' + utcStamp());
        if (slot && !slot.isAllDay && slot.startDateTime && slot.endDateTime) {
            lines.push('DTSTART;TZID=' + tz + ':' + toIcsLocalDateTime(slot.startDateTime));
            lines.push('DTEND;TZID=' + tz + ':' + toIcsLocalDateTime(slot.endDateTime));
        } else {
            lines.push('DTSTART;VALUE=DATE:' + day);
            lines.push('DTEND;VALUE=DATE:' + nextDayCompact(day));
        }
        lines.push('SUMMARY:' + escapeIcsText(summary));
        if (desc) lines.push('DESCRIPTION:' + escapeIcsText(desc));
        lines.push('END:VEVENT');
    });

    lines.push('END:VCALENDAR');
    return lines.join('\r\n');
}

function toIcsLocalDateTime(isoLocal) {
    const s = String(isoLocal || '').trim();
    const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(s);
    if (!m) return '';
    return m[1] + m[2] + m[3] + 'T' + m[4] + m[5] + '00';
}

function escapeIcsText(s) {
    return String(s || '')
        .replace(/\\/g, '\\\\')
        .replace(/;/g, '\\;')
        .replace(/,/g, '\\,')
        .replace(/\n/g, '\\n');
}

function utcStamp() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    return (
        d.getUTCFullYear() +
        p(d.getUTCMonth() + 1) +
        p(d.getUTCDate()) +
        'T' +
        p(d.getUTCHours()) +
        p(d.getUTCMinutes()) +
        p(d.getUTCSeconds()) +
        'Z'
    );
}

function nextDayCompact(yyyymmdd) {
    const y = Number(yyyymmdd.slice(0, 4));
    const m = Number(yyyymmdd.slice(4, 6));
    const d = Number(yyyymmdd.slice(6, 8));
    const dt = new Date(Date.UTC(y, m - 1, d + 1));
    const p = (n) => String(n).padStart(2, '0');
    return dt.getUTCFullYear() + p(dt.getUTCMonth() + 1) + p(dt.getUTCDate());
}

/**
 * @param {string} ics
 * @param {string} [filename]
 */
export function downloadIcs(ics, filename) {
    const blob = new Blob([ics], { type: 'text/calendar;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename || 'schularbeiten.ics';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
}

export function printTableCaption(items) {
    return `${(items || []).length} Termin(e) · Stand ${formatDeDate(new Date().toISOString().slice(0, 10))}`;
}
