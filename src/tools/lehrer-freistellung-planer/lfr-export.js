/**
 * iCal (.ics) für genehmigte Lehrer-Freistellungen.
 */
import {
    formatDeDateTimeRange,
    toIsoDateOnly,
    toIsoDateTimeLocal,
    normalizeStatus
} from './lfr-logic.js';

/**
 * @param {object[]} items
 * @param {string} [calendarName]
 */
export function buildIcs(items, calendarName) {
    const list = (items || []).filter((it) => normalizeStatus(it.status) === 'Genehmigt');
    const lines = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//MS365schule//Lehrer-Freistellungen//DE',
        'CALSCALE:GREGORIAN',
        'METHOD:PUBLISH',
        'X-WR-CALNAME:' + escapeIcsText(calendarName || 'Lehrer-Freistellungen')
    ];

    list.forEach((it) => {
        const uid = (it.antragId || it.itemId || 'lfr') + '@lehrer-freistellung.ms365schule';
        const beginLocal = toIsoDateTimeLocal(it.beginn);
        const endLocal = toIsoDateTimeLocal(it.ende);
        if (!beginLocal) return;
        const summary = (it.lehrerName ? it.lehrerName + ': ' : '') + (it.titel || 'Freistellung');
        const desc = [
            'Kategorie: ' + (it.kategorie || ''),
            'Zeitraum: ' + formatDeDateTimeRange(it.beginn, it.ende),
            it.beschreibung ? it.beschreibung : ''
        ]
            .filter(Boolean)
            .join('\\n');

        lines.push('BEGIN:VEVENT');
        lines.push('UID:' + uid);
        lines.push('DTSTAMP:' + utcStamp());
        const hasTime = beginLocal.includes('T') && !beginLocal.endsWith('T00:00');
        if (hasTime && endLocal) {
            lines.push('DTSTART:' + toIcsUtc(beginLocal));
            lines.push('DTEND:' + toIcsUtc(endLocal));
        } else {
            const start = String(toIsoDateOnly(beginLocal) || '').replace(/-/g, '');
            let endDay = String(toIsoDateOnly(endLocal || beginLocal) || '').replace(/-/g, '');
            endDay = nextDayCompact(endDay);
            lines.push('DTSTART;VALUE=DATE:' + start);
            lines.push('DTEND;VALUE=DATE:' + endDay);
        }
        lines.push('SUMMARY:' + escapeIcsText(summary));
        if (desc) lines.push('DESCRIPTION:' + escapeIcsText(desc));
        lines.push('END:VEVENT');
    });

    lines.push('END:VCALENDAR');
    return lines.join('\r\n');
}

function toIcsUtc(isoLocal) {
    const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(String(isoLocal || ''));
    if (!m) return utcStamp();
    const d = new Date(
        Number(m[1]),
        Number(m[2]) - 1,
        Number(m[3]),
        Number(m[4]),
        Number(m[5]),
        0
    );
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

export function downloadIcs(ics, filename) {
    const blob = new Blob([ics], { type: 'text/calendar;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename || 'lehrer-freistellungen.ics';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
}
