/**
 * ICS-Export für Schulaktivitäten.
 */
import { formatDeDate, labelMaps, statusLabel, typLabel } from './schulaktivitaeten-planer-state.js';
import { toIsoDateOnly } from './schulaktivitaeten-planer-logic.js';

/**
 * @param {object[]} items
 * @param {object} stammdaten
 * @param {string} [calendarName]
 */
export function buildIcs(items, stammdaten, calendarName) {
    const labels = labelMaps(stammdaten || {});
    const list = Array.isArray(items) ? items : [];
    const lines = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//MS365schule//Schulaktivitaeten-Planer//DE',
        'CALSCALE:GREGORIAN',
        'METHOD:PUBLISH',
        'X-WR-CALNAME:' + escapeIcsText(calendarName || 'Schulaktivitäten')
    ];

    list.forEach((akt) => {
        const start = String(toIsoDateOnly(akt.startdatum) || '').replace(/-/g, '');
        if (!/^\d{8}$/.test(start)) return;
        let endExclusive = String(toIsoDateOnly(akt.enddatum || akt.startdatum) || '').replace(/-/g, '');
        endExclusive = nextDayCompact(endExclusive);
        const uid = (akt.aktivitaetId || akt.itemId || start) + '@schulaktivitaeten.ms365schule';
        const summary =
            typLabel(akt.typ) +
            ': ' +
            (akt.titel || 'Aktivität') +
            ' · ' +
            (labels.klasse[akt.klasseCode] || akt.klasseCode || '');
        const desc = [
            'Status: ' + statusLabel(akt.status),
            akt.ort ? 'Ort: ' + akt.ort : '',
            akt.begleitung ? 'Begleitung: ' + akt.begleitung : '',
            akt.verkehrsmittel ? 'Verkehrsmittel: ' + akt.verkehrsmittel : '',
            akt.kostenHinweis ? 'Kosten: ' + akt.kostenHinweis : '',
            akt.notiz ? 'Notiz: ' + akt.notiz : '',
            'Zeitraum: ' + formatDeDate(akt.startdatum) + ' – ' + formatDeDate(akt.enddatum || akt.startdatum)
        ]
            .filter(Boolean)
            .join('\\n');

        lines.push('BEGIN:VEVENT');
        lines.push('UID:' + uid);
        lines.push('DTSTAMP:' + utcStamp());
        lines.push('DTSTART;VALUE=DATE:' + start);
        lines.push('DTEND;VALUE=DATE:' + endExclusive);
        lines.push('SUMMARY:' + escapeIcsText(summary));
        if (akt.ort) lines.push('LOCATION:' + escapeIcsText(akt.ort));
        if (desc) lines.push('DESCRIPTION:' + escapeIcsText(desc));
        lines.push('END:VEVENT');
    });

    lines.push('END:VCALENDAR');
    return lines.join('\r\n');
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
    a.download = filename || 'schulaktivitaeten.ics';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
}
