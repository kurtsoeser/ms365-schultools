/**
 * Export: CSV, iCal, Druck (Wochenplan / Teilnehmer).
 */
import { attendeesToCsv } from './projektwochen-bookings-logic.js';
import { buildWeekPlan, buildWeekPlanDisplayRows, PLAN_SLOTS, toIsoDateOnly } from './projektwochen-logic.js';

/**
 * @param {string} filename
 * @param {string} text
 * @param {string} [mime]
 */
export function downloadTextFile(filename, text, mime) {
    const blob = new Blob([text], { type: mime || 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename || 'export.csv';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
}

/**
 * @param {object[]} rows
 * @param {string} [filename]
 */
export function downloadAttendeesCsv(rows, filename) {
    const bom = '\uFEFF';
    downloadTextFile(filename || 'projektwochen-teilnehmer.csv', bom + attendeesToCsv(rows));
}

/**
 * @param {object[]} angebote
 * @returns {string}
 */
export function angeboteToCsv(angebote) {
    const header = [
        'Titel',
        'Status',
        'Datum',
        'Tag',
        'Slot',
        'Start',
        'Ende',
        'Ort',
        'Kapazität',
        'Preis',
        'Lehrer',
        'Zielklassen',
        'Kategorie',
        'BuchungAb',
        'SyncStatus',
        'BookingsUrl',
        'AngebotId'
    ];
    const lines = [header.join(';')];
    (angebote || []).forEach((a) => {
        const cells = [
            a.title,
            a.status,
            a.datum,
            a.tag,
            a.slot,
            a.startzeit,
            a.endzeit,
            a.ort,
            a.kapazitaet,
            a.preisEuro,
            a.lehrerCode || a.lehrerEmail,
            a.zielklassen,
            a.kategorie,
            a.buchungAb,
            a.syncStatus,
            a.bookingsBookingUrl,
            a.angebotId
        ].map(csvCell);
        lines.push(cells.join(';'));
    });
    return lines.join('\r\n');
}

function csvCell(v) {
    const s = String(v == null ? '' : v);
    if (/[;"\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
}

/**
 * @param {object[]} angebote
 * @param {string} [filename]
 */
export function downloadAngeboteCsv(angebote, filename) {
    downloadTextFile(filename || 'projektwochen-angebote.csv', '\uFEFF' + angeboteToCsv(angebote));
}

/**
 * @param {object[]} angebote
 * @param {object|null} aktion
 * @param {string} [calendarName]
 */
export function buildAngeboteIcs(angebote, aktion, calendarName) {
    const list = Array.isArray(angebote) ? angebote : [];
    const name =
        calendarName ||
        (aktion && aktion.title) ||
        'Projektwochen';
    const lines = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//MS365schule//Projektwochen//DE',
        'CALSCALE:GREGORIAN',
        'METHOD:PUBLISH',
        'X-WR-CALNAME:' + escapeIcsText(name)
    ];

    list.forEach((a) => {
        const day = String(toIsoDateOnly(a.datum) || '').replace(/-/g, '');
        if (!/^\d{8}$/.test(day)) return;
        const start = toIcsTime(day, a.startzeit || '08:00');
        const end = toIcsTime(day, a.endzeit || '12:00');
        const uid = (a.angebotId || a.itemId || day) + '@projektwochen.ms365schule';
        const summary = (a.title || 'Angebot') + (a.status ? ' [' + a.status + ']' : '');
        const desc = [
            a.ort ? 'Ort: ' + a.ort : '',
            a.treffpunkt ? 'Treffpunkt: ' + a.treffpunkt : '',
            a.lehrerCode ? 'Lehrer: ' + a.lehrerCode : '',
            a.kapazitaet != null ? 'Kapazität: ' + a.kapazitaet : '',
            a.preisEuro != null ? 'Preis: ' + a.preisEuro + ' €' : '',
            a.beschreibung || ''
        ]
            .filter(Boolean)
            .join('\\n');

        lines.push('BEGIN:VEVENT');
        lines.push('UID:' + uid);
        lines.push('DTSTAMP:' + utcStamp());
        lines.push('DTSTART:' + start);
        lines.push('DTEND:' + end);
        lines.push('SUMMARY:' + escapeIcsText(summary));
        if (desc) lines.push('DESCRIPTION:' + escapeIcsText(desc));
        if (a.ort) lines.push('LOCATION:' + escapeIcsText(a.ort));
        lines.push('END:VEVENT');
    });

    lines.push('END:VCALENDAR');
    return lines.join('\r\n');
}

/**
 * @param {string} ics
 * @param {string} [filename]
 */
export function downloadIcs(ics, filename) {
    downloadTextFile(filename || 'projektwochen.ics', ics, 'text/calendar;charset=utf-8');
}

function toIcsTime(yyyymmdd, hhmm) {
    const m = String(hhmm || '08:00').match(/^(\d{1,2}):(\d{2})/);
    const h = m ? String(Math.min(23, Number(m[1]))).padStart(2, '0') : '08';
    const min = m ? m[2] : '00';
    return yyyymmdd + 'T' + h + min + '00';
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

function escHtml(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * Druckbares HTML für Wochenplan.
 * @param {object[]} angebote
 * @param {object|null} aktion
 */
export function buildWeekPlanPrintHtml(angebote, aktion) {
    const grid = buildWeekPlan(angebote);
    const displayRows = buildWeekPlanDisplayRows(grid);
    const title = (aktion && aktion.title) || 'Projektwochen-Plan';
    const range =
        aktion && aktion.startdatum
            ? escHtml(aktion.startdatum) + ' – ' + escHtml(aktion.enddatum || '')
            : '';
    const slotHeads = PLAN_SLOTS.map((s) => '<th>' + escHtml(s) + '</th>').join('');

    function planCards(list) {
        return (list || [])
            .map((a) => {
                return (
                    '<div class="card"><strong>' +
                    escHtml(a.title) +
                    '</strong><br>' +
                    escHtml(a.ort || '') +
                    '<br>' +
                    escHtml(String(a.kapazitaet || '')) +
                    ' Plätze · ' +
                    escHtml(a.status || '') +
                    (a.lehrerCode ? '<br>' + escHtml(a.lehrerCode) : '') +
                    '</div>'
                );
            })
            .join('');
    }

    const rows = displayRows
        .map((day) => {
            const bandCount = day.bands.length;
            return day.bands
                .map((band, bi) => {
                    const th =
                        bi === 0
                            ? '<th' +
                              (bandCount > 1 ? ' rowspan="' + bandCount + '"' : '') +
                              '>' +
                              escHtml(day.tag) +
                              '</th>'
                            : '';
                    if (band.type === 'ganztags') {
                        return (
                            '<tr>' +
                            th +
                            '<td class="ganztags" colspan="' +
                            PLAN_SLOTS.length +
                            '">' +
                            (planCards(band.items) || '–') +
                            '</td></tr>'
                        );
                    }
                    const cells = PLAN_SLOTS.map((slot) => {
                        return '<td>' + (planCards(band.bySlot[slot]) || '–') + '</td>';
                    }).join('');
                    return '<tr>' + th + cells + '</tr>';
                })
                .join('');
        })
        .join('');

    return (
        '<!DOCTYPE html><html lang="de"><head><meta charset="UTF-8"><title>' +
        escHtml(title) +
        '</title><style>' +
        'body{font-family:Segoe UI,system-ui,sans-serif;margin:24px;color:#111}' +
        'h1{font-size:1.4rem;margin:0 0 4px} .meta{color:#555;margin-bottom:16px}' +
        'table{border-collapse:collapse;width:100%;font-size:12px}' +
        'th,td{border:1px solid #ccc;padding:8px;vertical-align:top}' +
        'th{background:#f3f4f6}' +
        'td.ganztags{background:#f0f7ff}' +
        '.card{margin-bottom:6px;padding:4px 6px;background:#f8fafc;border-radius:4px}' +
        '@media print{body{margin:12px} .no-print{display:none}}' +
        '</style></head><body>' +
        '<p class="no-print"><button onclick="window.print()">Drucken</button></p>' +
        '<h1>' +
        escHtml(title) +
        '</h1>' +
        '<p class="meta">' +
        range +
        ' · ' +
        (angebote || []).length +
        ' Angebote · Stand ' +
        escHtml(new Date().toLocaleString('de-AT')) +
        '</p>' +
        '<table><thead><tr><th>Tag</th>' +
        slotHeads +
        '</tr></thead><tbody>' +
        rows +
        '</tbody></table></body></html>'
    );
}

/**
 * Druckbares HTML für Teilnehmerliste.
 * @param {object[]} rows
 * @param {object|null} aktion
 */
export function buildAttendeesPrintHtml(rows, aktion) {
    const title = 'Teilnehmer – ' + ((aktion && aktion.title) || 'Projektwoche');
    const body = (rows || [])
        .map((r) => {
            return (
                '<tr><td>' +
                escHtml(r.name) +
                '</td><td>' +
                escHtml(r.email) +
                '</td><td>' +
                escHtml(r.klasse || '') +
                '</td><td>' +
                escHtml(r.angebotTitle) +
                '</td><td>' +
                escHtml(r.datum || '') +
                '</td></tr>'
            );
        })
        .join('');
    return (
        '<!DOCTYPE html><html lang="de"><head><meta charset="UTF-8"><title>' +
        escHtml(title) +
        '</title><style>' +
        'body{font-family:Segoe UI,system-ui,sans-serif;margin:24px}' +
        'table{border-collapse:collapse;width:100%;font-size:13px}' +
        'th,td{border:1px solid #ccc;padding:6px 8px;text-align:left}' +
        'th{background:#f3f4f6}' +
        '@media print{.no-print{display:none}}' +
        '</style></head><body>' +
        '<p class="no-print"><button onclick="window.print()">Drucken</button></p>' +
        '<h1>' +
        escHtml(title) +
        '</h1>' +
        '<p>' +
        (rows || []).length +
        ' Einträge · ' +
        escHtml(new Date().toLocaleString('de-AT')) +
        '</p>' +
        '<table><thead><tr><th>Name</th><th>E-Mail</th><th>Klasse</th><th>Angebot</th><th>Datum</th></tr></thead><tbody>' +
        (body || '<tr><td colspan="5">Keine Daten</td></tr>') +
        '</tbody></table></body></html>'
    );
}

/**
 * @param {string} html
 */
export function openPrintWindow(html) {
    const w = window.open('', '_blank');
    if (!w) throw new Error('Popup blockiert – bitte Popups erlauben.');
    w.document.open();
    w.document.write(html);
    w.document.close();
}
