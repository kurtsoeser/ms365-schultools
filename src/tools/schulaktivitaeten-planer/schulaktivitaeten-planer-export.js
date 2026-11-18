/**
 * ICS-Export / Feed-Inhalt für Schulaktivitäten.
 */
import {
    formatDeDate,
    labelMaps,
    statusLabel,
    typLabel,
    teilnahmeKurz
} from './schulaktivitaeten-planer-state.js';
import { toIsoDateOnly, parseCodeList } from './schulaktivitaeten-planer-logic.js';
import { applyUeToDraft, formatUeZeitraum } from './schulaktivitaeten-planer-zeitraster.js';

/**
 * @param {object[]} items
 * @param {object} stammdaten
 * @param {string} [calendarName]
 * @param {{ onlyApproved?: boolean }} [opts]
 */
export function buildIcs(items, stammdaten, calendarName, opts) {
    const labels = labelMaps(stammdaten || {});
    const onlyApproved = !!(opts && opts.onlyApproved);
    const list = (Array.isArray(items) ? items : []).filter((akt) => {
        const st = String((akt && akt.status) || '').toLowerCase();
        if (onlyApproved) return st === 'genehmigt';
        return st !== 'abgelehnt';
    });
    const calName = calendarName || 'Schulaktivitäten';
    const lines = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//MS365schule//Schulaktivitaeten-Planer//DE',
        'CALSCALE:GREGORIAN',
        'METHOD:PUBLISH',
        'X-WR-CALNAME:' + escapeIcsText(calName),
        'X-WR-CALDESC:' + escapeIcsText('Genehmigte Schulaktivitäten und Exkursionen'),
        'X-PUBLISHED-TTL:PT6H',
        'REFRESH-INTERVAL;VALUE=DURATION:PT6H'
    ];

    list.forEach((akt) => {
        const normalized = applyUeToDraft(akt).draft;
        const startIso = toIsoDateOnly(normalized.startdatum);
        const endIso = toIsoDateOnly(normalized.enddatum || normalized.startdatum);
        if (!startIso) return;
        const startCompact = startIso.replace(/-/g, '');
        if (!/^\d{8}$/.test(startCompact)) return;
        const endCompact = String(endIso || startIso).replace(/-/g, '');
        const uid = (normalized.aktivitaetId || normalized.itemId || startCompact) + '@schulaktivitaeten.ms365schule';

        const teilnahme = teilnahmeKurz(normalized, labels.klasse);
        const lehrerName =
            labels.lehrer[normalized.lehrerCode] || normalized.lehrerCode || normalized.lehrerEmail || '–';
        const begleitung = formatBegleitung(normalized, labels.lehrer);
        const ueLabel = formatUeZeitraum(normalized);
        const zeitraum =
            formatDeDate(normalized.startdatum) +
            (normalized.enddatum && normalized.enddatum !== normalized.startdatum
                ? ' – ' + formatDeDate(normalized.enddatum)
                : '') +
            (ueLabel ? ' · ' + ueLabel : '');

        const summary =
            typLabel(normalized.typ) +
            ': ' +
            (normalized.titel || 'Aktivität') +
            (teilnahme && teilnahme !== '–' ? ' · ' + teilnahme : '');

        const descLines = [
            'Titel: ' + (normalized.titel || '–'),
            'Typ: ' + typLabel(normalized.typ),
            'Status: ' + statusLabel(normalized.status),
            'Teilnahme: ' + (teilnahme || '–'),
            normalized.gruppenText ? 'Gruppe: ' + normalized.gruppenText : '',
            normalized.nichtTeilnehmende ? 'Nicht teilnehmend: ' + normalized.nichtTeilnehmende : '',
            'Verantwortlich: ' + lehrerName,
            normalized.lehrerEmail ? 'E-Mail: ' + normalized.lehrerEmail : '',
            begleitung ? 'Begleitung: ' + begleitung : '',
            normalized.externeVortragende ? 'Externe: ' + normalized.externeVortragende : '',
            normalized.ort ? 'Ort / Ziel: ' + normalized.ort : '',
            'Zeitraum: ' + zeitraum,
            normalized.unterrichtsbezug ? 'Unterrichtsbezug: ' + normalized.unterrichtsbezug : '',
            normalized.verkehrsmittel ? 'Verkehrsmittel: ' + normalized.verkehrsmittel : '',
            normalized.kostenHinweis ? 'Kosten: ' + normalized.kostenHinweis : '',
            normalized.ablauf ? 'Ablauf: ' + normalized.ablauf : '',
            normalized.notiz ? 'Notiz: ' + normalized.notiz : ''
        ].filter(Boolean);

        const htmlDesc = buildHtmlDescription({
            titel: normalized.titel || '–',
            typ: typLabel(normalized.typ),
            status: statusLabel(normalized.status),
            teilnahme: teilnahme || '–',
            gruppe: normalized.gruppenText || '',
            nichtTeilnehmende: normalized.nichtTeilnehmende || '',
            verantwortlich: lehrerName,
            email: normalized.lehrerEmail || '',
            begleitung: begleitung || '',
            externe: normalized.externeVortragende || '',
            ort: normalized.ort || '',
            zeitraum,
            unterrichtsbezug: normalized.unterrichtsbezug || '',
            verkehrsmittel: normalized.verkehrsmittel || '',
            kosten: normalized.kostenHinweis || '',
            ablauf: normalized.ablauf || '',
            notiz: normalized.notiz || ''
        });

        const timedSameDay =
            startIso === endIso &&
            /^\d{1,2}:\d{2}$/.test(String(normalized.startZeit || '')) &&
            /^\d{1,2}:\d{2}$/.test(String(normalized.endZeit || ''));

        lines.push('BEGIN:VEVENT');
        lines.push('UID:' + uid);
        lines.push('DTSTAMP:' + utcStamp());
        if (timedSameDay) {
            lines.push(
                'DTSTART:' + localFloating(startIso, normalized.startZeit),
                'DTEND:' + localFloating(endIso, normalized.endZeit)
            );
        } else {
            lines.push('DTSTART;VALUE=DATE:' + startCompact);
            lines.push('DTEND;VALUE=DATE:' + nextDayCompact(endCompact));
        }
        lines.push('SUMMARY:' + escapeIcsText(summary));
        lines.push('CATEGORIES:' + escapeIcsText(typLabel(normalized.typ)));
        if (normalized.ort) lines.push('LOCATION:' + escapeIcsText(normalized.ort));
        if (normalized.lehrerEmail) {
            lines.push(
                'ORGANIZER;CN=' +
                    escapeIcsParam(lehrerName) +
                    ':mailto:' +
                    String(normalized.lehrerEmail).trim()
            );
        }
        lines.push('DESCRIPTION:' + escapeIcsText(descLines.join('\n')));
        lines.push('X-ALT-DESC;FMTTYPE=text/html:' + escapeIcsText(htmlDesc));
        lines.push('END:VEVENT');
    });

    lines.push('END:VCALENDAR');
    return lines.join('\r\n');
}

/**
 * @param {object} item
 * @param {Record<string,string>} lehrerMap
 */
function formatBegleitung(item, lehrerMap) {
    const map = lehrerMap || {};
    const codes = parseCodeList(item.begleitungCodes || item.begleitung);
    if (codes.length) {
        return codes.map((c) => map[c] || c).join(', ');
    }
    return String(item.begleitung || '').trim();
}

/**
 * Tabellarische HTML-Beschreibung für Clients mit X-ALT-DESC (z. B. Outlook).
 * @param {Record<string,string>} row
 */
export function buildHtmlDescription(row) {
    const r = row || {};
    const rows = [
        ['Titel', r.titel],
        ['Typ', r.typ],
        ['Status', r.status],
        ['Teilnahme', r.teilnahme],
        r.gruppe ? ['Gruppe', r.gruppe] : null,
        r.nichtTeilnehmende ? ['Nicht teilnehmend', r.nichtTeilnehmende] : null,
        ['Verantwortlich', r.verantwortlich],
        r.email ? ['E-Mail', r.email] : null,
        r.begleitung ? ['Begleitung', r.begleitung] : null,
        r.externe ? ['Externe', r.externe] : null,
        r.ort ? ['Ort / Ziel', r.ort] : null,
        ['Zeitraum', r.zeitraum],
        r.unterrichtsbezug ? ['Unterrichtsbezug', r.unterrichtsbezug] : null,
        r.verkehrsmittel ? ['Verkehrsmittel', r.verkehrsmittel] : null,
        r.kosten ? ['Kosten', r.kosten] : null,
        r.ablauf ? ['Ablauf', r.ablauf] : null,
        r.notiz ? ['Notiz', r.notiz] : null
    ].filter(Boolean);

    const tr = rows
        .map(
            ([k, v]) =>
                '<tr><th style="text-align:left;padding:2px 10px 2px 0;vertical-align:top;white-space:nowrap;">' +
                escapeHtml(k) +
                '</th><td style="padding:2px 0;">' +
                escapeHtml(v) +
                '</td></tr>'
        )
        .join('');
    return (
        '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN"><html><body>' +
        '<table style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:12px;">' +
        tr +
        '</table></body></html>'
    );
}

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function localFloating(isoDate, hhmm) {
    const [h, m] = String(hhmm)
        .split(':')
        .map((x) => String(Number(x)).padStart(2, '0'));
    return String(isoDate).replace(/-/g, '') + 'T' + h + m + '00';
}

function escapeIcsText(s) {
    return String(s || '')
        .replace(/\\/g, '\\\\')
        .replace(/;/g, '\\;')
        .replace(/,/g, '\\,')
        .replace(/\n/g, '\\n');
}

function escapeIcsParam(s) {
    return String(s || '')
        .replace(/\\/g, '\\\\')
        .replace(/;/g, '\\;')
        .replace(/,/g, '\\,')
        .replace(/\n/g, ' ')
        .replace(/"/g, '');
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

/**
 * HTTPS → webcal für Kalender-Abos (Outlook/Apple).
 * @param {string} httpsUrl
 */
export function toWebcalUrl(httpsUrl) {
    const u = String(httpsUrl || '').trim();
    if (!u) return '';
    if (/^webcal:/i.test(u)) return u;
    return u.replace(/^https:/i, 'webcal:').replace(/^http:/i, 'webcal:');
}
