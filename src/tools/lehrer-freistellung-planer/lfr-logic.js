/**
 * Validierung & Filter – Freistellungen Lehrkräfte.
 */
import { STATUS_CHOICES, KATEGORIE_CHOICES } from './lfr-schema.js';
import {
    toIsoDateOnly,
    toIsoDateTimeLocal,
    toSharePointDateTime,
    toDateTimeMs,
    hasClockTime
} from '../freistellung-planer/freistellung-planer-logic.js';

export { toIsoDateOnly, toIsoDateTimeLocal, toSharePointDateTime, hasClockTime };

/** ISO-8601 Datum/Zeit inkl. optionalem Z / Offset (wie von Power Automate / Graph). */
const ISO_DT_IN_TEXT_RE =
    /\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2})?(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?/g;

/**
 * @param {object} draft
 * @returns {{ ok: boolean, errors: string[] }}
 */
export function validateAntrag(draft) {
    const errors = [];
    const titel = String(draft.titel || '').trim();
    if (titel.length < 2) errors.push('Bitte einen kurzen Titel angeben.');
    const beginn = toIsoDateTimeLocal(draft.beginn);
    const ende = toIsoDateTimeLocal(draft.ende);
    if (!beginn) errors.push('Beginn (Datum/Uhrzeit) fehlt.');
    if (!ende) errors.push('Ende (Datum/Uhrzeit) fehlt.');
    if (beginn && ende && toDateTimeMs(ende) < toDateTimeMs(beginn)) {
        errors.push('Ende liegt vor Beginn.');
    }
    const kat = String(draft.kategorie || '').trim();
    if (!kat) errors.push('Kategorie wählen.');
    else if (!KATEGORIE_CHOICES.some((k) => k.toLowerCase() === kat.toLowerCase())) {
        errors.push('Unbekannte Kategorie.');
    }
    const mail = String(draft.lehrerEmail || '').trim().toLowerCase();
    if (!mail.includes('@')) errors.push('Lehrer-E-Mail fehlt (Anmeldung oder Stammdaten).');
    return { ok: errors.length === 0, errors };
}

/**
 * @param {string} status
 */
export function normalizeStatus(status) {
    const s = String(status || '').trim();
    const hit = STATUS_CHOICES.find((x) => x.toLowerCase() === s.toLowerCase());
    return hit || 'Ausstehend';
}

/**
 * @param {object} item
 * @param {string} dayIso YYYY-MM-DD
 */
export function itemCoversDay(item, dayIso) {
    const day = toIsoDateOnly(dayIso);
    if (!day) return false;
    const b = toIsoDateOnly(item.beginn);
    const e = toIsoDateOnly(item.ende) || b;
    if (!b) return false;
    return day >= b && day <= e;
}

/**
 * E-Mail-Vergleich (UPN/Mail, Case).
 * @param {string} a
 * @param {string} b
 */
export function emailsMatchLfr(a, b) {
    const x = String(a || '')
        .trim()
        .toLowerCase();
    const y = String(b || '')
        .trim()
        .toLowerCase();
    return !!x && !!y && x === y;
}

function normPersonName(v) {
    return String(v || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ' ');
}

/**
 * Gehört der Antrag der angemeldeten Lehrkraft?
 * Mail, Author, Anzeigename oder SharePoint-„nur eigene“ (callerVisible).
 * @param {object} item
 * @param {string} accountEmail
 * @param {string} [accountName]
 */
export function itemOwnedByAccount(item, accountEmail, accountName) {
    if (item && item.callerVisible) return true;
    const email = String(accountEmail || '')
        .trim()
        .toLowerCase();
    const em = String((item && item.lehrerEmail) || '')
        .trim()
        .toLowerCase();
    if (email && emailsMatchLfr(em, email)) return true;
    const author = String((item && item.authorEmail) || '')
        .trim()
        .toLowerCase();
    if (email && emailsMatchLfr(author, email)) return true;
    const want = normPersonName(accountName);
    const got = normPersonName(item && item.lehrerName);
    return !!(want && got && want === got);
}

/**
 * @param {object[]} items
 * @param {object} filters
 * @param {{ scopeAll?: boolean, accountEmail?: string }} scope
 */
export function filterItems(items, filters, scope) {
    const f = filters || {};
    const email = String(scope.accountEmail || '').trim().toLowerCase();
    const scopeAll = !!scope.scopeAll;
    return (items || []).filter((it) => {
        if (!scopeAll) {
            if (!itemOwnedByAccount(it, email, scope.accountName)) return false;
        }
        if (f.status && normalizeStatus(it.status) !== normalizeStatus(f.status)) return false;
        if (f.kategorie && String(it.kategorie || '') !== String(f.kategorie)) return false;
        if (f.q) {
            const hay = [it.titel, it.beschreibung, it.lehrerName, it.kategorie]
                .join(' ')
                .toLowerCase();
            if (!hay.includes(String(f.q).trim().toLowerCase())) return false;
        }
        return true;
    });
}

/**
 * @param {number} year
 * @param {number} month 1–12
 * @returns {string[]} ISO dates for calendar grid (Mo-start, 6 weeks)
 */
export function monthGridDates(year, month) {
    const first = new Date(year, month - 1, 1);
    let dow = first.getDay();
    if (dow === 0) dow = 7;
    const start = new Date(year, month - 1, 1 - (dow - 1));
    const out = [];
    for (let i = 0; i < 42; i++) {
        const d = new Date(start.getFullYear(), start.getMonth(), start.getDate() + i);
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        out.push(`${y}-${m}-${day}`);
    }
    return out;
}

export function formatDeDate(iso) {
    const d = toIsoDateOnly(iso);
    if (!d) return '–';
    const [y, m, day] = d.split('-');
    return `${day}.${m}.${y}`;
}

/**
 * Datum/Uhrzeit für Anzeige (lokale Wandzeit, ohne Z/T-Rohformat).
 * @param {unknown} iso
 */
export function formatDeDateTime(iso) {
    const local = toIsoDateTimeLocal(iso);
    if (!local) return '–';
    const datePart = formatDeDate(local);
    if (!hasClockTime(local)) return datePart;
    const time = local.split('T')[1] || '';
    return datePart + ', ' + time + ' Uhr';
}

/**
 * Ersetzt ISO-Zeitstempel in Freitext (Titel, Bemerkung Direktion aus PA) für die Anzeige.
 * Speichert nichts um – nur Darstellung.
 * @param {unknown} text
 */
export function formatDisplayTextWithDeDates(text) {
    const s = String(text == null ? '' : text);
    if (!s) return '';
    return s.replace(ISO_DT_IN_TEXT_RE, (match) => {
        const pretty = formatDeDateTime(match);
        return pretty && pretty !== '–' ? pretty : match;
    });
}

export function formatDeDateTimeRange(beginn, ende) {
    const b = toIsoDateTimeLocal(beginn);
    const e = toIsoDateTimeLocal(ende);
    if (!b) return '–';
    const bd = formatDeDate(b);
    const ed = e ? formatDeDate(e) : bd;
    const bt = b.includes('T') ? b.split('T')[1] : '';
    const et = e && e.includes('T') ? e.split('T')[1] : '';
    if (bd === ed && bt && et && bt !== '00:00') return `${bd} ${bt}–${et}`;
    if (bd !== ed) return `${bd} – ${ed}`;
    return bd;
}
