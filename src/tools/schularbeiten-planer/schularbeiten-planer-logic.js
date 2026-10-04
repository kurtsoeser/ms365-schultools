/**
 * Regel-Engine für Schularbeiten (SchUG § 17 / LBVO § 7, HAK).
 * Rein, ohne DOM/fetch – Vitest-tauglich.
 */

/** @typedef {'beantragt'|'fixiert'|'abgelehnt'} SaStatus */
/** @typedef {'gesperrt'|'erlaubt'} FensterTyp */

/**
 * @typedef {object} Regelwerk
 * @property {number} maxProTag
 * @property {number} maxProWoche
 * @property {number} ankuendigungsfristTage
 * @property {number} sperreVorNotenkonferenzTage
 */

/**
 * @typedef {object} Terminfenster
 * @property {string} titel
 * @property {FensterTyp} typ
 * @property {string} startdatum ISO YYYY-MM-DD
 * @property {string} enddatum ISO YYYY-MM-DD
 */

/**
 * @typedef {object} Schularbeit
 * @property {string} [schularbeitId]
 * @property {string} fachCode
 * @property {string} klasseCode
 * @property {string} lehrerCode
 * @property {string} datum ISO YYYY-MM-DD
 * @property {string} [beginnUhrzeit] HH:mm (lokal)
 * @property {number} dauerMinuten
 * @property {string} [semester] WS|SS
 * @property {SaStatus} [status]
 */

/**
 * @typedef {object} ValidateOptions
 * @property {Schularbeit} draft
 * @property {Schularbeit[]} existing
 * @property {Regelwerk} rules
 * @property {Terminfenster[]} windows
 * @property {string} [today] ISO YYYY-MM-DD (Default: lokales Heute)
 * @property {{ proSemester?: number, minDauer?: number, maxDauer?: number }} [fachMeta]
 */

export const DEFAULT_RULES = {
    maxProTag: 1,
    maxProWoche: 2,
    ankuendigungsfristTage: 7,
    sperreVorNotenkonferenzTage: 7
};

/**
 * @param {unknown} value
 * @returns {string|null} YYYY-MM-DD
 */
export function toIsoDateOnly(value) {
    if (value == null) return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        const y = value.getFullYear();
        const m = String(value.getMonth() + 1).padStart(2, '0');
        const d = String(value.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }
    const s = String(value).trim();
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
    if (!m) return null;
    return `${m[1]}-${m[2]}-${m[3]}`;
}

/**
 * @param {string} iso
 * @returns {{ y: number, m: number, d: number }|null}
 */
export function parseIsoDateParts(iso) {
    const s = toIsoDateOnly(iso);
    if (!s) return null;
    const [y, m, d] = s.split('-').map(Number);
    if (!y || !m || !d) return null;
    return { y, m, d };
}

/** UTC-Mittag für DST-sichere Tagesdifferenz (unabhängig vom Browser-TZ). */
function utcNoonMs(iso) {
    const p = parseIsoDateParts(iso);
    if (!p) return NaN;
    return Date.UTC(p.y, p.m - 1, p.d, 12, 0, 0);
}

/** ISO-Datum (YYYY-MM-DD) aus UTC-Millisekunden – nie Browser-Lokalzeit. */
function isoDateFromUtcMs(ms) {
    if (!Number.isFinite(ms)) return null;
    const d = new Date(ms);
    const p = (n) => String(n).padStart(2, '0');
    return `${d.getUTCFullYear()}-${p(d.getUTCMonth() + 1)}-${p(d.getUTCDate())}`;
}

/**
 * @param {string} fromIso
 * @param {string} toIso
 * @returns {number|null}
 */
export function daysBetween(fromIso, toIso) {
    const a = utcNoonMs(fromIso);
    const b = utcNoonMs(toIso);
    if (Number.isNaN(a) || Number.isNaN(b)) return null;
    return Math.round((b - a) / 86400000);
}

/**
 * ISO-Kalenderwochen-Schlüssel, z. B. 2026-W42
 * @param {string} iso
 */
export function isoWeekKey(iso) {
    const p = parseIsoDateParts(iso);
    if (!p) return '';
    const date = new Date(Date.UTC(p.y, p.m - 1, p.d));
    // Donnerstag der aktuellen Woche bestimmt das ISO-Jahr
    const day = date.getUTCDay() || 7;
    date.setUTCDate(date.getUTCDate() + 4 - day);
    const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
    const week = Math.ceil(((date - yearStart) / 86400000 + 1) / 7);
    const isoYear = date.getUTCFullYear();
    return `${isoYear}-W${String(week).padStart(2, '0')}`;
}

/**
 * @param {string} iso
 * @param {number} deltaDays
 */
export function addDays(iso, deltaDays) {
    const ms = utcNoonMs(iso);
    if (Number.isNaN(ms)) return null;
    return isoDateFromUtcMs(ms + deltaDays * 86400000);
}

/**
 * @param {string} [raw]
 * @returns {string} HH:mm oder leer
 */
export function normalizeBeginnUhrzeit(raw) {
    const t = String(raw || '').trim();
    const m = /^(\d{1,2}):(\d{2})$/.exec(t);
    if (!m) return '';
    const h = parseInt(m[1], 10);
    const mi = parseInt(m[2], 10);
    if (!Number.isFinite(h) || !Number.isFinite(mi) || h < 0 || h > 23 || mi < 0 || mi > 59) return '';
    return String(h).padStart(2, '0') + ':' + String(mi).padStart(2, '0');
}

function timeToMinutes(hhmm) {
    const n = normalizeBeginnUhrzeit(hhmm);
    if (!n) return NaN;
    const [h, m] = n.split(':').map((x) => parseInt(x, 10));
    return h * 60 + m;
}

function minutesToTime(total) {
    const mins = ((total % (24 * 60)) + 24 * 60) % (24 * 60);
    const h = Math.floor(mins / 60);
    const m = mins % 60;
    return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
}

/**
 * Ende-Uhrzeit am selben Kalendertag (Überlauf nach Mitternacht wird gekürzt).
 * @param {string} beginnUhrzeit
 * @param {number} dauerMinuten
 */
export function computeEndeUhrzeit(beginnUhrzeit, dauerMinuten) {
    const start = normalizeBeginnUhrzeit(beginnUhrzeit);
    if (!start) return '';
    const dur = Number(dauerMinuten);
    if (!Number.isFinite(dur) || dur <= 0) return '';
    return minutesToTime(timeToMinutes(start) + dur);
}

/**
 * @param {object} sa
 */
export function formatSchularbeitZeitspanne(sa) {
    const start = normalizeBeginnUhrzeit(sa && sa.beginnUhrzeit);
    if (!start) return '';
    const end = computeEndeUhrzeit(start, sa.dauerMinuten);
    return end ? `${start} – ${end} Uhr` : `${start} Uhr`;
}

/** @param {object} sa */
export function formatSchularbeitZeitKurz(sa) {
    const start = normalizeBeginnUhrzeit(sa && sa.beginnUhrzeit);
    if (!start) return '';
    const end = computeEndeUhrzeit(start, sa.dauerMinuten);
    return end ? `${start}–${end}` : start;
}

/**
 * Minuten seit Mitternacht; ohne Uhrzeit → ans Ende des Tages sortieren.
 * @param {object} sa
 */
export function schularbeitBeginnMinuten(sa) {
    const n = normalizeBeginnUhrzeit(sa && sa.beginnUhrzeit);
    if (!n) return 24 * 60 + 1;
    return timeToMinutes(n);
}

/**
 * Sortierung innerhalb eines Tages: Beginn, Klasse, Fach, Titel.
 * @param {object} a
 * @param {object} b
 */
export function compareSchularbeitenByBeginn(a, b) {
    const ma = schularbeitBeginnMinuten(a);
    const mb = schularbeitBeginnMinuten(b);
    if (ma !== mb) return ma - mb;
    const k = String(a.klasseCode || '').localeCompare(String(b.klasseCode || ''), 'de');
    if (k) return k;
    const f = String(a.fachCode || '').localeCompare(String(b.fachCode || ''), 'de');
    if (f) return f;
    return schularbeitDisplayTitle(a).localeCompare(schularbeitDisplayTitle(b), 'de');
}

/**
 * @param {object[]} items
 */
export function sortSchularbeitenByBeginn(items) {
    return (Array.isArray(items) ? items : []).slice().sort(compareSchularbeitenByBeginn);
}

/**
 * Graph/ICS: Start- und End-ISO-Datumzeit als **Wanduhrzeit in Europe/Vienna**
 * (kein UTC-Offset im String). Sommer-/Winterzeit wird von Graph (`timeZone`) bzw.
 * ICS (`TZID` + `VTIMEZONE`) aufgelöst – Uhrzeiten aus dem Planer sind Schulortszeit.
 * @param {object} sa
 * @returns {{ isAllDay: boolean, startDateTime?: string, endDateTime?: string, startDate?: string, endDate?: string }|null}
 */
export function schularbeitTerminZeitfenster(sa) {
    const datum = toIsoDateOnly(sa && sa.datum);
    if (!datum) return null;
    const startTime = normalizeBeginnUhrzeit(sa && sa.beginnUhrzeit);
    const dur = Number(sa && sa.dauerMinuten);
    if (!startTime || !Number.isFinite(dur) || dur <= 0) {
        const endDate = addDays(datum, 1);
        return {
            isAllDay: true,
            startDate: datum,
            endDate: endDate || datum
        };
    }
    const startMins = timeToMinutes(startTime);
    const endMins = startMins + dur;
    const dayOffset = Math.floor(endMins / (24 * 60));
    const endDay = addDays(datum, dayOffset) || datum;
    const endTime = minutesToTime(endMins);
    return {
        isAllDay: false,
        startDateTime: `${datum}T${startTime}:00`,
        endDateTime: `${endDay}T${endTime}:00`
    };
}

/** IANA-Zeitzone für Graph/Outlook/ICS (Schulbetrieb AT, EU-DST-Regeln). */
export const SCHULARBEIT_CALENDAR_TZ = 'Europe/Vienna';

/**
 * RFC 5545 VTIMEZONE für Europe/Vienna (CET/CEST, letzter So im März/Oktober).
 * Pflicht für zuverlässigen ICS-Import mit TZID (Outlook, Apple, Google).
 */
export function icsViennaTimezoneBlock() {
    return [
        'BEGIN:VTIMEZONE',
        'TZID:Europe/Vienna',
        'X-LIC-LOCATION:Europe/Vienna',
        'BEGIN:DAYLIGHT',
        'TZOFFSETFROM:+0100',
        'TZOFFSETTO:+0200',
        'TZNAME:CEST',
        'DTSTART:19700329T020000',
        'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU',
        'END:DAYLIGHT',
        'BEGIN:STANDARD',
        'TZOFFSETFROM:+0200',
        'TZOFFSETTO:+0100',
        'TZNAME:CET',
        'DTSTART:19701025T030000',
        'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
        'END:STANDARD',
        'END:VTIMEZONE'
    ].join('\r\n');
}

/**
 * Betreff für Outlook/ICS: Feld **Titel** (Fallback Thema), mit Fach und Klasse.
 * @param {object} sa
 * @param {{ fach?: string, klasse?: string }} [labels]
 */
export function schularbeitCalendarSubject(sa, labels) {
    const title = schularbeitDisplayTitle(sa);
    const fach = String((labels && labels.fach) || (sa && sa.fachCode) || '').trim();
    const klasse = String((labels && labels.klasse) || (sa && sa.klasseCode) || '').trim();
    const ctx = [fach, klasse].filter(Boolean).join(' · ');
    const line = ctx ? `${title} · ${ctx}` : title;
    return line.slice(0, 250);
}

/**
 * Graph `dateTime` (lokal, mit Sekundenbruch).
 * @param {string} isoLocal z. B. 2026-11-11T08:35:00 oder 2026-11-11T08:35
 */
export function graphCalendarDateTime(isoLocal) {
    const s = String(isoLocal || '').trim();
    if (!s) return '';
    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(s)) {
        return s + '.0000000';
    }
    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(s)) {
        return s + ':00.0000000';
    }
    return s;
}

/**
 * @param {object} sa
 * @param {string} fallbackDatum ISO-Datum falls slot fehlt
 */
export function schularbeitGraphCalendarTimes(sa, fallbackDatum) {
    const datum = toIsoDateOnly(sa && sa.datum) || toIsoDateOnly(fallbackDatum);
    const slot = schularbeitTerminZeitfenster(sa);
    if (!datum) {
        return { isAllDay: true, startDateTime: '', endDateTime: '' };
    }
    if (!slot || slot.isAllDay) {
        const startDate = (slot && slot.startDate) || datum;
        const endDate = (slot && slot.endDate) || addDays(datum, 1) || datum;
        return {
            isAllDay: true,
            startDateTime: startDate + 'T00:00:00.0000000',
            endDateTime: endDate + 'T00:00:00.0000000'
        };
    }
    return {
        isAllDay: false,
        startDateTime: graphCalendarDateTime(slot.startDateTime),
        endDateTime: graphCalendarDateTime(slot.endDateTime)
    };
}

export function mondayOfWeekContaining(iso) {
    const base = toIsoDateOnly(iso) || toIsoDateOnly(new Date());
    const ms = utcNoonMs(base);
    if (Number.isNaN(ms)) return '';
    const dow = new Date(ms).getUTCDay();
    const diff = dow === 0 ? -6 : 1 - dow;
    return isoDateFromUtcMs(ms + diff * 86400000) || '';
}

/**
 * Unterrichtswoche Mo–Fr ab Montag.
 * @param {string} [mondayOrAnyDayInWeek] ISO; wird auf Montag normalisiert
 * @returns {{ monday: string, friday: string, days: { iso: string, weekday: string }[] }}
 */
export function buildSchoolWeekDays(mondayOrAnyDayInWeek) {
    const monday = mondayOfWeekContaining(mondayOrAnyDayInWeek);
    const weekdayLabels = ['Mo', 'Di', 'Mi', 'Do', 'Fr'];
    const days = [];
    for (let i = 0; i < 5; i++) {
        const iso = addDays(monday, i) || '';
        days.push({ iso, weekday: weekdayLabels[i] });
    }
    const friday = days[4] ? days[4].iso : addDays(monday, 4) || '';
    return { monday, friday, days };
}

/**
 * @param {string} iso
 * @param {Terminfenster} win
 */
export function dateInWindow(iso, win) {
    const d = toIsoDateOnly(iso);
    const a = toIsoDateOnly(win && win.startdatum);
    const b = toIsoDateOnly(win && win.enddatum);
    if (!d || !a || !b) return false;
    return d >= a && d <= b;
}

/**
 * @param {string} iso
 * @param {Terminfenster[]} windows
 * @param {FensterTyp} typ
 */
export function findWindowsCovering(iso, windows, typ) {
    const list = Array.isArray(windows) ? windows : [];
    return list.filter((w) => w && w.typ === typ && dateInWindow(iso, w));
}

/**
 * @param {Schularbeit} sa
 */
function isActiveStatus(sa) {
    const st = String((sa && sa.status) || 'beantragt').toLowerCase();
    return st === 'beantragt' || st === 'fixiert';
}

/**
 * @param {ValidateOptions} opts
 * @returns {{ errors: string[], warnings: string[], canSubmit: boolean }}
 */
/**
 * Anzeige-Titel in UI (Feld Titel, sonst Thema, sonst Fallback).
 * @param {{ titel?: string, thema?: string }} sa
 */
export function schularbeitDisplayTitle(sa) {
    const t = String((sa && sa.titel) || '').trim();
    if (t) return t;
    const th = String((sa && sa.thema) || '').trim();
    if (th) return th;
    return 'Schularbeit';
}

/**
 * SharePoint-Listeneintrag Title (nicht das Feld „Titel“).
 * @param {{ fachCode?: string, klasseCode?: string }} sa
 * @param {{ fachLabel?: string, klasseLabel?: string }} [labels]
 */
export function composeSchularbeitListItemTitle(sa, labels) {
    const fach =
        String((labels && labels.fachLabel) || (sa && sa.fachCode) || '').trim() || '?';
    const klasse =
        String((labels && labels.klasseLabel) || (sa && sa.klasseCode) || '').trim() || '?';
    return ('Schularbeit - ' + fach + ' - ' + klasse).slice(0, 250);
}

export function validateSchularbeit(opts) {
    const errors = [];
    const warnings = [];
    const draft = (opts && opts.draft) || {};
    const rules = { ...DEFAULT_RULES, ...(opts && opts.rules) };
    const windows = Array.isArray(opts && opts.windows) ? opts.windows : [];
    const existing = Array.isArray(opts && opts.existing) ? opts.existing : [];
    const fachMeta = (opts && opts.fachMeta) || {};
    const today = toIsoDateOnly(opts && opts.today) || toIsoDateOnly(new Date());

    const klasse = String(draft.klasseCode || '').trim();
    const fach = String(draft.fachCode || '').trim();
    const datum = toIsoDateOnly(draft.datum);
    const dauer = Number(draft.dauerMinuten);
    const draftId = String(draft.schularbeitId || '').trim();

    if (!klasse) errors.push('Bitte eine Klasse wählen.');
    if (!fach) errors.push('Bitte ein Fach wählen.');
    if (!datum) errors.push('Bitte ein gültiges Datum (JJJJ-MM-TT) angeben.');
    const beginnUhrzeit = normalizeBeginnUhrzeit(draft.beginnUhrzeit);
    if (!beginnUhrzeit) errors.push('Bitte eine gültige Beginn-Uhrzeit (HH:MM) angeben.');
    if (!Number.isFinite(dauer) || dauer <= 0) errors.push('Bitte eine gültige Dauer in Minuten angeben.');

    if (errors.length) {
        return { errors, warnings, canSubmit: false };
    }

    const peers = existing.filter((sa) => {
        if (!sa || !isActiveStatus(sa)) return false;
        if (String(sa.klasseCode || '').trim() !== klasse) return false;
        const id = String(sa.schularbeitId || '').trim();
        if (draftId && id && id === draftId) return false;
        return true;
    });

    const sameDay = peers.filter((sa) => toIsoDateOnly(sa.datum) === datum);
    const maxTag = Math.max(1, Number(rules.maxProTag) || 1);
    if (sameDay.length >= maxTag) {
        errors.push(
            `Maximal ${maxTag} Schularbeit${maxTag === 1 ? '' : 'en'} pro Tag und Klasse (bereits ${sameDay.length}).`
        );
    }

    const week = isoWeekKey(datum);
    const sameWeek = peers.filter((sa) => isoWeekKey(sa.datum) === week);
    const maxWoche = Math.max(1, Number(rules.maxProWoche) || 1);
    if (sameWeek.length >= maxWoche) {
        errors.push(
            `Maximal ${maxWoche} Schularbeit${maxWoche === 1 ? '' : 'en'} pro Kalenderwoche und Klasse (bereits ${sameWeek.length} in ${week}).`
        );
    }

    const frist = Math.max(0, Number(rules.ankuendigungsfristTage) || 0);
    const lead = daysBetween(today, datum);
    if (lead != null && lead < frist) {
        errors.push(`Ankündigungsfrist: mindestens ${frist} Tage Vorlauf (aktuell ${lead} Tag${lead === 1 ? '' : 'e'}).`);
    }
    if (lead != null && lead < 0) {
        errors.push('Das Datum liegt in der Vergangenheit.');
    }

    const blocked = findWindowsCovering(datum, windows, 'gesperrt');
    if (blocked.length) {
        const titles = blocked.map((w) => w.titel || 'Sperrzeit').join(', ');
        errors.push(`Termin liegt in einer Sperrzeit: ${titles}.`);
    }

    const prev = addDays(datum, -1);
    if (prev) {
        const prevBlocked = findWindowsCovering(prev, windows, 'gesperrt');
        if (prevBlocked.length) {
            warnings.push(
                'Tag nach schulfreiem Zeitraum (Sperrzeit) – laut LBVO möglichst vermeiden.'
            );
        }
    }

    const proSem = Number(fachMeta.proSemester);
    if (Number.isFinite(proSem) && proSem > 0 && draft.semester) {
        const sem = String(draft.semester).toUpperCase();
        const sameFachSem = peers.filter(
            (sa) =>
                String(sa.fachCode || '').trim() === fach &&
                String(sa.semester || '').toUpperCase() === sem
        );
        if (sameFachSem.length >= proSem) {
            warnings.push(
                `Kontingent Fach/Semester: üblich max. ${proSem} (bereits ${sameFachSem.length}).`
            );
        }
    }

    const minD = Number(fachMeta.minDauer) || 50;
    const maxD = Number(fachMeta.maxDauer) || 150;
    if (Number.isFinite(dauer) && (dauer < minD || dauer > maxD)) {
        warnings.push(`Dauer üblicherweise ${minD}–${maxD} Minuten (aktuell ${dauer}).`);
    }

    return {
        errors,
        warnings,
        canSubmit: errors.length === 0
    };
}

/**
 * KPI-Helfer für Dashboard.
 * @param {Schularbeit[]} items
 * @param {string} [todayIso]
 * @param {Regelwerk} [rules]
 */
export function computeDashboardKpis(items, todayIso, rules) {
    const today = toIsoDateOnly(todayIso) || toIsoDateOnly(new Date());
    const list = Array.isArray(items) ? items : [];
    const inTwoWeeks = addDays(today, 14);
    const rw = { ...DEFAULT_RULES, ...(rules || {}) };

    let offen = 0;
    let fixiertNaechste2Wochen = 0;
    let dieseWoche = 0;
    const weekNow = isoWeekKey(today);

    list.forEach((sa) => {
        const st = String((sa && sa.status) || '').toLowerCase();
        const d = toIsoDateOnly(sa && sa.datum);
        if (st === 'beantragt') offen++;
        if (st === 'fixiert' && d && today && inTwoWeeks && d >= today && d <= inTwoWeeks) {
            fixiertNaechste2Wochen++;
        }
        if ((st === 'beantragt' || st === 'fixiert') && d && isoWeekKey(d) === weekNow) {
            dieseWoche++;
        }
    });

    return {
        offen,
        fixiertNaechste2Wochen,
        konflikte: countRuleConflicts(list, rw),
        dieseWoche
    };
}

/**
 * Anzahl Konflikt-Gruppen (Klasse+Tag oder Klasse+Woche über Limit).
 * @param {Schularbeit[]} items
 * @param {Regelwerk} rules
 */
export function countRuleConflicts(items, rules) {
    const rw = { ...DEFAULT_RULES, ...(rules || {}) };
    const maxTag = Math.max(1, Number(rw.maxProTag) || 1);
    const maxWoche = Math.max(1, Number(rw.maxProWoche) || 1);
    const active = (Array.isArray(items) ? items : []).filter((sa) => {
        const st = String((sa && sa.status) || '').toLowerCase();
        return st === 'beantragt' || st === 'fixiert';
    });

    const byDay = new Map();
    const byWeek = new Map();
    active.forEach((sa) => {
        const klasse = String(sa.klasseCode || '').trim();
        const d = toIsoDateOnly(sa.datum);
        if (!klasse || !d) return;
        const dayKey = klasse + '|' + d;
        const weekKey = klasse + '|' + isoWeekKey(d);
        byDay.set(dayKey, (byDay.get(dayKey) || 0) + 1);
        byWeek.set(weekKey, (byWeek.get(weekKey) || 0) + 1);
    });

    let n = 0;
    byDay.forEach((count) => {
        if (count > maxTag) n++;
    });
    byWeek.forEach((count) => {
        if (count > maxWoche) n++;
    });
    return n;
}

/**
 * Wochenverteilung für Dashboard-Chart.
 * @param {Schularbeit[]} items
 * @param {{ today?: string, weekCount?: number }} [opts]
 * @returns {{ weeks: { key: string, label: string, total: number, byFach: Record<string, number> }[], faecher: string[] }}
 */
export function buildWeeklyDistribution(items, opts) {
    const today = toIsoDateOnly(opts && opts.today) || toIsoDateOnly(new Date());
    const weekCount = Math.max(1, Number(opts && opts.weekCount) || 8);
    const active = (Array.isArray(items) ? items : []).filter((sa) => {
        const st = String((sa && sa.status) || '').toLowerCase();
        return (st === 'beantragt' || st === 'fixiert') && toIsoDateOnly(sa.datum);
    });

    // Anker: Montag der aktuellen ISO-Woche
    const anchorMonday = mondayOfIsoWeek(today);
    const weekKeys = [];
    for (let i = 0; i < weekCount; i++) {
        const monday = addDays(anchorMonday, i * 7);
        const key = isoWeekKey(monday);
        weekKeys.push({ key, monday, label: key.replace(/^\d{4}-/, '') });
    }
    const keySet = new Set(weekKeys.map((w) => w.key));

    /** @type {Map<string, Record<string, number>>} */
    const map = new Map();
    weekKeys.forEach((w) => map.set(w.key, {}));

    const fachSet = new Set();
    active.forEach((sa) => {
        const key = isoWeekKey(sa.datum);
        if (!keySet.has(key)) return;
        const fach = String(sa.fachCode || '').trim() || '?';
        fachSet.add(fach);
        const bucket = map.get(key) || {};
        bucket[fach] = (bucket[fach] || 0) + 1;
        map.set(key, bucket);
    });

    const faecher = Array.from(fachSet).sort((a, b) => a.localeCompare(b, 'de'));
    const weeks = weekKeys.map((w) => {
        const byFach = map.get(w.key) || {};
        let total = 0;
        Object.keys(byFach).forEach((f) => {
            total += byFach[f];
        });
        return { key: w.key, label: w.label, total, byFach };
    });

    return { weeks, faecher };
}

/**
 * @param {string} iso
 */
function mondayOfIsoWeek(iso) {
    const p = parseIsoDateParts(iso);
    if (!p) return iso;
    const dt = new Date(Date.UTC(p.y, p.m - 1, p.d, 12));
    const day = dt.getUTCDay() || 7; // So=7
    dt.setUTCDate(dt.getUTCDate() - (day - 1));
    return toIsoDateOnly(dt);
}
