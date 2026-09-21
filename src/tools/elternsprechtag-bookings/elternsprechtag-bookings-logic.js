/**
 * Reine Hilfsfunktionen für Elternsprechtag / Microsoft Bookings (ohne DOM/Graph).
 */

const WEEKDAYS = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
const WEEKDAY_DE = {
    sunday: 'Sonntag',
    monday: 'Montag',
    tuesday: 'Dienstag',
    wednesday: 'Mittwoch',
    thursday: 'Donnerstag',
    friday: 'Freitag',
    saturday: 'Samstag'
};

/** @param {string} isoDate YYYY-MM-DD */
export function weekdayFromIsoDate(isoDate) {
    const m = String(isoDate || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    const y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);
    const dt = new Date(y, mo - 1, d, 12, 0, 0);
    if (Number.isNaN(dt.getTime())) return null;
    return WEEKDAYS[dt.getDay()] || null;
}

export function weekdayLabelDe(day) {
    return WEEKDAY_DE[day] || String(day || '');
}

/** @param {string} hhmm "15:00" oder "15:00:00" */
export function toBookingsTime(hhmm) {
    const s = String(hhmm || '').trim();
    const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) return null;
    const h = Math.min(23, Math.max(0, Number(m[1])));
    const min = Math.min(59, Math.max(0, Number(m[2])));
    const sec = m[3] != null ? Math.min(59, Math.max(0, Number(m[3]))) : 0;
    const pad = function (n) {
        return String(n).padStart(2, '0');
    };
    return pad(h) + ':' + pad(min) + ':' + pad(sec) + '.0000000';
}

/** @param {number} minutes */
export function durationIsoFromMinutes(minutes) {
    const n = Math.round(Number(minutes));
    if (!Number.isFinite(n) || n < 5 || n > 120) return null;
    return 'PT' + n + 'M';
}

/**
 * Wochenmuster für einen Wochentag (Hilfsbau für customAvailabilities / businessHours).
 * @param {{ weekday: string, startTime: string, endTime: string }} opts
 */
export function buildBusinessHoursForDay(opts) {
    const day = String((opts && opts.weekday) || '').toLowerCase();
    const start = opts && opts.startTime;
    const end = opts && opts.endTime;
    if (!WEEKDAYS.includes(day) || !start || !end) return null;
    return WEEKDAYS.map(function (wd) {
        if (wd !== day) {
            return { day: wd, timeSlots: [] };
        }
        return {
            day: wd,
            timeSlots: [{ startTime: start, endTime: end }]
        };
    });
}

/**
 * Kalendertage zwischen zwei ISO-Daten (to − from), ganzzahlig.
 * @param {string} fromIso
 * @param {string} toIso
 */
export function calendarDaysBetween(fromIso, toIso) {
    const a = String(fromIso || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
    const b = String(toIso || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!a || !b) return null;
    const d0 = new Date(Number(a[1]), Number(a[2]) - 1, Number(a[3]), 12, 0, 0);
    const d1 = new Date(Number(b[1]), Number(b[2]) - 1, Number(b[3]), 12, 0, 0);
    if (Number.isNaN(d0.getTime()) || Number.isNaN(d1.getTime())) return null;
    return Math.round((d1.getTime() - d0.getTime()) / 86400000);
}

/**
 * maximumAdvance so, dass der Sprechtag erst ab bookingOpenDate buchbar wird.
 * @param {string} bookingOpenIso
 * @param {string} eventIso
 */
export function maximumAdvanceForOpenDate(bookingOpenIso, eventIso) {
    const days = calendarDaysBetween(bookingOpenIso, eventIso);
    if (days == null) return null;
    if (days < 0) return null;
    return 'P' + Math.max(0, Math.min(365, days)) + 'D';
}

/**
 * Dienst-Scheduling: generell nicht buchbar, nur am konkreten Sprechtag.
 * Freischaltung über maximumAdvance (ab bookingOpenDate).
 * @param {{ eventDate: string, startHhmm: string, endHhmm: string, durationMin: number, bookingOpenDate: string }} opts
 */
export function buildSingleDayServiceSchedulingPolicy(opts) {
    const eventDate = String((opts && opts.eventDate) || '').trim();
    const openDate = String((opts && opts.bookingOpenDate) || '').trim();
    const weekday = weekdayFromIsoDate(eventDate);
    const startTime = toBookingsTime(opts && opts.startHhmm);
    const endTime = toBookingsTime(opts && opts.endHhmm);
    const duration = durationIsoFromMinutes(opts && opts.durationMin);
    const maxAdvance = maximumAdvanceForOpenDate(openDate || eventDate, eventDate);
    if (!weekday || !startTime || !endTime || !duration || !maxAdvance) return null;
    const hours = buildBusinessHoursForDay({
        weekday: weekday,
        startTime: startTime,
        endTime: endTime
    });
    if (!hours) return null;
    return {
        allowStaffSelection: true,
        timeSlotInterval: duration,
        minimumLeadTime: 'PT30M',
        maximumAdvance: maxAdvance,
        sendConfirmationsToOwner: true,
        isMeetingInviteToCustomersEnabled: true,
        generalAvailability: {
            availabilityType: 'notBookable'
        },
        customAvailabilities: [
            {
                '@odata.type': '#microsoft.graph.bookingsAvailabilityWindow',
                availabilityType: 'customWeeklyHours',
                startDate: eventDate,
                endDate: eventDate,
                businessHours: hours
            }
        ]
    };
}

/**
 * @param {{ teachers: Array<{ code?: string, name?: string, email?: string }>, selectedEmails?: Set<string>|null }} input
 */
export function normalizeTeacherRows(input) {
    const list = Array.isArray(input && input.teachers) ? input.teachers : [];
    const selected = input && input.selectedEmails instanceof Set ? input.selectedEmails : null;
    return list
        .map(function (t) {
            const email = String((t && t.email) || '')
                .trim()
                .toLowerCase();
            const code = String((t && t.code) || '')
                .trim()
                .toUpperCase();
            const name = String((t && t.name) || '').trim() || code || email;
            if (!email) {
                return {
                    code: code,
                    name: name,
                    email: '',
                    selected: false,
                    skipReason: 'keine E-Mail'
                };
            }
            return {
                code: code,
                name: name,
                email: email,
                selected: selected ? selected.has(email) : true,
                skipReason: ''
            };
        })
        .filter(function (t) {
            return t.email || t.code || t.name;
        })
        .sort(function (a, b) {
            return String(a.name || a.code).localeCompare(String(b.name || b.code), 'de');
        });
}

/**
 * @param {string} a
 * @param {string} b
 */
export function emailsEqual(a, b) {
    return String(a || '')
        .trim()
        .toLowerCase() ===
        String(b || '')
            .trim()
            .toLowerCase();
}

/**
 * Tage von heute bis inkl. Ereignistag (mindestens 1).
 * @param {string} isoDate
 */
export function daysUntilInclusive(isoDate) {
    const m = String(isoDate || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return 30;
    const target = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 23, 59, 0);
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
    const diff = Math.ceil((target.getTime() - start.getTime()) / 86400000);
    return Math.max(1, Math.min(365, diff));
}
