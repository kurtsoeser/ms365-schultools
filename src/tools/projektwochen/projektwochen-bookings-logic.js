/**
 * Reine Bookings-Hilfen für Projektwochen (ohne DOM/Graph).
 * Nutzt Muster aus elternsprechtag-bookings-logic (maximumAdvance, Einzeltag).
 */
import {
    weekdayFromIsoDate,
    toBookingsTime,
    maximumAdvanceForOpenDate,
    buildBusinessHoursForDay,
    calendarDaysBetween
} from '../elternsprechtag-bookings/elternsprechtag-bookings-logic.js';
import { toIsoDateOnly, effectiveBuchungAb } from './projektwochen-logic.js';

/**
 * Minuten zwischen HH:mm – für Ganztagsangebote bis 10h.
 * @param {string} startHhmm
 * @param {string} endHhmm
 */
export function durationMinutesFromTimes(startHhmm, endHhmm) {
    const a = String(startHhmm || '').match(/^(\d{1,2}):(\d{2})/);
    const b = String(endHhmm || '').match(/^(\d{1,2}):(\d{2})/);
    if (!a || !b) return null;
    const m0 = Number(a[1]) * 60 + Number(a[2]);
    const m1 = Number(b[1]) * 60 + Number(b[2]);
    const d = m1 - m0;
    if (!Number.isFinite(d) || d < 5) return null;
    return Math.min(600, d);
}

/**
 * @param {number} minutes
 * @returns {string|null} ISO-8601 Dauer
 */
export function durationIsoFromMinutesPw(minutes) {
    const n = Math.round(Number(minutes));
    if (!Number.isFinite(n) || n < 5 || n > 600) return null;
    if (n % 60 === 0 && n >= 60) return 'PT' + n / 60 + 'H';
    return 'PT' + n + 'M';
}

/**
 * BuchungAb (dateTime) → YYYY-MM-DD für maximumAdvance.
 * @param {string} buchungAb
 * @param {string} eventIso
 */
export function bookingOpenDateFromBuchungAb(buchungAb, eventIso) {
    const d = toIsoDateOnly(buchungAb);
    if (d) return d;
    return toIsoDateOnly(eventIso) || '';
}

/**
 * Scheduling-Policy: ein Tag, Kapazität über Service.maximumAttendeesCount.
 * @param {{ eventDate: string, startHhmm: string, endHhmm: string, bookingOpenDate: string }} opts
 */
export function buildPwServiceSchedulingPolicy(opts) {
    const eventDate = toIsoDateOnly(opts && opts.eventDate);
    const openDate = bookingOpenDateFromBuchungAb(opts && opts.bookingOpenDate, eventDate);
    const weekday = weekdayFromIsoDate(eventDate);
    const startTime = toBookingsTime(opts && opts.startHhmm);
    const endTime = toBookingsTime(opts && opts.endHhmm);
    const mins = durationMinutesFromTimes(opts && opts.startHhmm, opts && opts.endHhmm);
    const duration = durationIsoFromMinutesPw(mins);
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
        minimumLeadTime: 'PT0M',
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
 * Dienst-Payload aus Angebot + Aktion.
 * @param {object} angebot
 * @param {object|null} aktion
 */
export function buildServicePayloadFromAngebot(angebot, aktion) {
    const eventDate = toIsoDateOnly(angebot && angebot.datum);
    const start = String((angebot && angebot.startzeit) || '08:00');
    const end = String((angebot && angebot.endzeit) || '12:00');
    const openRaw = effectiveBuchungAb(angebot, aktion);
    const openDate = bookingOpenDateFromBuchungAb(openRaw, eventDate);
    const mins = durationMinutesFromTimes(start, end) || 240;
    const duration = durationIsoFromMinutesPw(mins);
    const policy = buildPwServiceSchedulingPolicy({
        eventDate,
        startHhmm: start,
        endHhmm: end,
        bookingOpenDate: openDate
    });
    if (!eventDate || !duration || !policy) {
        return { ok: false, error: 'Scheduling ungültig (Datum/Zeiten/BuchungAb prüfen).', payload: null };
    }

    const kap = Math.max(1, Math.min(100, Number(angebot.kapazitaet) || 1));
    const preis = Number(angebot.preisEuro);
    const free = !Number.isFinite(preis) || preis <= 0;
    const descParts = [
        String((angebot && angebot.beschreibung) || '').trim(),
        angebot.ort ? 'Ort: ' + angebot.ort : '',
        angebot.treffpunkt ? 'Treffpunkt: ' + angebot.treffpunkt : '',
        !free ? 'Preis: ' + preis + ' €' + (angebot.kostenHinweis ? ' (' + angebot.kostenHinweis + ')' : '') : '',
        angebot.hinweisEltern ? 'Eltern: ' + angebot.hinweisEltern : '',
        angebot.angebotId ? '[PW:' + angebot.angebotId + ']' : ''
    ].filter(Boolean);

    return {
        ok: true,
        error: '',
        duration,
        openDate,
        maxAdvance: policy.maximumAdvance,
        payload: {
            displayName: String((angebot && angebot.title) || 'Angebot').slice(0, 100),
            description: descParts.join('\n').slice(0, 1500),
            defaultDuration: duration,
            defaultPrice: free ? 0 : preis,
            defaultPriceType: free ? 'free' : 'fixedPrice',
            isHiddenFromCustomers: false,
            maximumAttendeesCount: kap,
            schedulingPolicy: policy
        }
    };
}

/**
 * Vorschlag Business-Anzeigename.
 * @param {object} aktion
 */
export function defaultBusinessName(aktion) {
    const t = String((aktion && aktion.title) || '').trim();
    if (t) return t.slice(0, 80);
    const id = String((aktion && aktion.aktionId) || 'pw').trim();
    return 'Projektwoche ' + id;
}

/**
 * Appointments → Teilnehmerzeilen + Belegung je Service.
 * @param {object[]} appointments
 * @param {object[]} angebote  mit bookingsServiceId
 */
export function normalizeAppointments(appointments, angebote) {
    const byService = {};
    ((angebote || []) || []).forEach((a) => {
        const sid = String((a && a.bookingsServiceId) || '').trim();
        if (sid) byService[sid] = a;
    });

    /** @type {object[]} */
    const rows = [];
    /** @type {Record<string, { filled: number, max: number, angebotId: string, title: string }>} */
    const occupancy = {};

    const list = Array.isArray(appointments) ? appointments : [];
    for (let i = 0; i < list.length; i++) {
        const ap = list[i];
        if (!ap) continue;
        const serviceId = String(ap.serviceId || '').trim();
        const angebot = byService[serviceId] || null;
        const max = Number(ap.maximumAttendeesCount) || (angebot && angebot.kapazitaet) || 0;
        const filled =
            Number(ap.filledAttendeesCount) >= 0
                ? Number(ap.filledAttendeesCount)
                : Array.isArray(ap.customers)
                  ? ap.customers.length
                  : 0;

        if (serviceId) {
            if (!occupancy[serviceId]) {
                occupancy[serviceId] = {
                    filled: 0,
                    max: max,
                    angebotId: angebot ? angebot.angebotId : '',
                    title: angebot ? angebot.title : String(ap.serviceName || serviceId)
                };
            }
            occupancy[serviceId].filled = Math.max(occupancy[serviceId].filled, filled);
            if (max) occupancy[serviceId].max = max;
        }

        const customers = Array.isArray(ap.customers) ? ap.customers : [];
        if (!customers.length && (ap.customerName || ap.customerEmailAddress)) {
            customers.push({
                name: ap.customerName || '',
                emailAddress: ap.customerEmailAddress || ''
            });
        }
        for (let c = 0; c < customers.length; c++) {
            const cust = customers[c] || {};
            rows.push({
                appointmentId: String(ap.id || ''),
                serviceId,
                angebotId: angebot ? angebot.angebotId : '',
                angebotTitle: angebot ? angebot.title : String(ap.serviceName || ''),
                datum: angebot ? angebot.datum : toIsoDateOnly(ap.start && ap.start.dateTime),
                name: String(cust.name || '').trim(),
                email: String(cust.emailAddress || cust.email || '')
                    .trim()
                    .toLowerCase(),
                phone: String(cust.phone || '').trim()
            });
        }
    }

    return { rows, occupancy };
}

/**
 * Join Klasse aus Stammdaten-Schülern.
 * @param {object[]} rows
 * @param {object[]} students
 */
export function enrichAttendeesWithKlasse(rows, students) {
    const byMail = {};
    (students || []).forEach((s) => {
        const em = String((s && s.email) || '')
            .trim()
            .toLowerCase();
        if (em) byMail[em] = s.klasse || s.klasseCode || s.classCode || '';
    });
    return (rows || []).map((r) => ({
        ...r,
        klasse: byMail[r.email] || ''
    }));
}

/**
 * @param {object[]} rows
 * @returns {string} CSV
 */
export function attendeesToCsv(rows) {
    const header = ['Name', 'E-Mail', 'Klasse', 'Angebot', 'Datum', 'ServiceId'];
    const lines = [header.join(';')];
    (rows || []).forEach((r) => {
        const cells = [r.name, r.email, r.klasse, r.angebotTitle, r.datum, r.serviceId].map((v) => {
            const s = String(v == null ? '' : v);
            if (/[;"\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
            return s;
        });
        lines.push(cells.join(';'));
    });
    return lines.join('\r\n');
}

export {
    weekdayFromIsoDate,
    toBookingsTime,
    maximumAdvanceForOpenDate,
    calendarDaysBetween
};
