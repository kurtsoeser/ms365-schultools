/**
 * Buchungen aus Bookings als Zeilen für die Liste Elternsprechtag-Termine.
 * Bookings bleibt die Quelle. Ein Übernehmen-Lauf ersetzt den Tag.
 */

import { clockLabel, normEmail } from './elternsprechtag-logic.js';

export const ELTERNSPRECHTAG_LIST_TITLE = 'Elternsprechtag-Termine';

/** @type {import('../../shared/spo-graph-column.js').GraphColumnDef[]} */
export const ELTERNSPRECHTAG_LIST_COLUMNS = [
    { name: 'TerminId', displayName: 'Termin-ID', text: { allowMultipleLines: false, maxLength: 255 } },
    { name: 'Datum', displayName: 'Datum', text: { allowMultipleLines: false, maxLength: 10 } },
    { name: 'Startzeit', displayName: 'Start', text: { allowMultipleLines: false, maxLength: 8 } },
    { name: 'Endzeit', displayName: 'Ende', text: { allowMultipleLines: false, maxLength: 8 } },
    { name: 'Lehrkraft', displayName: 'Lehrkraft', text: { allowMultipleLines: false, maxLength: 255 } },
    { name: 'Eltern', displayName: 'Eltern', text: { allowMultipleLines: false, maxLength: 200 } },
    { name: 'ElternMail', displayName: 'E-Mail', text: { allowMultipleLines: false, maxLength: 255 } },
    { name: 'Telefon', displayName: 'Telefon', text: { allowMultipleLines: false, maxLength: 40 } },
    { name: 'Kind', displayName: 'Kind', text: { allowMultipleLines: false, maxLength: 120 } },
    { name: 'Klasse', displayName: 'Klasse', text: { allowMultipleLines: false, maxLength: 40 } },
    { name: 'Hinweis', displayName: 'Hinweis', text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'Ort', displayName: 'Ort', text: { allowMultipleLines: false, maxLength: 200 } },
    { name: 'OnlineLink', displayName: 'Online-Link', text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'FensterStart', displayName: 'Dienst von', text: { allowMultipleLines: false, maxLength: 8 } },
    { name: 'FensterEnde', displayName: 'Dienst bis', text: { allowMultipleLines: false, maxLength: 8 } },
    { name: 'SlotMinuten', displayName: 'Slotlänge', text: { allowMultipleLines: false, maxLength: 3 } }
];

/**
 * @param {unknown} setup
 * @param {string} [fallbackUrl]
 */
export function resolveSpiegelSiteUrl(setup, fallbackUrl) {
    const own = String(setup && setup.siteUrl ? setup.siteUrl : '').trim();
    if (/^https:\/\//i.test(own)) return own;
    const fb = String(fallbackUrl || '').trim();
    return /^https:\/\//i.test(fb) ? fb : '';
}

/**
 * @param {string} dateTime
 */
function isoDay(dateTime) {
    const m = String(dateTime || '').match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : '';
}

/**
 * @param {object[]} staff
 */
export function staffEmailById(staff) {
    /** @type {Record<string, string>} */
    const out = {};
    (staff || []).forEach(function (member) {
        const id = member && member.id ? String(member.id) : '';
        const email = normEmail(member && member.emailAddress);
        if (id && email) out[id] = email;
    });
    return out;
}

/**
 * @param {object|null|undefined} customer
 */
function answerLines(customer) {
    const raw =
        customer && (customer.customQuestionAnswers || customer.customQuestionAnswer);
    const list = Array.isArray(raw) ? raw : [];
    return list
        .map(function (entry) {
            const question = String((entry && (entry.question || entry.questionId)) || '').trim();
            const selected = Array.isArray(entry && entry.selectedOptions)
                ? entry.selectedOptions.join(', ')
                : '';
            const answer = String((entry && entry.answer) || selected || '').trim();
            return { question: question, answer: answer };
        })
        .filter(function (line) {
            return line.answer;
        });
}

/**
 * @param {string} question
 */
function answerKind(question) {
    const q = String(question || '').toLowerCase();
    if (/klasse/.test(q)) return 'klasse';
    if (/kind|schüler|schueler|sohn|tochter/.test(q)) return 'kind';
    return '';
}

/**
 * @param {object} appointment
 * @param {object} customer
 * @param {number} index
 */
function customerParts(appointment, customer, index) {
    const ap = appointment || {};
    const cust = customer || {};
    const lines = answerLines(cust);
    let kind = '';
    let klasse = '';
    const notes = [];
    lines.forEach(function (line) {
        const kindOf = answerKind(line.question);
        if (kindOf === 'kind' && !kind) kind = line.answer;
        else if (kindOf === 'klasse' && !klasse) klasse = line.answer;
        else notes.push((line.question ? line.question + ': ' : '') + line.answer);
    });
    const free = String(cust.notes || cust.customerNotes || ap.customerNotes || ap.additionalInformation || '').trim();
    if (free) notes.push(free);
    const name = String(cust.name || (index === 0 ? ap.customerName : '') || '').trim();
    const email = normEmail(cust.emailAddress || cust.email || (index === 0 ? ap.customerEmailAddress : ''));
    const phone = String(cust.phone || (index === 0 ? ap.customerPhone : '') || '').trim();
    return { name: name, email: email, phone: phone, kind: kind, klasse: klasse, hinweis: notes.join('\n') };
}

/**
 * @param {object} appointment
 */
function customersOf(appointment) {
    const ap = appointment || {};
    const list = Array.isArray(ap.customers) ? ap.customers.slice() : [];
    if (!list.length && (ap.customerName || ap.customerEmailAddress)) {
        list.push({
            name: ap.customerName || '',
            emailAddress: ap.customerEmailAddress || '',
            phone: ap.customerPhone || '',
            notes: ap.customerNotes || ''
        });
    }
    return list;
}

/**
 * @param {object[]} appointments
 * @param {object[]} staff
 * @param {{ eventDate?: string, serviceName?: string }} [opts]
 * @returns {{ rows: object[], skippedNoTeacher: number }}
 */
export function rowsFromAppointments(appointments, staff, opts) {
    const emails = staffEmailById(staff);
    const eventDate = String((opts && opts.eventDate) || '').trim();
    const serviceName = String((opts && opts.serviceName) || '').trim().toLowerCase();
    /** @type {object[]} */
    const rows = [];
    let skippedNoTeacher = 0;
    (Array.isArray(appointments) ? appointments : []).forEach(function (ap) {
        if (!ap || ap.isCancelled) return;
        const startRaw = ap.start && ap.start.dateTime ? String(ap.start.dateTime) : '';
        const endRaw = ap.end && ap.end.dateTime ? String(ap.end.dateTime) : '';
        const day = isoDay(startRaw);
        if (eventDate && day && day !== eventDate) return;
        const service = String(ap.serviceName || '').trim();
        if (serviceName && service && service.toLowerCase() !== serviceName) return;
        const customers = customersOf(ap);
        if (!customers.length) return;
        const staffIds = Array.isArray(ap.staffMemberIds) ? ap.staffMemberIds : [];
        const teachers = [];
        staffIds.forEach(function (id) {
            const email = emails[String(id || '')];
            if (email && teachers.indexOf(email) < 0) teachers.push(email);
        });
        if (!teachers.length) {
            skippedNoTeacher += customers.length;
            return;
        }
        const place =
            ap.serviceLocation && ap.serviceLocation.displayName
                ? String(ap.serviceLocation.displayName).trim()
                : '';
        const link = String(ap.joinWebUrl || '').trim();
        customers.forEach(function (cust, index) {
            const parts = customerParts(ap, cust, index);
            if (!parts.name && !parts.email) return;
            teachers.forEach(function (teacher) {
                const title = parts.name || parts.email || 'Buchung';
                rows.push({
                    terminId: String(ap.id || '') + '#' + String(index) + '#' + teacher,
                    datum: day || eventDate,
                    start: clockLabel(startRaw),
                    end: clockLabel(endRaw),
                    lehrkraft: teacher,
                    eltern: title,
                    email: parts.email,
                    phone: parts.phone,
                    kind: parts.kind,
                    klasse: parts.klasse,
                    hinweis: parts.hinweis,
                    ort: place,
                    link: link
                });
            });
        });
    });
    rows.sort(function (a, b) {
        return String(a.start).localeCompare(String(b.start)) || String(a.eltern).localeCompare(String(b.eltern));
    });
    return { rows: rows, skippedNoTeacher: skippedNoTeacher };
}

/**
 * @param {object} row
 */
export function fieldsForRow(row) {
    const r = row || {};
    /** @type {Record<string, string>} */
    const fields = {
        Title: String(r.eltern || 'Buchung').slice(0, 200),
        TerminId: String(r.terminId || '').slice(0, 255),
        Datum: String(r.datum || ''),
        Lehrkraft: String(r.lehrkraft || '')
    };
    function put(key, value) {
        const text = String(value || '').trim();
        if (text) fields[key] = text;
    }
    put('Startzeit', r.start);
    put('Endzeit', r.end);
    put('Eltern', r.eltern);
    put('ElternMail', r.email);
    put('Telefon', r.phone);
    put('Kind', r.kind);
    put('Klasse', r.klasse);
    put('Hinweis', r.hinweis);
    put('Ort', r.ort);
    put('OnlineLink', r.link);
    put('FensterStart', r.fensterStart);
    put('FensterEnde', r.fensterEnde);
    put('SlotMinuten', r.slotMinuten);
    return fields;
}

/**
 * @param {object} item
 * @param {string} eventDate
 */
export function itemMatchesDate(item, eventDate) {
    const fields = item && item.fields ? item.fields : {};
    return String(fields.Datum || '') === String(eventDate || '');
}

/**
 * @param {object[]} items
 * @param {string} eventDate
 */
export function itemIdsForDate(items, eventDate) {
    return (Array.isArray(items) ? items : [])
        .filter(function (item) {
            return item && item.id != null && itemMatchesDate(item, eventDate);
        })
        .map(function (item) {
            return String(item.id);
        });
}

/**
 * @param {object} fields
 */
function slotFromFields(fields, itemId) {
    const f = fields || {};
    const kind = String(f.Kind || '').trim();
    const klasse = String(f.Klasse || '').trim();
    const hinweis = String(f.Hinweis || '').trim();
    const detail = [];
    if (kind) detail.push('Kind: ' + kind);
    if (klasse) detail.push('Klasse: ' + klasse);
    if (hinweis) detail.push(hinweis);
    const link = String(f.OnlineLink || '').trim();
    const start = String(f.Startzeit || '').trim();
    return {
        id: String(f.TerminId || itemId || ''),
        startLabel: start,
        endLabel: String(f.Endzeit || '').trim(),
        sortKey: String(f.Datum || '') + 'T' + start,
        subject: String(f.Title || f.Eltern || '').trim(),
        parentName: String(f.Eltern || f.Title || '').trim(),
        parentEmail: normEmail(f.ElternMail),
        phone: String(f.Telefon || '').trim(),
        childName: kind,
        className: klasse,
        note: hinweis,
        detail: detail.join(' · '),
        place: String(f.Ort || '').trim(),
        webLink: '',
        joinUrl: link,
        online: !!link,
        booking: true
    };
}

/**
 * Eigene Zeilen eines bereits übernommenen Tages.
 * @param {object[]} items
 * @param {string} selfEmail
 * @param {string} eventDate
 */
export function slotsFromSpiegelItems(items, selfEmail, eventDate) {
    const me = normEmail(selfEmail);
    return (Array.isArray(items) ? items : [])
        .filter(function (item) {
            if (!itemMatchesDate(item, eventDate)) return false;
            const fields = item.fields || {};
            return me && normEmail(fields.Lehrkraft) === me;
        })
        .map(function (item) {
            return slotFromFields(item.fields, item.id);
        })
        .sort(function (a, b) {
            return String(a.sortKey).localeCompare(String(b.sortKey));
        });
}

/**
 * Dienstfenster nur mit gültiger Uhrzeit und Slotlänge.
 * @param {{ start?: string, end?: string, minutes?: number|string }|null|undefined} raw
 */
export function serviceWindow(raw) {
    const start = String((raw && raw.start) || '').trim();
    const end = String((raw && raw.end) || '').trim();
    const minutes = Number(raw && raw.minutes);
    if (!/^\d{2}:\d{2}$/.test(start) || !/^\d{2}:\d{2}$/.test(end)) return null;
    if (!(minutes >= 5 && minutes <= 120) || start >= end) return null;
    return { start: start, end: end, minutes: minutes };
}

/**
 * Schreibt das Fenster auf jede Zeile. Ohne Buchung bleibt eine Markerzeile, damit Lehrkräfte die freien Slots sehen.
 * @param {object[]} rows
 * @param {{ start?: string, end?: string, minutes?: number|string }|null|undefined} window
 * @param {string} eventDate
 */
export function withServiceWindow(rows, window, eventDate) {
    const w = serviceWindow(window);
    const list = (Array.isArray(rows) ? rows : []).slice();
    if (!w) return list;
    list.forEach(function (row) {
        row.fensterStart = w.start;
        row.fensterEnde = w.end;
        row.slotMinuten = String(w.minutes);
    });
    if (!list.length) {
        list.push({
            terminId: 'fenster#' + String(eventDate || ''),
            datum: String(eventDate || ''),
            lehrkraft: '',
            eltern: 'Dienstfenster',
            fensterStart: w.start,
            fensterEnde: w.end,
            slotMinuten: String(w.minutes)
        });
    }
    return list;
}

/**
 * @param {object[]} items
 * @param {string} eventDate
 */
export function windowFromItems(items, eventDate) {
    const list = Array.isArray(items) ? items : [];
    for (let i = 0; i < list.length; i++) {
        if (!itemMatchesDate(list[i], eventDate)) continue;
        const fields = list[i].fields || {};
        const window = serviceWindow({
            start: fields.FensterStart,
            end: fields.FensterEnde,
            minutes: fields.SlotMinuten
        });
        if (window) return window;
    }
    return null;
}

/**
 * Übernommene Tage, aufsteigend.
 * @param {object[]} items
 */
export function datesFromItems(items) {
    /** @type {string[]} */
    const out = [];
    (Array.isArray(items) ? items : []).forEach(function (item) {
        const fields = item && item.fields ? item.fields : {};
        const day = String(fields.Datum || '').trim();
        if (/^\d{4}-\d{2}-\d{2}$/.test(day) && out.indexOf(day) < 0) out.push(day);
    });
    out.sort();
    return out;
}

/**
 * @param {object[]} items
 * @param {string} eventDate
 */
export function latestUpdate(items, eventDate) {
    let best = '';
    (Array.isArray(items) ? items : []).forEach(function (item) {
        if (eventDate && !itemMatchesDate(item, eventDate)) return;
        const stamp = String((item && (item.lastModifiedDateTime || item.createdDateTime)) || '');
        if (stamp && stamp > best) best = stamp;
    });
    return best;
}

/**
 * Der Tag gilt als übernommen, sobald irgendeine Zeile dieses Datums in der Liste steht.
 * @param {object[]} items
 * @param {string} eventDate
 */
export function dayWasPublished(items, eventDate) {
    return (Array.isArray(items) ? items : []).some(function (item) {
        return itemMatchesDate(item, eventDate);
    });
}
