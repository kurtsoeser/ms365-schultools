/**
 * Tagesleiste: freie Slots im Dienstfenster, Klasse über die Eltern-Mail.
 */

import { normEmail } from './elternsprechtag-logic.js';

/**
 * @param {string} value
 */
function toMinutes(value) {
    const m = String(value || '').match(/^(\d{2}):(\d{2})$/);
    if (!m) return null;
    return Number(m[1]) * 60 + Number(m[2]);
}

/**
 * @param {number} minutes
 */
function fromMinutes(minutes) {
    const h = Math.floor(minutes / 60);
    const m = minutes % 60;
    return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
}

/**
 * @param {object[]} students
 * @param {object[]} guardians
 * @returns {Record<string, { name: string, klasse: string }[]>}
 */
export function childrenByParentEmail(students, guardians) {
    /** @type {Record<string, string>} */
    const emailById = {};
    (guardians || []).forEach(function (guardian) {
        const id = guardian && guardian.id ? String(guardian.id) : '';
        const email = normEmail(guardian && guardian.email);
        if (id && email) emailById[id] = email;
    });
    /** @type {Record<string, { name: string, klasse: string }[]>} */
    const out = {};
    function add(email, child) {
        const key = normEmail(email);
        if (!key || !child) return;
        if (!out[key]) out[key] = [];
        const name = String(child.name || '').trim();
        const klasse = String(child.klasse || '').trim();
        const known = out[key].some(function (row) {
            return row.name === name && row.klasse === klasse;
        });
        if (!known) out[key].push({ name: name, klasse: klasse });
    }
    (students || []).forEach(function (student) {
        if (!student) return;
        const child = { name: student.name, klasse: student.klasse || student.className || '' };
        (student.guardianIds || []).forEach(function (id) {
            add(emailById[String(id || '')], child);
        });
        const pairs = student.parentPairs || student.parents || [];
        (Array.isArray(pairs) ? pairs : []).forEach(function (pair) {
            add(pair && pair.email, child);
        });
    });
    return out;
}

/**
 * Stammdaten gewinnen gegenüber einer leeren oder abweichenden Klassenangabe aus dem Buchungstext.
 * @param {object} slot
 * @param {Record<string, { name: string, klasse: string }[]>} index
 */
export function applyParentClass(slot, index) {
    if (!slot || slot.free) return slot;
    const hits = (index && index[normEmail(slot.parentEmail)]) || [];
    if (!hits.length) return slot;
    const next = Object.assign({}, slot);
    const classes = [];
    const names = [];
    hits.forEach(function (hit) {
        if (hit.klasse && classes.indexOf(hit.klasse) < 0) classes.push(hit.klasse);
        if (hit.name && names.indexOf(hit.name) < 0) names.push(hit.name);
    });
    if (classes.length) next.className = classes.join(', ');
    if (!next.childName && names.length === 1) next.childName = names[0];
    return next;
}

/**
 * @param {object[]} slots
 * @param {{ start: string, end: string, minutes: number }|null|undefined} window
 */
export function fillFreeSlots(slots, window) {
    const start = window ? toMinutes(window.start) : null;
    const end = window ? toMinutes(window.end) : null;
    const step = window ? Number(window.minutes) : 0;
    const bookings = (Array.isArray(slots) ? slots : []).filter(function (slot) {
        return slot && !slot.free;
    });
    if (start == null || end == null || !(step >= 5 && step <= 120) || end <= start) {
        return bookings.slice();
    }
    /** @type {Record<string, object[]>} */
    const byStart = {};
    /** @type {object[]} */
    const outside = [];
    bookings.forEach(function (slot) {
        const at = toMinutes(slot.startLabel);
        if (at == null || at < start || at >= end) {
            outside.push(slot);
            return;
        }
        const key = fromMinutes(at);
        if (!byStart[key]) byStart[key] = [];
        byStart[key].push(slot);
    });
    /** @type {object[]} */
    const out = [];
    for (let t = start; t < end; t += step) {
        const key = fromMinutes(t);
        const hit = byStart[key];
        if (hit && hit.length) out.push.apply(out, hit);
        else {
            const slotEnd = Math.min(t + step, end);
            out.push({
                id: 'frei-' + key,
                free: true,
                startLabel: key,
                endLabel: fromMinutes(slotEnd),
                sortKey: 'T' + key,
                parentName: '',
                parentEmail: '',
                childName: '',
                className: '',
                detail: '',
                note: '',
                place: '',
                phone: '',
                webLink: '',
                joinUrl: '',
                online: false,
                booking: false
            });
        }
        delete byStart[key];
    }
    Object.keys(byStart).forEach(function (key) {
        outside.push.apply(outside, byStart[key]);
    });
    outside.sort(function (a, b) {
        return String(a.sortKey || a.startLabel).localeCompare(String(b.sortKey || b.startLabel));
    });
    return out.concat(outside);
}

/**
 * Aufeinanderfolgende freie Slots werden eine Zeile.
 * @param {object[]} slots
 */
export function collapseFreeRuns(slots) {
    /** @type {object[]} */
    const out = [];
    (Array.isArray(slots) ? slots : []).forEach(function (slot) {
        if (!slot) return;
        const prev = out[out.length - 1];
        if (slot.free && prev && prev.free) {
            prev.endLabel = slot.endLabel || prev.endLabel;
            return;
        }
        out.push(Object.assign({}, slot));
    });
    return out;
}

/**
 * Nächster, aktueller oder erster gebuchter Slot.
 * @param {object[]} slots
 * @param {{ isToday?: boolean, nowMinutes?: number }} [opts]
 */
export function focusBooking(slots, opts) {
    const parents = (Array.isArray(slots) ? slots : []).filter(function (slot) {
        return slot && !slot.free && slot.booking !== false && (slot.parentName || slot.subject);
    });
    if (!parents.length) return null;
    const o = opts || {};
    if (!o.isToday) return { kind: 'erster', slot: parents[0] };
    const now = Number(o.nowMinutes);
    if (!Number.isFinite(now)) return { kind: 'erster', slot: parents[0] };
    for (let i = 0; i < parents.length; i++) {
        const start = toMinutes(parents[i].startLabel);
        const end = toMinutes(parents[i].endLabel);
        if (start != null && end != null && now >= start && now < end) {
            return { kind: 'jetzt', slot: parents[i] };
        }
    }
    for (let i = 0; i < parents.length; i++) {
        const start = toMinutes(parents[i].startLabel);
        if (start != null && start > now) return { kind: 'naechster', slot: parents[i] };
    }
    return { kind: 'vorbei', slot: parents[parents.length - 1] };
}

/**
 * Heute, sonst der nächste übernommene Tag, sonst der letzte vergangene.
 * @param {string[]} dates
 * @param {string} today
 */
export function pickConferenceDate(dates, today) {
    const list = (Array.isArray(dates) ? dates : [])
        .map(function (value) {
            return String(value || '').trim();
        })
        .filter(function (value) {
            return /^\d{4}-\d{2}-\d{2}$/.test(value);
        })
        .sort();
    if (!list.length) return '';
    const day = String(today || '');
    if (list.indexOf(day) >= 0) return day;
    for (let i = 0; i < list.length; i++) {
        if (list[i] > day) return list[i];
    }
    return list[list.length - 1];
}

/**
 * @param {object[]} slots
 * @param {{ start: string, end: string, minutes: number }|null|undefined} window
 * @param {{ students?: object[], guardians?: object[] }} [directory]
 */
export function presentDay(slots, window, directory) {
    const dir = directory || {};
    const index = childrenByParentEmail(dir.students, dir.guardians);
    const matched = (Array.isArray(slots) ? slots : []).map(function (slot) {
        return applyParentClass(slot, index);
    });
    return collapseFreeRuns(fillFreeSlots(matched, window));
}
