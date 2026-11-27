/**
 * Tagesansicht Elternsprechtag für die angemeldete Lehrkraft.
 */

import {
    dayHeading,
    readSetupFromStorage,
    slotsFromEvents
} from './elternsprechtag-logic.js';
import { loadMyDayEvents } from './elternsprechtag-calendar.js';
import { focusBooking, pickConferenceDate, presentDay } from './elternsprechtag-day.js';
import { loadSpiegelSlots } from './elternsprechtag-spiegel.js';
import { resolveSpiegelSiteUrl, serviceWindow } from './elternsprechtag-spiegel-logic.js';

/**
 * @param {string} value
 */
function esc(value) {
    return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function signedInEmail() {
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            return String((info && (info.username || info.upn)) || '')
                .trim()
                .toLowerCase();
        }
    } catch {
        /* ignore */
    }
    return '';
}

function todayIso() {
    const d = new Date();
    const pad = function (n) {
        return String(n).padStart(2, '0');
    };
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
}

/**
 * @param {HTMLElement} host
 */
export function mountElternsprechtag(host) {
    const setup = readSetupFromStorage(typeof localStorage !== 'undefined' ? localStorage : null);
    const demo = /(?:\?|&)demo=1(?:&|$)/.test(String(location.search || ''));
    let date = setup.eventDate || todayIso();
    let datePicked = !!setup.eventDate;
    let includeOther = false;
    let moreOpen = false;
    let slots = [];
    let knownDates = [];
    let updatedAt = '';
    let status = '';
    let statusTone = '';
    let loading = false;
    let source = 'kalender';

    const paint = function () {
        const heading = dayHeading(date);
        const booked = slots.filter(function (slot) {
            return slot && !slot.free;
        }).length;
        const focus = focusBooking(slots, { isToday: date === todayIso(), nowMinutes: nowMinutes() });
        const rows = slots
            .map(function (slot) {
                const who = slot.free ? 'frei' : slot.parentName || slot.subject || 'Termin';
                const childLine = [slot.childName, slot.className].filter(Boolean).join(', ');
                const meta = [childLine, slot.note, slot.place].filter(Boolean).join(' · ');
                const current = focus && focus.slot && focus.slot.id === slot.id;
                return (
                    '<li class="est-slot' +
                    (slot.free ? ' est-slot--free' : '') +
                    (current ? ' est-slot--now' : '') +
                    '">' +
                    '<div class="est-slot__time">' +
                    esc(slot.startLabel) +
                    (slot.endLabel ? '<span>–' + esc(slot.endLabel) + '</span>' : '') +
                    '</div>' +
                    '<div class="est-slot__body">' +
                    '<strong>' +
                    esc(who) +
                    '</strong>' +
                    (meta ? '<p>' + esc(meta) + '</p>' : '') +
                    '</div>' +
                    primaryAction(slot) +
                    '</li>'
                );
            })
            .join('');
        const otherDates = knownDates.filter(function (day) {
            return day !== date;
        });
        const dateButtons = otherDates
            .map(function (day) {
                return '<button type="button" class="est-date" data-date="' + esc(day) + '">' + esc(shortDate(day)) + '</button>';
            })
            .join('');
        const empty = !loading && !slots.length
            ? '<p class="est-empty">Am ' +
              esc(heading || 'diesem Tag') +
              ' ist nichts gebucht.' +
              (dateButtons ? ' Andere übernommene Tage stehen unter dem Datum.' : '') +
              '</p>'
            : '';
        const banner = focus
            ? '<p class="est-next"><span>' +
              esc(focusLabel(focus.kind)) +
              '</span> ' +
              esc(focus.slot.startLabel || '') +
              ' · ' +
              esc(focus.slot.parentName || focus.slot.subject || 'Termin') +
              (focus.slot.childName || focus.slot.className
                  ? ' · ' + esc([focus.slot.childName, focus.slot.className].filter(Boolean).join(', '))
                  : '') +
              '</p>'
            : '';
        host.innerHTML =
            '<section class="est">' +
            '<header class="est__head">' +
            '<div><h2>' +
            esc(heading || 'Elternsprechtag') +
            '</h2></div>' +
            '<p class="est__count"><strong>' +
            String(booked) +
            '</strong><span>' +
            (booked === 1 ? 'Termin' : 'Termine') +
            '</span></p>' +
            '</header>' +
            banner +
            '<form class="est__bar">' +
            '<label>Datum <input type="date" name="date" value="' +
            esc(date) +
            '" required></label>' +
            (dateButtons ? '<div class="est-dates">' + dateButtons + '</div>' : '') +
            '<button type="button" class="est-retry"' +
            (loading ? ' disabled' : '') +
            '>' +
            (loading ? 'Laden …' : 'Erneut laden') +
            '</button>' +
            '</form>' +
            '<details class="est-more"' +
            (moreOpen ? ' open' : '') +
            '>' +
            '<summary>Termin fehlt?</summary>' +
            '<label class="est__check"><input type="checkbox" name="all"' +
            (includeOther ? ' checked' : '') +
            '> Auch andere Termine aus dem Kalender</label>' +
            '</details>' +
            (status
                ? '<p class="est-status' +
                  (statusTone === 'warn' ? ' est-status--warn' : '') +
                  '">' +
                  esc(status) +
                  '</p>'
                : '') +
            (rows ? '<ol class="est-day">' + rows + '</ol>' : empty) +
            '</section>';
        const form = host.querySelector('form');
        const dateInput = form && form.querySelector('input[name="date"]');
        if (dateInput) {
            dateInput.addEventListener('change', function () {
                date = String(dateInput.value || '').trim();
                datePicked = true;
                void reload();
            });
        }
        host.querySelectorAll('.est-date').forEach(function (button) {
            button.addEventListener('click', function () {
                date = button.getAttribute('data-date') || date;
                datePicked = true;
                void reload();
            });
        });
        const retry = host.querySelector('.est-retry');
        if (retry) {
            retry.addEventListener('click', function () {
                void reload();
            });
        }
        const more = host.querySelector('.est-more');
        const allBox = more && more.querySelector('input[name="all"]');
        if (more) {
            more.addEventListener('toggle', function () {
                moreOpen = more.open;
            });
        }
        if (allBox) {
            allBox.addEventListener('change', function () {
                includeOther = !!allBox.checked;
                moreOpen = true;
                void reload();
            });
        }
    };

    const reload = async function () {
        loading = true;
        status = '';
        paint();
        try {
            const directory = readDirectory();
            if (demo) {
                source = 'kalender';
                knownDates = [];
                updatedAt = '';
                const raw = slotsFromEvents(demoEvents(date), setup, 'lehrer@schule.at', {
                    includeOther: includeOther
                });
                const window = serviceWindow(setup) || { start: '15:00', end: '16:00', minutes: 10 };
                slots = finishDay(raw, window, directory, includeOther);
                status = 'Beispiel.';
            } else {
                const email = signedInEmail();
                const mirrored = await readMirror(date, email, setup);
                if (Array.isArray(mirrored.dates) && mirrored.dates.length) knownDates = mirrored.dates;
                if (!datePicked && knownDates.length) {
                    const better = pickConferenceDate(knownDates, todayIso());
                    datePicked = true;
                    if (better && better !== date) {
                        date = better;
                        loading = false;
                        return reload();
                    }
                }
                const window = mirrored.window || serviceWindow(setup);
                updatedAt = mirrored.updatedAt || '';
                if (mirrored.published) {
                    source = 'spiegel';
                    let raw = mirrored.slots;
                    if (includeOther) {
                        const events = await loadMyDayEvents(date);
                        raw = raw.concat(
                            slotsFromEvents(events, setup, email, { includeOther: true }).filter(function (slot) {
                                return !slot.booking;
                            })
                        );
                    }
                    slots = finishDay(raw, window, directory, includeOther);
                    const bookedNow = slots.filter(function (slot) {
                        return !slot.free;
                    }).length;
                    const when = takenLabel(updatedAt);
                    status = bookedNow
                        ? 'Übernommen' + (when ? ' ' + when : '') + '.'
                        : 'Übernommen' + (when ? ' ' + when : '') + '. Für dieses Konto keine Buchung.';
                } else {
                    source = 'kalender';
                    const events = await loadMyDayEvents(date);
                    const raw = slotsFromEvents(events, setup, email, {
                        includeOther: includeOther
                    });
                    slots = finishDay(raw, window, directory, includeOther);
                    status = mirrored.note || 'Noch aus dem eigenen Kalender.';
                }
            }
            statusTone = '';
        } catch (err) {
            slots = [];
            statusTone = 'warn';
            status = softError(err);
        }
        loading = false;
        paint();
    };

    paint();
    void reload();
}

/**
 * @param {string} isoDate
 * @param {string} email
 * @param {ReturnType<typeof import('./elternsprechtag-logic.js').readSetup>} setup
 */
async function readMirror(isoDate, email, setup) {
    let fallback = '';
    try {
        if (typeof window !== 'undefined' && window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
            const { suggestPrimarySiteUrl } = await import('../../shared/school-tool-sites.js');
            fallback = suggestPrimarySiteUrl(window.ms365AppDataV2.getSetup());
        }
    } catch {
        fallback = '';
    }
    const siteUrl = resolveSpiegelSiteUrl(setup, fallback);
    if (!siteUrl) return { published: false, slots: [], window: null, dates: [], updatedAt: '', note: '' };
    try {
        const loaded = await loadSpiegelSlots(siteUrl, isoDate, email);
        return {
            published: !!loaded.published,
            slots: loaded.slots || [],
            window: loaded.window || null,
            dates: loaded.dates || [],
            updatedAt: loaded.updatedAt || '',
            note: ''
        };
    } catch (err) {
        const raw = err && err.message ? String(err.message) : String(err || '');
        if (/403|AccessDenied|ErrorAccessDenied/i.test(raw)) {
            return {
                published: false,
                slots: [],
                window: null,
                dates: [],
                updatedAt: '',
                note: 'Liste nicht lesbar. Es gilt der eigene Kalender.'
            };
        }
        return { published: false, slots: [], window: null, dates: [], updatedAt: '', note: '' };
    }
}

function readDirectory() {
    try {
        const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
        if (!api || typeof api.getContainer !== 'function' || typeof api.getYearBucket !== 'function') {
            return { students: [], guardians: [] };
        }
        const container = api.getContainer() || {};
        const year = container.years && container.years.current;
        const bucket = api.getYearBucket(year) || {};
        return {
            students: bucket.students || [],
            guardians: bucket.guardians || []
        };
    } catch {
        return { students: [], guardians: [] };
    }
}

/**
 * @param {{ start: string, end: string, minutes: number }|null} window
 */
function nowMinutes() {
    const d = new Date();
    return d.getHours() * 60 + d.getMinutes();
}

function shortDate(iso) {
    const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return iso;
    return String(Number(m[3])) + '.' + String(Number(m[2])) + '.';
}

function focusLabel(kind) {
    if (kind === 'jetzt') return 'Jetzt';
    if (kind === 'naechster') return 'Nächster';
    if (kind === 'vorbei') return 'Letzter';
    return 'Erster Termin';
}

function takenLabel(iso) {
    const d = new Date(iso);
    if (!iso || Number.isNaN(d.getTime())) return '';
    let time = '';
    try {
        time = new Intl.DateTimeFormat('de-AT', {
            hour: '2-digit',
            minute: '2-digit',
            timeZone: 'Europe/Vienna'
        }).format(d);
    } catch {
        time = '';
    }
    const now = new Date();
    const same =
        d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth() && d.getDate() === now.getDate();
    if (same) return 'heute ' + time;
    try {
        return (
            new Intl.DateTimeFormat('de-AT', { day: 'numeric', month: 'short', timeZone: 'Europe/Vienna' }).format(d) +
            (time ? ' ' + time : '')
        );
    } catch {
        return time;
    }
}

function primaryAction(slot) {
    if (!slot || slot.free) return '';
    if (slot.joinUrl) {
        return (
            '<a class="est-action" href="' +
            esc(slot.joinUrl) +
            '" target="_blank" rel="noopener">Beitreten</a>'
        );
    }
    const phone = String(slot.phone || '').replace(/[^\d+]/g, '');
    if (phone) return '<a class="est-action" href="tel:' + esc(phone) + '">Anrufen</a>';
    if (slot.parentEmail) {
        return '<a class="est-action" href="mailto:' + esc(slot.parentEmail) + '">Mail</a>';
    }
    return '';
}

/**
 * @param {object[]} raw
 * @param {{ start: string, end: string, minutes: number }|null} window
 * @param {{ students?: object[], guardians?: object[] }} directory
 * @param {boolean} keepOutside
 */
function finishDay(raw, window, directory, keepOutside) {
    const bookings = (raw || []).filter(function (slot) {
        return slot && (slot.booking || !keepOutside);
    });
    const extras = keepOutside
        ? (raw || []).filter(function (slot) {
              return slot && !slot.booking;
          })
        : [];
    const grid = presentDay(bookings, window, directory);
    return grid.concat(extras).sort(function (a, b) {
        const byTime = String(a.startLabel || '').localeCompare(String(b.startLabel || ''));
        if (byTime) return byTime;
        return (a.free ? 1 : 0) - (b.free ? 1 : 0);
    });
}

function softError(err) {
    const raw = err && err.message ? String(err.message) : String(err || '');
    if (/login|anmeld|interaction|AADSTS|popup/i.test(raw)) {
        return 'Bitte oben anmelden, danach erneut aktualisieren.';
    }
    if (/403|AccessDenied|ErrorAccessDenied/i.test(raw)) {
        return 'Kein Zugriff auf den eigenen Kalender. Die Anmeldung braucht die Kalender-Berechtigung.';
    }
    return raw || 'Kalender konnte nicht geladen werden.';
}

/**
 * @param {string} isoDate
 */
function demoEvents(isoDate) {
    const day = isoDate || '2026-03-18';
    return [
        {
            id: '1',
            subject: 'Termin Elternsprechtag',
            isCancelled: false,
            start: { dateTime: day + 'T15:00:00', timeZone: 'Europe/Vienna' },
            end: { dateTime: day + 'T15:10:00', timeZone: 'Europe/Vienna' },
            organizer: { emailAddress: { address: 'sprechtag@bookings.local', name: 'Bookings' } },
            attendees: [
                { emailAddress: { address: 'lehrer@schule.at', name: 'Lehrkraft' } },
                { emailAddress: { address: 'anna.huber@mail.at', name: 'Anna Huber' } }
            ],
            bodyPreview: 'Name: Anna Huber\nKind: Lena Huber\nKlasse: 3A\nTelefon: +43 660 1234567',
            location: { displayName: 'Zimmer 12' },
            webLink: 'https://outlook.office.com/calendar/'
        },
        {
            id: '2',
            subject: 'Konferenz',
            isCancelled: false,
            start: { dateTime: day + 'T16:00:00', timeZone: 'Europe/Vienna' },
            end: { dateTime: day + 'T17:00:00', timeZone: 'Europe/Vienna' },
            organizer: { emailAddress: { address: 'direktion@schule.at', name: 'Direktion' } },
            bodyPreview: 'Interne Besprechung'
        }
    ];
}
