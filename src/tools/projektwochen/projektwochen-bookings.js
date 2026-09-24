/**
 * Microsoft Bookings Graph-API für Projektwochen.
 */
import {
    buildServicePayloadFromAngebot,
    defaultBusinessName,
    normalizeAppointments,
    enrichAttendeesWithKlasse,
    toBookingsTime,
    weekdayFromIsoDate,
    buildPwServiceSchedulingPolicy
} from './projektwochen-bookings-logic.js';
import { buildBusinessHoursForDay } from '../elternsprechtag-bookings/elternsprechtag-bookings-logic.js';
import { toIsoDateOnly, effectiveBuchungAb } from './projektwochen-logic.js';

const BOOKINGS_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Bookings.Read.All',
    'https://graph.microsoft.com/Bookings.ReadWrite.All',
    'https://graph.microsoft.com/Bookings.Manage.All'
];

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

async function token() {
    return await G().getGraphToken(BOOKINGS_SCOPES);
}

function bizPath(id) {
    return '/solutions/bookingBusinesses/' + encodeURIComponent(id);
}

async function listAllPages(tok, path) {
    let next = path;
    const out = [];
    while (next) {
        const data = await G().graphJson('GET', next.indexOf('http') === 0 ? next : next, tok, undefined, 'v1.0');
        const rows = (data && data.value) || [];
        for (let i = 0; i < rows.length; i++) out.push(rows[i]);
        next = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
    }
    return out;
}

function sleep(ms) {
    return G().sleep(ms);
}

/**
 * @param {(msg: string) => void} [logFn]
 */
export async function listBookingBusinesses(logFn) {
    const write = typeof logFn === 'function' ? logFn : function () {};
    const tok = await token();
    write('Lade Bookings-Umgebungen …');
    const list = await listAllPages(tok, '/solutions/bookingBusinesses');
    return list
        .map((b) => ({
            id: String((b && b.id) || ''),
            displayName: String((b && b.displayName) || b.id || ''),
            publicUrl: String((b && b.publicUrl) || ''),
            isPublished: !!(b && b.isPublished)
        }))
        .filter((b) => b.id);
}

/**
 * Business anlegen oder vorhandenes per ID nutzen.
 * @param {object} aktion
 * @param {{ businessId?: string, displayName?: string }} [opts]
 * @param {(msg: string) => void} [logFn]
 */
export async function ensureBookingBusiness(aktion, opts, logFn) {
    const write = typeof logFn === 'function' ? logFn : function () {};
    const tok = await token();
    const existingId = String((opts && opts.businessId) || (aktion && aktion.bookingsBusinessId) || '').trim();

    if (existingId) {
        write('Prüfe Bookings-Business ' + existingId + ' …');
        const one = await G().graphJson('GET', bizPath(existingId), tok, undefined, 'v1.0');
        return {
            id: String((one && one.id) || existingId),
            displayName: String((one && one.displayName) || ''),
            publicUrl: String((one && one.publicUrl) || ''),
            isPublished: !!(one && one.isPublished),
            created: false
        };
    }

    const name = String((opts && opts.displayName) || defaultBusinessName(aktion)).trim();
    if (!name) throw new Error('Name für Bookings-Umgebung fehlt.');

    write('Suche vorhandene Umgebung „' + name + '" …');
    try {
        const found = await listAllPages(
            tok,
            '/solutions/bookingBusinesses?query=' + encodeURIComponent(name)
        );
        const hit = found.find(
            (b) => String((b && b.displayName) || '').toLowerCase() === name.toLowerCase()
        );
        if (hit && hit.id) {
            write('Vorhandene Umgebung wird genutzt: ' + hit.id);
            return {
                id: String(hit.id),
                displayName: String(hit.displayName || name),
                publicUrl: String(hit.publicUrl || ''),
                isPublished: !!hit.isPublished,
                created: false
            };
        }
    } catch (e) {
        write('Hinweis Suche: ' + ((e && e.message) || e));
    }

    const start = toIsoDateOnly(aktion && aktion.startdatum);
    const end = toIsoDateOnly(aktion && aktion.enddatum) || start;
    const body = {
        displayName: name,
        defaultCurrencyIso: 'EUR'
    };
    if (start) {
        const weekday = weekdayFromIsoDate(start);
        const startTime = toBookingsTime('08:00');
        const endTime = toBookingsTime('17:00');
        if (weekday && startTime && endTime) {
            const hours = buildBusinessHoursForDay({
                weekday,
                startTime,
                endTime
            });
            // Für die ganze Woche: alle Werktage 08–17
            if (hours && end) {
                const days = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'];
                body.businessHours = days.map((day) => ({
                    day,
                    timeSlots: [{ startTime, endTime }]
                }));
                ['sunday', 'saturday'].forEach((day) => {
                    body.businessHours.push({ day, timeSlots: [] });
                });
            }
        }
        const open = toIsoDateOnly(aktion && aktion.buchungAbDefault) || start;
        const policy = buildPwServiceSchedulingPolicy({
            eventDate: start,
            startHhmm: '08:00',
            endHhmm: '12:00',
            bookingOpenDate: open
        });
        if (policy) {
            body.schedulingPolicy = {
                timeSlotInterval: 'PT60M',
                minimumLeadTime: 'PT0M',
                maximumAdvance: policy.maximumAdvance,
                sendConfirmationsToOwner: true,
                allowStaffSelection: true
            };
        }
    }

    write('Lege Bookings-Business an: ' + name);
    const created = await G().graphJson('POST', '/solutions/bookingBusinesses', tok, body, 'v1.0');
    const id = String((created && created.id) || '');
    if (!id) throw new Error('Bookings-Business ohne ID.');
    await sleep(1500);
    try {
        write('Veröffentliche Buchungsseite …');
        await G().graphJson('POST', bizPath(id) + '/publish', tok, {}, 'v1.0');
    } catch (e) {
        write('Hinweis Publish: ' + ((e && e.message) || e));
    }
    let publicUrl = String((created && created.publicUrl) || '');
    try {
        const again = await G().graphJson('GET', bizPath(id), tok, undefined, 'v1.0');
        publicUrl = String((again && again.publicUrl) || publicUrl);
    } catch {
        /* ignore */
    }
    return {
        id,
        displayName: name,
        publicUrl,
        isPublished: true,
        created: true
    };
}

async function ensureStaffMember(tok, businessId, email, displayName, logFn) {
    const write = typeof logFn === 'function' ? logFn : function () {};
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return '';

    const staff = await listAllPages(tok, bizPath(businessId) + '/staffMembers');
    const hit = staff.find(
        (s) =>
            String((s && s.emailAddress) || '')
                .trim()
                .toLowerCase() === em
    );
    if (hit && hit.id) {
        write('Staff vorhanden: ' + em);
        return String(hit.id);
    }

    write('Lege Staff an: ' + em);
    const created = await G().graphJson(
        'POST',
        bizPath(businessId) + '/staffMembers',
        tok,
        {
            '@odata.type': '#microsoft.graph.bookingStaffMember',
            displayName: String(displayName || em).slice(0, 100),
            emailAddress: em,
            role: 'externalGuest',
            useBusinessHours: true
        },
        'v1.0'
    );
    await sleep(200);
    return String((created && created.id) || '');
}

/**
 * Ein freigegebenes Angebot als Bookings-Service syncen.
 * @param {string} businessId
 * @param {object} angebot
 * @param {object|null} aktion
 * @param {(msg: string) => void} [logFn]
 */
export async function syncAngebotToBookings(businessId, angebot, aktion, logFn) {
    const write = typeof logFn === 'function' ? logFn : function () {};
    if (!businessId) throw new Error('Bookings-Business-ID fehlt.');
    if (!angebot || angebot.status !== 'freigegeben') {
        throw new Error('Nur freigegebene Angebote können gesynct werden.');
    }

    const built = buildServicePayloadFromAngebot(angebot, aktion);
    if (!built.ok) throw new Error(built.error || 'Payload ungültig.');

    const tok = await token();
    const staffId = await ensureStaffMember(
        tok,
        businessId,
        angebot.lehrerEmail,
        angebot.lehrerCode || angebot.lehrerEmail,
        write
    );
    const payload = { ...built.payload };
    if (staffId) payload.staffMemberIds = [staffId];
    else payload.staffMemberIds = [];

    write(
        'Dienst „' +
            payload.displayName +
            '": Kapazität ' +
            payload.maximumAttendeesCount +
            ', maximumAdvance ' +
            built.maxAdvance +
            ' (Buchung ab ' +
            built.openDate +
            ').'
    );

    let serviceId = String(angebot.bookingsServiceId || '').trim();
    let webUrl = String(angebot.bookingsBookingUrl || '').trim();
    let created = false;

    if (serviceId) {
        write('Aktualisiere Service ' + serviceId + ' …');
        try {
            const updated = await G().graphJson(
                'PATCH',
                bizPath(businessId) + '/services/' + encodeURIComponent(serviceId),
                tok,
                payload,
                'v1.0'
            );
            webUrl = String((updated && updated.webUrl) || webUrl);
        } catch (e) {
            write('Update fehlgeschlagen, lege neu an: ' + ((e && e.message) || e));
            serviceId = '';
        }
    }

    if (!serviceId) {
        // Namenssuche
        try {
            const services = await listAllPages(tok, bizPath(businessId) + '/services');
            const marker = angebot.angebotId ? '[PW:' + angebot.angebotId + ']' : '';
            const hit = services.find((s) => {
                const desc = String((s && s.description) || '');
                if (marker && desc.indexOf(marker) !== -1) return true;
                return (
                    String((s && s.displayName) || '').toLowerCase() ===
                    payload.displayName.toLowerCase()
                );
            });
            if (hit && hit.id) {
                serviceId = String(hit.id);
                write('Bestehenden Dienst gefunden: ' + serviceId);
                const updated = await G().graphJson(
                    'PATCH',
                    bizPath(businessId) + '/services/' + encodeURIComponent(serviceId),
                    tok,
                    payload,
                    'v1.0'
                );
                webUrl = String((updated && updated.webUrl) || hit.webUrl || webUrl);
            }
        } catch (e) {
            write('Hinweis Services lesen: ' + ((e && e.message) || e));
        }
    }

    if (!serviceId) {
        write('Lege neuen Dienst an …');
        const createdSvc = await G().graphJson(
            'POST',
            bizPath(businessId) + '/services',
            tok,
            payload,
            'v1.0'
        );
        serviceId = String((createdSvc && createdSvc.id) || '');
        webUrl = String((createdSvc && createdSvc.webUrl) || '');
        created = true;
        if (!serviceId) throw new Error('Service ohne ID.');
    }

    if (!webUrl) {
        try {
            const biz = await G().graphJson('GET', bizPath(businessId), tok, undefined, 'v1.0');
            webUrl = String((biz && biz.publicUrl) || '');
        } catch {
            /* ignore */
        }
    }

    return {
        serviceId,
        bookingUrl: webUrl,
        created,
        openDate: built.openDate,
        maxAdvance: built.maxAdvance
    };
}

/**
 * Termine/Teilnehmer im Aktionsfenster laden.
 * @param {string} businessId
 * @param {object|null} aktion
 * @param {object[]} angebote
 * @param {object[]} [students]
 * @param {(msg: string) => void} [logFn]
 */
export async function loadBookingsAttendees(businessId, aktion, angebote, students, logFn) {
    const write = typeof logFn === 'function' ? logFn : function () {};
    if (!businessId) throw new Error('Bookings-Business-ID fehlt.');
    const tok = await token();

    const start = toIsoDateOnly(aktion && aktion.startdatum) || '1970-01-01';
    const end = toIsoDateOnly(aktion && aktion.enddatum) || start;
    // calendarView erwartet DateTimeOffset
    const startParam = encodeURIComponent(start + 'T00:00:00.0000000');
    const endParam = encodeURIComponent(end + 'T23:59:59.0000000');

    write('Lade calendarView ' + start + ' – ' + end + ' …');
    let appointments = [];
    try {
        appointments = await listAllPages(
            tok,
            bizPath(businessId) + '/calendarView?start=' + startParam + '&end=' + endParam
        );
    } catch (e) {
        write('calendarView fehlgeschlagen, versuche appointments: ' + ((e && e.message) || e));
        appointments = await listAllPages(tok, bizPath(businessId) + '/appointments');
    }

    // Details nachladen wenn customers leer
    const enriched = [];
    for (let i = 0; i < appointments.length; i++) {
        const ap = appointments[i];
        const hasCust =
            (Array.isArray(ap.customers) && ap.customers.length) || ap.customerEmailAddress;
        if (hasCust || !ap.id) {
            enriched.push(ap);
            continue;
        }
        try {
            const detail = await G().graphJson(
                'GET',
                bizPath(businessId) + '/appointments/' + encodeURIComponent(ap.id),
                tok,
                undefined,
                'v1.0'
            );
            enriched.push(detail || ap);
            await sleep(80);
        } catch {
            enriched.push(ap);
        }
    }

    const norm = normalizeAppointments(enriched, angebote);
    const rows = enrichAttendeesWithKlasse(norm.rows, students);
    write('Termine: ' + enriched.length + ', Teilnehmerzeilen: ' + rows.length);
    return {
        appointments: enriched,
        rows,
        occupancy: norm.occupancy
    };
}

export { BOOKINGS_SCOPES, effectiveBuchungAb };
