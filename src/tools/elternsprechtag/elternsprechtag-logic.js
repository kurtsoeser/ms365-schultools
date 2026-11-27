/**
 * Elternsprechtag für Lehrkräfte: eigene Kalendertermine des Sprechtags.
 * Kein Bookings-Admin-Zugriff.
 */

export const ELTERNSPRECHTAG_SETUP_KEY = 'ms365-elternsprechtag-bookings-v1';

/**
 * @param {string} value
 */
export function normEmail(value) {
    return String(value || '')
        .trim()
        .toLowerCase();
}

/**
 * @param {unknown} raw
 */
export function readSetup(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const minutes = Number(o.minutes);
    return {
        eventDate: String(o.eventDate || '').trim(),
        start: String(o.start || '').trim(),
        end: String(o.end || '').trim(),
        minutes: Number.isFinite(minutes) ? minutes : 0,
        serviceName: String(o.serviceName || '').trim(),
        bizName: String(o.bizName || '').trim(),
        businessId: String(o.businessId || '').trim(),
        publicUrl: String(o.publicUrl || '').trim(),
        siteUrl: String(o.siteUrl || '').trim()
    };
}

/**
 * @param {Storage|null|undefined} storage
 */
export function readSetupFromStorage(storage) {
    try {
        const raw = storage && storage.getItem(ELTERNSPRECHTAG_SETUP_KEY);
        if (!raw) return readSetup(null);
        return readSetup(JSON.parse(raw));
    } catch {
        return readSetup(null);
    }
}

/**
 * @param {string} html
 */
export function plainText(html) {
    return String(html || '')
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/p>/gi, '\n')
        .replace(/<[^>]+>/g, ' ')
        .replace(/&nbsp;/gi, ' ')
        .replace(/&amp;/gi, '&')
        .replace(/\r/g, '')
        .replace(/[ \t]+\n/g, '\n')
        .replace(/\n{3,}/g, '\n\n')
        .trim();
}

/**
 * @param {string} dateTime
 */
export function clockLabel(dateTime) {
    const s = String(dateTime || '').trim();
    const local = s.match(/T(\d{2}):(\d{2})/);
    const hasZone = /Z$/i.test(s) || /[+-]\d{2}:\d{2}$/.test(s);
    if (local && !hasZone) return local[1] + ':' + local[2];
    const d = new Date(s);
    if (Number.isNaN(d.getTime())) return '';
    try {
        return new Intl.DateTimeFormat('de-AT', {
            hour: '2-digit',
            minute: '2-digit',
            timeZone: 'Europe/Vienna'
        }).format(d);
    } catch {
        return local ? local[1] + ':' + local[2] : '';
    }
}

/**
 * @param {string} isoDate YYYY-MM-DD
 */
export function dayHeading(isoDate) {
    const m = String(isoDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return '';
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
    if (Number.isNaN(d.getTime())) return '';
    try {
        return new Intl.DateTimeFormat('de-AT', {
            weekday: 'long',
            day: 'numeric',
            month: 'long',
            year: 'numeric'
        }).format(d);
    } catch {
        return isoDate;
    }
}

/**
 * @param {string} isoDate
 * @returns {{ start: Date, end: Date }|null}
 */
export function dayBounds(isoDate) {
    const m = String(isoDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    const y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);
    const start = new Date(y, mo - 1, d, 0, 0, 0, 0);
    const end = new Date(y, mo - 1, d + 1, 0, 0, 0, 0);
    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return null;
    return { start, end };
}

/**
 * @param {object|null|undefined} ev
 */
function organizerEmail(ev) {
    const org = ev && ev.organizer && ev.organizer.emailAddress;
    return normEmail(org && org.address);
}

/**
 * @param {object|null|undefined} ev
 */
function eventBlob(ev) {
    const e = ev || {};
    const loc =
        e.location && typeof e.location === 'object'
            ? e.location.displayName
            : e.location;
    const body =
        e.bodyPreview != null
            ? e.bodyPreview
            : e.body && e.body.content
              ? e.body.content
              : '';
    return [e.subject, loc, body, organizerEmail(e)].join('\n');
}

/**
 * Termin kommt von der Bookings-Seite, nicht aus einem beliebigen Kalendereintrag.
 * @param {object|null|undefined} ev
 * @param {ReturnType<typeof readSetup>} setup
 */
export function eventLooksLikeBooking(ev, setup) {
    const e = ev || {};
    if (e.isCancelled || e.isAllDay) return false;
    const cfg = setup || readSetup(null);
    const blob = eventBlob(e);
    const org = organizerEmail(e);
    const biz = normEmail(cfg.businessId);
    if (biz && (org === biz || blob.toLowerCase().indexOf(biz) >= 0)) return true;
    if (/bookings\.cloud\.microsoft|outlook\.office(?:365)?\.com\/book\/|microsoft bookings/i.test(blob)) {
        return true;
    }
    if (org && /bookings/i.test(org)) return true;
    const service = String(cfg.serviceName || '').trim();
    const subject = String(e.subject || '');
    if (service && service.length >= 4 && subject.toLowerCase().indexOf(service.toLowerCase()) >= 0) {
        return true;
    }
    const bizName = String(cfg.bizName || '').trim();
    if (bizName && bizName.length >= 4 && subject.toLowerCase().indexOf(bizName.toLowerCase()) >= 0) {
        return true;
    }
    return false;
}

/**
 * @param {object|null|undefined} ev
 * @param {string} selfEmail
 */
export function parentFromEvent(ev, selfEmail) {
    const e = ev || {};
    const self = normEmail(selfEmail);
    const org = organizerEmail(e);
    const attendees = Array.isArray(e.attendees) ? e.attendees : [];
    const other = attendees
        .map(function (a) {
            const mail = a && a.emailAddress ? a.emailAddress : {};
            return {
                name: String(mail.name || '').trim(),
                email: normEmail(mail.address)
            };
        })
        .filter(function (a) {
            return a.email && a.email !== self && a.email !== org && !/bookings/i.test(a.email);
        });
    const body = plainText(
        e.bodyPreview != null ? e.bodyPreview : e.body && e.body.content ? e.body.content : ''
    );
    const mailMatch = body.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
    const bodyMail = mailMatch ? normEmail(mailMatch[0]) : '';
    const nameMatch = body.match(/(?:Name|Kunde|Kundin|Customer)\s*[:：]\s*([^\n,]{2,80})/i);
    const phoneMatch = body.match(/(?:Telefon|Phone|Tel\.?)\s*[:：]\s*([+\d][\d\s/()-]{5,})/i);
    const kindMatch = body.match(/\bKind\s*[:：]\s*([^\n,]{2,80})/i);
    const klasseMatch = body.match(/\bKlasse\s*[:：]\s*([^\n,]{1,40})/i);
    const email =
        (other[0] && other[0].email) ||
        (bodyMail && bodyMail !== self && bodyMail !== org && !/bookings/i.test(bodyMail) ? bodyMail : '');
    const attendeeName = other[0] && other[0].name && !/bookings/i.test(other[0].name) ? other[0].name : '';
    const name = attendeeName || (nameMatch ? nameMatch[1].trim() : '') || email;
    const extras = [];
    if (kindMatch) extras.push('Kind: ' + kindMatch[1].trim());
    if (klasseMatch) extras.push('Klasse: ' + klasseMatch[1].trim());
    return {
        name: name,
        email: email,
        phone: phoneMatch ? phoneMatch[1].trim() : '',
        childName: kindMatch ? kindMatch[1].trim() : '',
        className: klasseMatch ? klasseMatch[1].trim() : '',
        detail: extras.join(' · ')
    };
}

/**
 * @param {object} ev
 * @param {ReturnType<typeof readSetup>} setup
 * @param {string} selfEmail
 */
export function toSlot(ev, setup, selfEmail) {
    const e = ev || {};
    const startRaw = e.start && e.start.dateTime ? e.start.dateTime : '';
    const endRaw = e.end && e.end.dateTime ? e.end.dateTime : '';
    const parent = parentFromEvent(e, selfEmail);
    const loc =
        e.location && typeof e.location === 'object'
            ? String(e.location.displayName || '').trim()
            : String(e.location || '').trim();
    const join =
        e.onlineMeeting && e.onlineMeeting.joinUrl
            ? String(e.onlineMeeting.joinUrl).trim()
            : '';
    return {
        id: String(e.id || ''),
        startLabel: clockLabel(startRaw),
        endLabel: clockLabel(endRaw),
        sortKey: startRaw || '',
        subject: String(e.subject || '').trim(),
        parentName: parent.name,
        parentEmail: parent.email,
        phone: parent.phone,
        childName: parent.childName,
        className: parent.className,
        detail: parent.detail,
        place: loc,
        webLink: String(e.webLink || '').trim(),
        joinUrl: join,
        online: !!e.isOnlineMeeting || !!join,
        booking: eventLooksLikeBooking(e, setup)
    };
}

/**
 * @param {object[]} events
 * @param {ReturnType<typeof readSetup>} setup
 * @param {string} selfEmail
 * @param {{ includeOther?: boolean }} [opts]
 */
export function slotsFromEvents(events, setup, selfEmail, opts) {
    const onlyBookings = !(opts && opts.includeOther);
    return (Array.isArray(events) ? events : [])
        .filter(function (ev) {
            if (!ev || ev.isCancelled) return false;
            if (onlyBookings) return eventLooksLikeBooking(ev, setup);
            return !ev.isAllDay;
        })
        .map(function (ev) {
            return toSlot(ev, setup, selfEmail);
        })
        .sort(function (a, b) {
            return String(a.sortKey).localeCompare(String(b.sortKey));
        });
}
