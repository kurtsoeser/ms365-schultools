/**
 * Kalender und Termine ↔ Outlook (Graph calendarView / events).
 * Kein Parallel-SoR: Outlook-Kalender (Schul-/Lehrer-/Mein) sind die Quelle.
 */

import { fetchAllPages, getGraphToken, graphJson } from '../../shared/graph-client.js';
import { SICHTBARKEIT, SYNC_STATUS, TERMIN_TYP, toIsoDateOnly, toIsoDateTime } from './termine-hub-model.js';
import { SOURCE_IDS } from './termine-hub-settings.js';

export const TH_CALENDAR_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Calendars.ReadWrite',
    'https://graph.microsoft.com/Calendars.ReadWrite.Shared',
    // Mitgliedschaften auflisten (ARGE-/Lehrer-Gruppenkalender)
    'https://graph.microsoft.com/Group.Read.All'
];

const TZ = 'Europe/Vienna';

function graphCalendarHeaders() {
    return {
        Prefer: 'outlook.timezone="' + TZ + '", outlook.send-event-invitations="false"'
    };
}

/**
 * @param {unknown} err
 * @param {string} [calendarUser]
 */
export function formatThCalendarAccessError(err, calendarUser) {
    const raw = err && err.message ? String(err.message) : String(err || '');
    if (!/ErrorAccessDenied|Access is denied|403/i.test(raw)) return raw;
    const target = String(calendarUser || '').trim().toLowerCase();
    return (
        'Kein Zugriff auf den Kalender' +
        (target ? ' von ' + target : '') +
        '. Für fremde/Shared-Mailboxen braucht das angemeldete Konto Bearbeiten-Recht in Outlook ' +
        '(Calendars.ReadWrite.Shared). Global Admin ersetzt das nicht.'
    );
}

/**
 * @param {{ kind: 'me'|'user'|'group', upn?: string, groupId?: string, calendarId?: string, signedInEmail?: string }} target
 */
export function calendarBasePath(target) {
    const t = target || {};
    if (t.kind === 'group') {
        const gid = String(t.groupId || '').trim();
        if (!gid) throw new Error('Gruppen-ID fehlt.');
        return '/groups/' + encodeURIComponent(gid) + '/calendar';
    }
    const cal = String(t.calendarId || '').trim();
    const signedIn = String(t.signedInEmail || '')
        .trim()
        .toLowerCase();
    const user = String(t.upn || '')
        .trim()
        .toLowerCase();
    const asMe = t.kind === 'me' || (!!signedIn && !!user && signedIn === user);
    if (asMe) {
        if (cal) return '/me/calendars/' + encodeURIComponent(cal);
        return '/me/calendar';
    }
    if (!user) throw new Error('Kalender-UPN fehlt.');
    const base = '/users/' + encodeURIComponent(user);
    if (cal) return base + '/calendars/' + encodeURIComponent(cal);
    return base + '/calendar';
}

/**
 * @param {Date|string} value
 */
export function toGraphDateTimeParam(value) {
    const d = value instanceof Date ? value : new Date(value);
    if (Number.isNaN(d.getTime())) throw new Error('Ungültiges Datum für calendarView.');
    return d.toISOString();
}

/**
 * Graph-Event → Hub-Termin.
 * @param {object} ev
 * @param {{ id: string, label: string }} source
 */
export function mapGraphEventToTermin(ev, source) {
    const e = ev && typeof ev === 'object' ? ev : {};
    const isAllDay = !!e.isAllDay;
    const startRaw = e.start && e.start.dateTime ? e.start.dateTime : e.start;
    const endRaw = e.end && e.end.dateTime ? e.end.dateTime : e.end;
    let start = toIsoDateTime(startRaw);
    let end = toIsoDateTime(endRaw);
    if (isAllDay) {
        const sd = toIsoDateOnly(startRaw);
        const ed = toIsoDateOnly(endRaw);
        // Graph all-day end ist exklusiv → Anzeige auf Vortag ziehen
        let displayEnd = ed;
        if (sd && ed && ed > sd) {
            const dt = new Date(ed + 'T12:00:00');
            dt.setDate(dt.getDate() - 1);
            displayEnd = toIsoDateOnly(dt);
        }
        start = sd ? new Date(sd + 'T12:00:00').toISOString() : start;
        end = displayEnd ? new Date(displayEnd + 'T12:00:00').toISOString() : end;
    }
    const loc =
        e.location && typeof e.location === 'object'
            ? String(e.location.displayName || '').trim()
            : String(e.location || '').trim();
    const body =
        e.bodyPreview != null
            ? String(e.bodyPreview)
            : e.body && e.body.content
              ? String(e.body.content)
              : '';
    const sourceId = source && source.id ? String(source.id) : SOURCE_IDS.ME;
    return {
        id: String(e.id || '').trim(),
        graphEventId: String(e.id || '').trim(),
        title: String(e.subject || '').trim() || '(Ohne Titel)',
        start,
        end: end || start,
        allDay: isAllDay,
        ort: loc,
        beschreibung: body.trim(),
        sichtbarkeit: sourceId === SOURCE_IDS.SCHUL ? SICHTBARKEIT.OEFFENTLICH : SICHTBARKEIT.INTERN,
        terminTyp: TERMIN_TYP.SONSTIGES,
        zielgruppe: '',
        syncStatus: SYNC_STATUS.SYNCED,
        sourceId,
        sourceLabel: source && source.label ? String(source.label) : '',
        webLink: String(e.webLink || '').trim(),
        createdAt: String(e.createdDateTime || '').trim(),
        updatedAt: String(e.lastModifiedDateTime || '').trim()
    };
}

/**
 * Hub-Draft → Graph event body.
 * @param {object} draft
 */
export function buildGraphEventBody(draft) {
    const d = draft && typeof draft === 'object' ? draft : {};
    const title = String(d.title || '').trim();
    if (!title) throw new Error('Titel fehlt.');
    const allDay = !!d.allDay;
    const startIso = toIsoDateTime(d.start);
    const endIso = toIsoDateTime(d.end || d.start);
    if (!startIso) throw new Error('Beginn fehlt.');
    const ort = String(d.ort || '').trim();
    const beschreibung = String(d.beschreibung || '').trim();
    const categories = [];
    if (d.terminTyp) categories.push('MS365-' + String(d.terminTyp));
    if (d.sichtbarkeit) categories.push('Sichtbarkeit-' + String(d.sichtbarkeit));

    if (allDay) {
        const start = toIsoDateOnly(d.start || startIso);
        let end = toIsoDateOnly(d.end || endIso || start);
        if (!start) throw new Error('Beginn (Datum) fehlt.');
        if (!end) end = start;
        // Graph: Enddatum exklusiv
        const endDt = new Date(end + 'T12:00:00');
        endDt.setDate(endDt.getDate() + 1);
        const endExclusive = toIsoDateOnly(endDt);
        return {
            subject: title,
            body: { contentType: 'text', content: beschreibung },
            isAllDay: true,
            start: { dateTime: start, timeZone: TZ },
            end: { dateTime: endExclusive, timeZone: TZ },
            location: ort ? { displayName: ort } : undefined,
            categories,
            isReminderOn: false,
            responseRequested: false
        };
    }

    const startLocal = startIso.replace(/\.\d{3}Z$/, '').replace(/Z$/, '').slice(0, 19);
    const endLocal = (endIso || startIso).replace(/\.\d{3}Z$/, '').replace(/Z$/, '').slice(0, 19);
    return {
        subject: title,
        body: { contentType: 'text', content: beschreibung },
        isAllDay: false,
        start: { dateTime: startLocal, timeZone: TZ },
        end: { dateTime: endLocal, timeZone: TZ },
        location: ort ? { displayName: ort } : undefined,
        categories,
        isReminderOn: true,
        reminderMinutesBeforeStart: 15,
        responseRequested: false
    };
}

/**
 * @param {{ kind: 'me'|'user', upn?: string, calendarId?: string, signedInEmail?: string }} target
 * @param {Date|string} start
 * @param {Date|string} end
 */
export async function fetchCalendarView(target, start, end) {
    const base = calendarBasePath(target);
    const startParam = encodeURIComponent(toGraphDateTimeParam(start));
    const endParam = encodeURIComponent(toGraphDateTimeParam(end));
    const select =
        'id,subject,start,end,isAllDay,location,bodyPreview,webLink,createdDateTime,lastModifiedDateTime,categories';
    const path =
        base +
        '/calendarView?startDateTime=' +
        startParam +
        '&endDateTime=' +
        endParam +
        '&$select=' +
        select +
        '&$orderby=start/dateTime&$top=100';
    const token = await getGraphToken(TH_CALENDAR_SCOPES);
    const { items, truncated } = await fetchAllPages(token, path, {
        maxItems: 800,
        maxPages: 20,
        extraHeaders: graphCalendarHeaders()
    });
    return { items: items || [], truncated: !!truncated };
}

/**
 * @param {Array<{ id: string, label: string, kind: 'me'|'user', upn: string }>} sources
 * @param {Date|string} start
 * @param {Date|string} end
 * @param {{ signedInEmail?: string }} [opts]
 */
export async function fetchEventsForSources(sources, start, end, opts) {
    const signedInEmail = String((opts && opts.signedInEmail) || '')
        .trim()
        .toLowerCase();
    const list = Array.isArray(sources) ? sources : [];
    /** @type {object[]} */
    const events = [];
    /** @type {string[]} */
    const errors = [];
    let truncated = false;

    for (const src of list) {
        try {
            const target = {
                kind: src.kind,
                upn: src.upn,
                groupId: src.groupId,
                signedInEmail
            };
            const res = await fetchCalendarView(target, start, end);
            if (res.truncated) truncated = true;
            for (const ev of res.items) {
                events.push(mapGraphEventToTermin(ev, { id: src.id, label: src.label }));
            }
        } catch (err) {
            errors.push(
                (src.label || src.id) +
                    ': ' +
                    formatThCalendarAccessError(err, src.upn || src.groupId)
            );
        }
    }

    events.sort((a, b) => String(a.start || '').localeCompare(String(b.start || '')));
    return { events, errors, truncated };
}

/**
 * Unified M365-Gruppen der angemeldeten Person (für Gruppenkalender).
 * @returns {Promise<Array<{ id: string, displayName: string, mail: string }>>}
 */
export async function listMyUnifiedGroups() {
    const token = await getGraphToken(TH_CALENDAR_SCOPES);
    const path =
        '/me/memberOf/microsoft.graph.group?$select=id,displayName,mail,groupTypes&$top=100';
    const { items } = await fetchAllPages(token, path, { maxItems: 400, maxPages: 8 });
    return (items || [])
        .filter((g) => {
            const types = Array.isArray(g && g.groupTypes) ? g.groupTypes : [];
            return types.map(String).includes('Unified');
        })
        .map((g) => ({
            id: String(g.id || '').trim(),
            displayName: String(g.displayName || '').trim() || String(g.mail || '').trim() || 'Gruppe',
            mail: String(g.mail || '')
                .trim()
                .toLowerCase()
        }))
        .filter((g) => g.id)
        .sort((a, b) => a.displayName.localeCompare(b.displayName, 'de'));
}

/**
 * @param {object} draft
 * @param {{ kind: 'me'|'user', upn?: string, signedInEmail?: string }} target
 */
/**
 * @param {object} draft
 * @param {{ kind: 'me'|'user', upn?: string, signedInEmail?: string, sourceId?: string, sourceLabel?: string }} target
 */
export async function createOutlookEvent(draft, target) {
    const body = buildGraphEventBody(draft);
    const path = calendarBasePath(target) + '/events';
    const token = await getGraphToken(TH_CALENDAR_SCOPES);
    const sourceId =
        (target && target.sourceId) ||
        (target && target.kind === 'me' ? SOURCE_IDS.ME : SOURCE_IDS.LEHRER);
    const sourceLabel =
        (target && target.sourceLabel) ||
        (target && target.kind === 'me' ? 'Mein Kalender' : String((target && target.upn) || ''));
    try {
        const created = await graphJson('POST', path, token, body, graphCalendarHeaders());
        return mapGraphEventToTermin(created, { id: sourceId, label: sourceLabel });
    } catch (err) {
        throw new Error(formatThCalendarAccessError(err, target && target.upn));
    }
}

/**
 * @param {string} eventId
 * @param {{ kind: 'me'|'user', upn?: string, signedInEmail?: string }} target
 */
export async function deleteOutlookEvent(eventId, target) {
    const id = String(eventId || '').trim();
    if (!id) throw new Error('Event-ID fehlt.');
    const path = calendarBasePath(target) + '/events/' + encodeURIComponent(id);
    const token = await getGraphToken(TH_CALENDAR_SCOPES);
    try {
        await graphJson('DELETE', path, token, undefined, graphCalendarHeaders());
        return { ok: true };
    } catch (err) {
        throw new Error(formatThCalendarAccessError(err, target && target.upn));
    }
}

/**
 * @returns {string}
 */
export function signedInEmail() {
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const a = window.ms365AuthGetAccountInfo();
            return String((a && (a.username || a.upn || a.mail || a.email)) || '')
                .trim()
                .toLowerCase();
        }
        if (typeof window.ms365AuthGetUserPrincipalName === 'function') {
            return String(window.ms365AuthGetUserPrincipalName() || '')
                .trim()
                .toLowerCase();
        }
    } catch {
        /* ignore */
    }
    return '';
}

/**
 * IT-Diagnose: Kalenderquelle kurz anlesen.
 * @param {{ kind: 'me'|'user', upn?: string, signedInEmail?: string, sourceId?: string, sourceLabel?: string }} target
 * @returns {Promise<{ ok: boolean, status: 'ok'|'readonly'|'denied'|'unreachable'|'missing', message: string }>}
 */
export async function probeCalendarSource(target) {
    const t = target || {};
    if (t.kind === 'user' && !String(t.upn || '').trim()) {
        return {
            ok: false,
            status: 'missing',
            message: 'UPN fehlt – Kalenderquelle ist noch nicht eingerichtet.'
        };
    }
    const from = new Date();
    from.setHours(0, 0, 0, 0);
    const to = new Date(from);
    to.setDate(to.getDate() + 1);
    try {
        await fetchCalendarView(t, from, to);
        return {
            ok: true,
            status: 'ok',
            message: 'Erreichbar – Lesen funktioniert.'
        };
    } catch (err) {
        const raw = err && err.message ? String(err.message) : String(err || '');
        if (/ErrorAccessDenied|Access is denied|403|Keine? Zugriff/i.test(raw)) {
            return { ok: false, status: 'denied', message: raw };
        }
        return {
            ok: false,
            status: 'unreachable',
            message: raw || 'Kalender nicht erreichbar.'
        };
    }
}
