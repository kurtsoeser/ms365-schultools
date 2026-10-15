/**
 * Genehmigte Lehrer-Freistellungen → freigegebener Outlook-Kalender (Graph).
 */
import { getGraphToken, graphJson } from '../../shared/graph-client.js';
import { toIsoDateOnly, normalizeStatus } from './lfr-logic.js';
import { loadSetupCfg } from './lfr-state.js';

export const LFR_CALENDAR_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Calendars.ReadWrite',
    /** Für fremde/freigegebene Postfächer (nicht Global-Admin-Ersatz). */
    'https://graph.microsoft.com/Calendars.ReadWrite.Shared'
];

const EVENT_MAP_LS = 'ms365-lfr-outlook-event-v1';
const TZ = 'Europe/Vienna';

/**
 * Delegierter Graph-Zugriff auf fremde Kalender braucht Freigabe + Shared-Scope.
 * Global Admin allein reicht nicht.
 * @param {unknown} err
 * @param {string} [calendarUser]
 */
export function formatLfrCalendarAccessError(err, calendarUser) {
    const raw = err && err.message ? String(err.message) : String(err || '');
    if (!/ErrorAccessDenied|Access is denied|403/i.test(raw)) return raw;
    const target = String(calendarUser || '').trim().toLowerCase();
    return (
        'Kein Schreibzugriff auf den Kalender' +
        (target ? ' von ' + target : '') +
        '. Mit dem angemeldeten Konto geht nur der eigene Kalender – oder ein Kalender, der dir mit Bearbeiten-Recht freigegeben wurde. ' +
        'Global Admin ersetzt das nicht. Lösung: im IT-Setup deinen UPN als Kalender-Besitzer setzen, oder „In meinen Kalender“ nutzen, ' +
        'oder den Zielkalender in Outlook für dich freigeben.'
    );
}

function graphCalendarWriteHeaders() {
    return {
        Prefer: 'outlook.timezone="' + TZ + '", outlook.send-event-invitations="false"'
    };
}

function eventMapKey(antragId) {
    return String(antragId || '').trim();
}

export function getOutlookEventId(antragId) {
    try {
        const raw = JSON.parse(localStorage.getItem(EVENT_MAP_LS) || '{}');
        return String(raw[eventMapKey(antragId)] || '').trim();
    } catch {
        return '';
    }
}

function setOutlookEventId(antragId, eventId) {
    try {
        const raw = JSON.parse(localStorage.getItem(EVENT_MAP_LS) || '{}');
        const k = eventMapKey(antragId);
        if (!k) return;
        if (eventId) raw[k] = String(eventId);
        else delete raw[k];
        localStorage.setItem(EVENT_MAP_LS, JSON.stringify(raw));
    } catch {
        /* ignore */
    }
}

/**
 * @returns {{ calendarUser: string, calendarId: string }}
 */
export function loadCalendarTargetFromSetup() {
    const c = loadSetupCfg();
    return {
        calendarUser: String(c.outlookCalendarUser || '').trim().toLowerCase(),
        calendarId: String(c.outlookCalendarId || '').trim()
    };
}

/**
 * @param {object} it Planner-Item
 */
export function buildLfrCalendarEvent(it) {
    const beginn = String(it.beginn || '').trim();
    const ende = String(it.ende || '').trim();
    if (!beginn || !ende) throw new Error('Beginn/Ende fehlen.');
    const subject =
        'Freistellung: ' +
        String(it.lehrerName || it.titel || 'Lehrkraft').trim() +
        (it.kategorie ? ' (' + it.kategorie + ')' : '');
    const body = [
        'Lehrer-Freistellung (genehmigt)',
        it.antragId ? 'Antrag-ID: ' + it.antragId : '',
        it.lehrerEmail ? 'E-Mail: ' + it.lehrerEmail : '',
        it.beschreibung ? String(it.beschreibung).trim() : ''
    ]
        .filter(Boolean)
        .join(' · ');

    const allDay = beginn.length <= 10 && ende.length <= 10;
    if (allDay) {
        const start = toIsoDateOnly(beginn) || beginn.slice(0, 10);
        let end = toIsoDateOnly(ende) || ende.slice(0, 10);
        const endDt = new Date(end + 'T12:00:00Z');
        endDt.setUTCDate(endDt.getUTCDate() + 1);
        end =
            endDt.getUTCFullYear() +
            '-' +
            String(endDt.getUTCMonth() + 1).padStart(2, '0') +
            '-' +
            String(endDt.getUTCDate()).padStart(2, '0');
        return {
            subject,
            body: { contentType: 'text', content: body },
            isAllDay: true,
            start: { dateTime: start, timeZone: TZ },
            end: { dateTime: end, timeZone: TZ },
            showAs: 'free',
            categories: ['Lehrer-Freistellung'],
            isReminderOn: false,
            responseRequested: false
        };
    }

    const startDt = beginn.replace('Z', '').slice(0, 19);
    const endDt = ende.replace('Z', '').slice(0, 19);
    return {
        subject,
        body: { contentType: 'text', content: body },
        isAllDay: false,
        start: { dateTime: startDt, timeZone: TZ },
        end: { dateTime: endDt, timeZone: TZ },
        showAs: 'free',
        categories: ['Lehrer-Freistellung'],
        isReminderOn: false,
        responseRequested: false
    };
}

/**
 * @param {{ calendarUser?: string, calendarId?: string, signedInEmail?: string, useMe?: boolean }} target
 */
export function calendarEventsPath(target) {
    const t = target || {};
    const cal = String(t.calendarId || '').trim();
    const signedIn = String(t.signedInEmail || '')
        .trim()
        .toLowerCase();
    const user = String(t.calendarUser || '')
        .trim()
        .toLowerCase();
    const asMe = t.useMe === true || (!!signedIn && !!user && signedIn === user);
    if (asMe) {
        if (cal) return '/me/calendars/' + encodeURIComponent(cal) + '/events';
        return '/me/calendar/events';
    }
    if (!user) throw new Error('Kalender-Besitzer (UPN) fehlt – im IT-Setup eintragen.');
    const base = '/users/' + encodeURIComponent(user);
    if (cal) return base + '/calendars/' + encodeURIComponent(cal) + '/events';
    return base + '/calendar/events';
}

/**
 * @param {object} it
 * @param {{ calendarUser?: string, calendarId?: string, signedInEmail?: string, useMe?: boolean }} [opts]
 */
export async function upsertLfrOutlookEvent(it, opts) {
    const setup = loadCalendarTargetFromSetup();
    const target = {
        calendarUser: (opts && opts.calendarUser) || setup.calendarUser,
        calendarId: (opts && opts.calendarId) || setup.calendarId,
        signedInEmail: (opts && opts.signedInEmail) || '',
        useMe: !!(opts && opts.useMe)
    };
    const pathBase = calendarEventsPath(target);
    const body = buildLfrCalendarEvent(it);
    const tok = await getGraphToken(LFR_CALENDAR_SCOPES);
    const antragId = it.antragId || it.itemId;
    let existingId = getOutlookEventId(antragId);

    try {
        if (existingId) {
            try {
                await graphJson(
                    'PATCH',
                    pathBase + '/' + encodeURIComponent(existingId),
                    tok,
                    body,
                    graphCalendarWriteHeaders()
                );
                return { eventId: existingId, created: false };
            } catch (err) {
                const msg = err && err.message ? String(err.message) : String(err);
                if (!/404|ErrorItemNotFound|not found/i.test(msg)) throw err;
                existingId = '';
            }
        }

        const created = await graphJson('POST', pathBase, tok, body, graphCalendarWriteHeaders());
        const eventId = created && created.id != null ? String(created.id) : '';
        if (!eventId) throw new Error('Graph hat keine Event-ID zurückgegeben.');
        setOutlookEventId(antragId, eventId);
        return { eventId, created: true };
    } catch (e) {
        throw new Error(formatLfrCalendarAccessError(e, target.useMe ? '' : target.calendarUser));
    }
}

/**
 * @param {object[]} items
 * @param {{ calendarUser?: string, calendarId?: string, signedInEmail?: string, useMe?: boolean, onlyApproved?: boolean }} [opts]
 */
export async function syncLfrItemsToOutlookCalendar(items, opts) {
    const o = opts || {};
    const setup = loadCalendarTargetFromSetup();
    const target = {
        calendarUser: o.calendarUser || setup.calendarUser,
        calendarId: o.calendarId || setup.calendarId,
        signedInEmail: o.signedInEmail || '',
        useMe: !!o.useMe
    };
    const onlyApproved = o.onlyApproved !== false;
    const list = (items || []).filter((it) => {
        const st = normalizeStatus(it.status);
        if (onlyApproved && st !== 'Genehmigt') return false;
        return true;
    });

    let ok = 0;
    let fail = 0;
    const errors = [];
    for (let i = 0; i < list.length; i++) {
        const it = list[i];
        try {
            await upsertLfrOutlookEvent(it, target);
            ok++;
        } catch (e) {
            fail++;
            errors.push(
                (it.lehrerName || it.titel || '?') +
                    ': ' +
                    (e && e.message ? e.message : String(e))
            );
        }
    }
    return { ok, fail, errors, total: list.length };
}
