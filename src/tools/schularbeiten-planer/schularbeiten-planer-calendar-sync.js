/**
 * Sync fixierter Schularbeiten → Gruppenkalender der Klassen-Teams (Graph).
 */
import { getGraphToken, graphJson } from '../../shared/graph-client.js';
import {
    toIsoDateOnly,
    formatSchularbeitZeitspanne,
    schularbeitCalendarSubject,
    schularbeitGraphCalendarTimes,
    SCHULARBEIT_CALENDAR_TZ
} from './schularbeiten-planer-logic.js';
import { saMarker } from './schularbeiten-planer-sync.js';
import { formatDeDate, isOwnSchularbeit } from './schularbeiten-planer-state.js';

export const CALENDAR_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.ReadWrite.All'
];

/** Persönlicher Outlook-Kalender des angemeldeten Benutzers. */
export const USER_CALENDAR_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Calendars.ReadWrite'
];

const PERSONAL_CAL_LS = 'ms365-sa-personal-cal-v1';

function personalCalStorageKey(schularbeitId, accountEmail) {
    return String(schularbeitId || '').trim() + '|' + String(accountEmail || '').trim().toLowerCase();
}

/**
 * @param {string} schularbeitId
 * @param {string} accountEmail
 */
export function countPersonalCalendarLinked(items, accountEmail) {
    const email = String(accountEmail || '').trim().toLowerCase();
    if (!email) return 0;
    let n = 0;
    (items || []).forEach((sa) => {
        if (getPersonalCalendarEventId(sa.schularbeitId, email)) n++;
    });
    return n;
}

export function getPersonalCalendarEventId(schularbeitId, accountEmail) {
    try {
        const raw = JSON.parse(localStorage.getItem(PERSONAL_CAL_LS) || '{}');
        return String(raw[personalCalStorageKey(schularbeitId, accountEmail)] || '').trim();
    } catch {
        return '';
    }
}

function setPersonalCalendarEventId(schularbeitId, accountEmail, eventId) {
    try {
        const raw = JSON.parse(localStorage.getItem(PERSONAL_CAL_LS) || '{}');
        const key = personalCalStorageKey(schularbeitId, accountEmail);
        if (eventId) raw[key] = String(eventId);
        else delete raw[key];
        localStorage.setItem(PERSONAL_CAL_LS, JSON.stringify(raw));
    } catch {
        /* ignore */
    }
}

const TZ = SCHULARBEIT_CALENDAR_TZ;

/** Graph-Header: Zeitzone + keine Einladungs-/Update-Benachrichtigungen an Teilnehmer (es gibt keine). */
function graphCalendarWriteHeaders() {
    return {
        Prefer: 'outlook.timezone="' + TZ + '", outlook.send-event-invitations="false"'
    };
}

function normCode(s) {
    return String(s || '')
        .trim()
        .toUpperCase();
}

/**
 * Folgetag als YYYY-MM-DD (all-day Enddatum für Graph).
 * @param {string} isoDate
 */
export function nextIsoDate(isoDate) {
    const d = toIsoDateOnly(isoDate);
    if (!d) return '';
    const dt = new Date(d + 'T12:00:00Z');
    dt.setUTCDate(dt.getUTCDate() + 1);
    const y = dt.getUTCFullYear();
    const m = String(dt.getUTCMonth() + 1).padStart(2, '0');
    const day = String(dt.getUTCDate()).padStart(2, '0');
    return y + '-' + m + '-' + day;
}

/**
 * Klasse → Graph Group-ID (classTeams, sonst classGroupMatchByKey).
 * @param {string} klasseCode
 * @returns {string}
 */
export function resolveClassGroupId(klasseCode) {
    const code = String(klasseCode || '').trim();
    if (!code) {
        throw new Error('Klassen-Code fehlt – Kalender-Sync nicht möglich.');
    }
    if (/[,;~]/.test(code) || /(?:\d+[A-Za-z]+){2,}/.test(code)) {
        throw new Error(
            'Klasse „' + code + '“ ist kein einzelnes Klassenteam (Mehrklassen-Code). Bitte in Stammdaten prüfen.'
        );
    }

    const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
    if (!api || typeof api.getContainer !== 'function') {
        throw new Error('App-Daten (ms365AppDataV2) nicht geladen.');
    }

    const container = api.getContainer();
    const teams =
        typeof api.normalizeCoreClassTeams === 'function'
            ? api.normalizeCoreClassTeams((container.core && container.core.classTeams) || [])
            : (container.core && container.core.classTeams) || [];

    const want = normCode(code);
    for (let i = 0; i < teams.length; i++) {
        const t = teams[i];
        if (!t) continue;
        const cc = normCode(t.classCode || '');
        const dn = String(t.displayName || '').trim();
        const match = (dn && code === dn) || (cc && want === cc);
        const gid = t.graphGroupId ? String(t.graphGroupId).trim() : '';
        if (match && gid) return gid;
    }

    const setup = typeof api.getSetup === 'function' ? api.getSetup() : null;
    const map = (setup && setup.classGroupMatchByKey) || {};
    const entry = map[want] || map[code] || null;
    const fromMatch = entry && (entry.groupId || entry.graphGroupId);
    if (fromMatch) return String(fromMatch).trim();

    throw new Error(
        'Keine Teams-Gruppe für Klasse „' +
            code +
            '“ verknüpft. Bitte in Stammdaten (Klassen-Teams / Gruppenabgleich) zuordnen.'
    );
}

/**
 * @param {object} sa
 * @param {{ fach?: string, klasse?: string }} [labels]
 */
export function buildGroupCalendarEvent(sa, labels) {
    const datum = toIsoDateOnly(sa.datum);
    if (!datum) throw new Error('Datum fehlt für Kalender-Event.');
    const times = schularbeitGraphCalendarTimes(sa, datum);
    const subject = schularbeitCalendarSubject(sa, {
        fach: labels && labels.fach,
        klasse: labels && labels.klasse
    });
    const zeit = formatSchularbeitZeitspanne(sa);
    const marker = saMarker(sa.schularbeitId);
    const bodyLines = [
        marker,
        'Schularbeit',
        zeit ? 'Zeit: ' + zeit : '',
        sa.lehrerCode ? 'Lehrer: ' + sa.lehrerCode : '',
        sa.dauerMinuten ? 'Dauer: ' + sa.dauerMinuten + ' Min.' : '',
        sa.notiz ? String(sa.notiz).trim() : ''
    ].filter(Boolean);

    // dateTime = Wanduhrzeit in `timeZone` (Graph/Exchange wendet CET/CEST inkl. DST an).
    return {
        subject,
        body: {
            contentType: 'text',
            content: bodyLines.join(' · ')
        },
        isAllDay: times.isAllDay,
        start: {
            dateTime: times.startDateTime,
            timeZone: TZ
        },
        end: {
            dateTime: times.endDateTime,
            timeZone: TZ
        },
        showAs: 'busy',
        categories: ['Schularbeit'],
        /** Keine Outlook-Erinnerung (weder für Organisator noch bei Gruppenkalender). */
        isReminderOn: false,
        reminderMinutesBeforeStart: 0,
        /** Keine Meeting-Einladungen / keine Antwortanfragen – Termin nur im Kalender. */
        responseRequested: false,
        isOnlineMeeting: false,
        allowNewTimeProposals: false
    };
}

/**
 * @param {object} sa
 * @param {{ fachLabel?: string, klasseLabel?: string }} [opts]
 * @returns {Promise<{ eventId: string, created: boolean, groupId: string }>}
 */
export async function upsertGroupCalendarEvent(sa, opts) {
    if (!sa || !sa.schularbeitId) throw new Error('SchularbeitId fehlt für Klassenkalender-Sync.');
    const groupId = resolveClassGroupId(sa.klasseCode);
    const body = buildGroupCalendarEvent(sa, {
        fach: opts && opts.fachLabel,
        klasse: opts && opts.klasseLabel
    });
    const tok = await getGraphToken(CALENDAR_SCOPES);
    const existingId = sa.teamsCalendarEventId ? String(sa.teamsCalendarEventId).trim() : '';

    if (existingId) {
        try {
            await graphJson(
                'PATCH',
                '/groups/' + encodeURIComponent(groupId) + '/calendar/events/' + encodeURIComponent(existingId),
                tok,
                body,
                graphCalendarWriteHeaders()
            );
            return { eventId: existingId, created: false, groupId };
        } catch (err) {
            const msg = err && err.message ? String(err.message) : String(err);
            // Event gelöscht / andere Gruppe → neu anlegen
            if (!/404|ErrorItemNotFound|not found/i.test(msg)) throw err;
        }
    }

    const created = await graphJson(
        'POST',
        '/groups/' + encodeURIComponent(groupId) + '/calendar/events',
        tok,
        body,
        graphCalendarWriteHeaders()
    );
    const eventId = created && created.id != null ? String(created.id) : '';
    if (!eventId) throw new Error('Graph hat keine Event-ID zurückgegeben.');
    return { eventId, created: true, groupId };
}

/**
 * @param {string} groupId
 * @param {string} eventId
 */
export async function deleteGroupCalendarEvent(groupId, eventId) {
    const gid = String(groupId || '').trim();
    const eid = String(eventId || '').trim();
    if (!gid || !eid) return;
    const tok = await getGraphToken(CALENDAR_SCOPES);
    try {
        await graphJson(
            'DELETE',
            '/groups/' + encodeURIComponent(gid) + '/calendar/events/' + encodeURIComponent(eid),
            tok
        );
    } catch (err) {
        const msg = err && err.message ? String(err.message) : String(err);
        if (/404|ErrorItemNotFound|not found/i.test(msg)) return;
        throw err;
    }
}

/**
 * Löscht Event anhand gespeicherter ID; groupId wird aus Klasse aufgelöst.
 * @param {object} sa
 */
/**
 * @param {string} klasseCode
 * @returns {{ ok: boolean, groupId?: string, teamName?: string, message: string }}
 */
export function describeClassGroupLink(klasseCode) {
    const code = String(klasseCode || '').trim();
    if (!code) {
        return { ok: false, message: 'Bitte eine Klasse im Filter wählen oder Termine mit Klassen-Code laden.' };
    }
    try {
        const groupId = resolveClassGroupId(code);
        let teamName = '';
        const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
        if (api && typeof api.getContainer === 'function') {
            const container = api.getContainer();
            const teams =
                typeof api.normalizeCoreClassTeams === 'function'
                    ? api.normalizeCoreClassTeams((container.core && container.core.classTeams) || [])
                    : (container.core && container.core.classTeams) || [];
            const want = normCode(code);
            for (let i = 0; i < teams.length; i++) {
                const t = teams[i];
                if (!t) continue;
                const cc = normCode(t.classCode || '');
                const dn = String(t.displayName || '').trim();
                const match = (dn && code === dn) || (cc && want === cc);
                if (match && String(t.graphGroupId || '').trim() === groupId) {
                    teamName = dn || String(t.mailNickname || '').trim();
                    break;
                }
            }
        }
        const label = teamName ? '„' + teamName + '“' : 'Gruppe ' + groupId.slice(0, 8) + '…';
        return {
            ok: true,
            groupId,
            teamName,
            message: 'Kalender der Klasse ' + code + ' → ' + label
        };
    } catch (e) {
        return { ok: false, message: e && e.message ? e.message : String(e) };
    }
}

/**
 * Fixierte Schularbeiten in die jeweiligen Microsoft-365-Gruppenkalender schreiben (Upsert).
 * @param {object[]} items
 * @param {{ fach?: Record<string,string>, klasse?: Record<string,string>, lehrer?: Record<string,string>, persistEventId?: (sa: object, eventId: string) => Promise<void> }} labelMapsOrOpts
 */
/**
 * @param {object} sa
 * @param {{ accountEmail: string, fachLabel?: string, klasseLabel?: string }} opts
 */
export async function upsertUserCalendarEvent(sa, opts) {
    const email = String((opts && opts.accountEmail) || '').trim().toLowerCase();
    if (!email) throw new Error('Bitte anmelden – persönlicher Kalender braucht Ihre Microsoft-Konto-E-Mail.');
    if (!sa || !sa.schularbeitId) throw new Error('SchularbeitId fehlt für Kalender-Sync.');

    const body = buildGroupCalendarEvent(sa, {
        fach: opts && opts.fachLabel,
        klasse: opts && opts.klasseLabel
    });
    const tok = await getGraphToken(USER_CALENDAR_SCOPES);
    let existingId = getPersonalCalendarEventId(sa.schularbeitId, email);

    if (existingId) {
        try {
            await graphJson(
                'PATCH',
                '/me/calendar/events/' + encodeURIComponent(existingId),
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

    const created = await graphJson('POST', '/me/calendar/events', tok, body, graphCalendarWriteHeaders());
    const eventId = created && created.id != null ? String(created.id) : '';
    if (!eventId) throw new Error('Graph hat keine Event-ID zurückgegeben.');
    setPersonalCalendarEventId(sa.schularbeitId, email, eventId);
    return { eventId, created: true };
}

/**
 * Eigene Schularbeiten → Outlook-Kalender von /me (nur angemeldete Person).
 * @param {object[]} items
 * @param {{ accountEmail: string, scope?: object, fach?: Record<string,string>, klasse?: Record<string,string>, includeStatuses?: string[] }} opts
 */
export async function syncSchularbeitenToUserCalendar(items, opts) {
    const email = String((opts && opts.accountEmail) || '').trim().toLowerCase();
    const scope = opts && opts.scope;
    const fach = (opts && opts.fach) || {};
    const klasse = (opts && opts.klasse) || {};
    const includeStatuses = (opts && opts.includeStatuses) || ['fixiert', 'beantragt'];
    const statusSet = new Set(includeStatuses.map((s) => String(s).toLowerCase()));

    const list = (items || []).filter((sa) => {
        if (!statusSet.has(String(sa.status || '').toLowerCase())) return false;
        if (scope && !isOwnSchularbeit(sa, scope)) return false;
        return true;
    });

    let ok = 0;
    let fail = 0;
    const errors = [];

    for (let i = 0; i < list.length; i++) {
        const sa = list[i];
        try {
            await upsertUserCalendarEvent(sa, {
                accountEmail: email,
                fachLabel: fach[sa.fachCode] || sa.fachCode,
                klasseLabel: klasse[sa.klasseCode] || sa.klasseCode
            });
            ok++;
        } catch (e) {
            fail++;
            const msg = e && e.message ? e.message : String(e);
            errors.push(
                formatDeDate(sa.datum) + ' · ' + (sa.fachCode || '') + ': ' + msg
            );
        }
    }

    return { ok, fail, errors, total: list.length };
}

export async function syncSchularbeitenToGroupCalendars(items, labelMapsOrOpts) {
    const labels = labelMapsOrOpts || {};
    const fach = labels.fach || {};
    const klasse = labels.klasse || {};
    const persist = labels.persistEventId;
    const fixed = (items || []).filter((sa) => String(sa.status || '').toLowerCase() === 'fixiert');
    let ok = 0;
    let fail = 0;
    const errors = [];

    for (let i = 0; i < fixed.length; i++) {
        const sa = fixed[i];
        try {
            const result = await upsertGroupCalendarEvent(sa, {
                fachLabel: fach[sa.fachCode] || sa.fachCode,
                klasseLabel: klasse[sa.klasseCode] || sa.klasseCode
            });
            if (typeof persist === 'function' && result.eventId) {
                await persist(sa, result.eventId);
            }
            ok++;
        } catch (e) {
            fail++;
            const msg = e && e.message ? e.message : String(e);
            errors.push(
                (sa.klasseCode || '?') +
                    ' · ' +
                    formatDeDate(sa.datum) +
                    ' · ' +
                    (sa.fachCode || '') +
                    ': ' +
                    msg
            );
        }
    }

    return { ok, fail, errors, total: fixed.length };
}

export async function removeGroupCalendarEventForSa(sa) {
    const eventId = sa && sa.teamsCalendarEventId ? String(sa.teamsCalendarEventId).trim() : '';
    if (!eventId) return { removed: false };
    let groupId = '';
    try {
        groupId = resolveClassGroupId(sa.klasseCode);
    } catch {
        return { removed: false, skipped: true };
    }
    await deleteGroupCalendarEvent(groupId, eventId);
    return { removed: true, groupId, eventId };
}
