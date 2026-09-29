/**
 * Sync fixierter Schularbeiten → Gruppenkalender der Klassen-Teams (Graph).
 */
import { getGraphToken, graphJson } from '../../shared/graph-client.js';
import { toIsoDateOnly } from './schularbeiten-planer-logic.js';
import { saMarker } from './schularbeiten-planer-sync.js';

export const CALENDAR_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.ReadWrite.All'
];

const TZ = 'Europe/Vienna';

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
    const endDate = nextIsoDate(datum);
    const fach = (labels && labels.fach) || sa.fachCode || 'Fach';
    const klasse = (labels && labels.klasse) || sa.klasseCode || '';
    const subject =
        fach + (klasse ? ' · ' + klasse : '') + (sa.thema ? ' – ' + String(sa.thema).slice(0, 80) : '');
    const marker = saMarker(sa.schularbeitId);
    const bodyLines = [
        marker,
        'Schularbeit',
        sa.lehrerCode ? 'Lehrer: ' + sa.lehrerCode : '',
        sa.dauerMinuten ? sa.dauerMinuten + ' Min.' : '',
        sa.notiz ? String(sa.notiz).trim() : ''
    ].filter(Boolean);

    return {
        subject: subject.slice(0, 250),
        body: {
            contentType: 'text',
            content: bodyLines.join(' · ')
        },
        isAllDay: true,
        start: {
            dateTime: datum + 'T00:00:00.0000000',
            timeZone: TZ
        },
        end: {
            dateTime: endDate + 'T00:00:00.0000000',
            timeZone: TZ
        },
        showAs: 'busy',
        categories: ['Schularbeit']
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
                body
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
        body
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
