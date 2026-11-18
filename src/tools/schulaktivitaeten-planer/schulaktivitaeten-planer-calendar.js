/**
 * Genehmigte Schulaktivität → persönlicher Outlook-Kalender / Einladungen / ICS.
 * Nutzt denselben Graph-Stack wie der Planer (`ms365SpoGraph` / spo-graph-shared).
 */
import { enrichClassWithM365Group } from '../../shared/school-stammdaten-spo-lists.js';
import {
    klasseCodesOf,
    parseCodeList,
    toIsoDateOnly
} from './schulaktivitaeten-planer-logic.js';
import { applyUeToDraft, formatUeZeitraum } from './schulaktivitaeten-planer-zeitraster.js';
import {
    formatDeDate,
    labelMaps,
    statusLabel,
    teilnahmeKurz,
    typLabel
} from './schulaktivitaeten-planer-state.js';
import { buildIcs, downloadIcs } from './schulaktivitaeten-planer-export.js';

export const AKT_CALENDAR_TZ = 'Europe/Vienna';

export const AKT_USER_CALENDAR_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Calendars.ReadWrite'
];

/** Kalender schreiben + Gruppen-Mail für Klassen-Teams auflösen. */
export const AKT_INVITE_CALENDAR_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Calendars.ReadWrite',
    'https://graph.microsoft.com/Group.Read.All'
];

const PERSONAL_CAL_LS = 'ms365-akt-personal-cal-v1';

function spoGraph() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api || typeof api.getGraphToken !== 'function') {
        throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared). Bitte Seite neu laden.');
    }
    return api;
}

async function getGraphToken(scopes) {
    return spoGraph().getGraphToken(scopes);
}

/**
 * Graph-JSON mit optionalen Extra-Headern (z. B. Prefer für Zeitzone / Einladungen).
 * spo-graph-shared.graphJson kennt kein Prefer – deshalb lokaler Fetch.
 * @param {string} method
 * @param {string} pathOrUrl
 * @param {string} token
 * @param {unknown} [body]
 * @param {Record<string, string>} [extraHeaders]
 */
async function graphJson(method, pathOrUrl, token, body, extraHeaders) {
    const path = String(pathOrUrl || '');
    const url =
        path.indexOf('http') === 0
            ? path
            : 'https://graph.microsoft.com/v1.0' + (path.indexOf('/') === 0 ? path : '/' + path);
    const headers = Object.assign(
        { Authorization: 'Bearer ' + token },
        extraHeaders && typeof extraHeaders === 'object' ? extraHeaders : {}
    );
    if (body !== undefined) headers['Content-Type'] = 'application/json';
    const res = await fetch(url, {
        method: method,
        headers: headers,
        body: body !== undefined ? JSON.stringify(body) : undefined
    });
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = text;
        }
    }
    if (!res.ok) {
        const msg =
            typeof data === 'object' && data && data.error
                ? data.error.message || JSON.stringify(data.error)
                : text || String(res.status);
        const err = new Error(method + ' ' + path + ': ' + msg);
        err.status = res.status;
        err.payload = data;
        throw err;
    }
    return data || {};
}

function personalCalStorageKey(aktivitaetId, accountEmail) {
    return String(aktivitaetId || '').trim() + '|' + String(accountEmail || '').trim().toLowerCase();
}

/**
 * @param {string} aktivitaetId
 * @param {string} accountEmail
 */
export function getAktPersonalCalendarEventId(aktivitaetId, accountEmail) {
    try {
        const raw = JSON.parse(localStorage.getItem(PERSONAL_CAL_LS) || '{}');
        return String(raw[personalCalStorageKey(aktivitaetId, accountEmail)] || '').trim();
    } catch {
        return '';
    }
}

function setAktPersonalCalendarEventId(aktivitaetId, accountEmail, eventId) {
    try {
        const raw = JSON.parse(localStorage.getItem(PERSONAL_CAL_LS) || '{}');
        const key = personalCalStorageKey(aktivitaetId, accountEmail);
        if (eventId) raw[key] = String(eventId);
        else delete raw[key];
        localStorage.setItem(PERSONAL_CAL_LS, JSON.stringify(raw));
    } catch {
        /* ignore */
    }
}

function normCode(s) {
    return String(s || '')
        .trim()
        .toUpperCase();
}

function isEmail(v) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || '').trim());
}

function graphCalendarHeaders(sendInvites) {
    const prefer =
        'outlook.timezone="' +
        AKT_CALENDAR_TZ +
        '"' +
        (sendInvites
            ? ', outlook.send-event-invitations="true"'
            : ', outlook.send-event-invitations="false"');
    return { Prefer: prefer };
}

/**
 * Folgetag YYYY-MM-DD (Graph all-day Enddatum exklusiv).
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
 * @param {object} item
 * @param {object} [stammdaten]
 */
export function buildAktCalendarSubject(item, stammdaten) {
    const labels = labelMaps(stammdaten || {});
    const normalized = applyUeToDraft(item || {}).draft;
    const teilnahme = teilnahmeKurz(normalized, labels.klasse);
    return (
        typLabel(normalized.typ) +
        ': ' +
        (normalized.titel || 'Aktivität') +
        (teilnahme && teilnahme !== '–' ? ' · ' + teilnahme : '')
    );
}

/**
 * @param {object} item
 * @param {object} [stammdaten]
 */
export function buildAktCalendarBodyText(item, stammdaten) {
    const labels = labelMaps(stammdaten || {});
    const normalized = applyUeToDraft(item || {}).draft;
    const teilnahme = teilnahmeKurz(normalized, labels.klasse);
    const lehrerName =
        labels.lehrer[normalized.lehrerCode] ||
        normalized.lehrerCode ||
        normalized.lehrerEmail ||
        '–';
    const begleitungCodes = parseCodeList(normalized.begleitungCodes || normalized.begleitung);
    const begleitung = begleitungCodes.length
        ? begleitungCodes.map((c) => labels.lehrer[c] || c).join(', ')
        : String(normalized.begleitung || '').trim();
    const ueLabel = formatUeZeitraum(normalized);
    const zeitraum =
        formatDeDate(normalized.startdatum) +
        (normalized.enddatum && normalized.enddatum !== normalized.startdatum
            ? ' – ' + formatDeDate(normalized.enddatum)
            : '') +
        (ueLabel ? ' · ' + ueLabel : '');
    return [
        'Schulaktivität (MS365-Schul-Tools)',
        'Titel: ' + (normalized.titel || '–'),
        'Typ: ' + typLabel(normalized.typ),
        'Status: ' + statusLabel(normalized.status),
        'Teilnahme: ' + (teilnahme || '–'),
        normalized.gruppenText ? 'Gruppe: ' + normalized.gruppenText : '',
        normalized.nichtTeilnehmende ? 'Nicht teilnehmend: ' + normalized.nichtTeilnehmende : '',
        'Verantwortlich: ' + lehrerName,
        begleitung ? 'Begleitung: ' + begleitung : '',
        normalized.ort ? 'Ort: ' + normalized.ort : '',
        'Zeitraum: ' + zeitraum,
        normalized.verkehrsmittel ? 'Verkehrsmittel: ' + normalized.verkehrsmittel : '',
        normalized.kostenHinweis ? 'Kosten: ' + normalized.kostenHinweis : '',
        normalized.ablauf ? 'Ablauf: ' + normalized.ablauf : '',
        normalized.notiz ? 'Notiz: ' + normalized.notiz : ''
    ]
        .filter(Boolean)
        .join('\n');
}

/**
 * @param {object} item
 * @returns {{ isAllDay: boolean, start: {dateTime:string,timeZone:string}, end: {dateTime:string,timeZone:string} }}
 */
export function buildAktGraphTimes(item) {
    const normalized = applyUeToDraft(item || {}).draft;
    const startIso = toIsoDateOnly(normalized.startdatum);
    const endIso = toIsoDateOnly(normalized.enddatum || normalized.startdatum) || startIso;
    if (!startIso) throw new Error('Startdatum fehlt – Kalendereintrag nicht möglich.');

    const startZeit = String(normalized.startZeit || '').trim();
    const endZeit = String(normalized.endZeit || '').trim();
    const timedSameDay =
        startIso === endIso && /^\d{1,2}:\d{2}$/.test(startZeit) && /^\d{1,2}:\d{2}$/.test(endZeit);

    if (timedSameDay) {
        const pad = (t) => {
            const m = /^(\d{1,2}):(\d{2})$/.exec(t);
            return m ? String(m[1]).padStart(2, '0') + ':' + m[2] + ':00' : t + ':00';
        };
        return {
            isAllDay: false,
            start: { dateTime: startIso + 'T' + pad(startZeit), timeZone: AKT_CALENDAR_TZ },
            end: { dateTime: endIso + 'T' + pad(endZeit), timeZone: AKT_CALENDAR_TZ }
        };
    }

    const endExclusive = nextIsoDate(endIso);
    return {
        isAllDay: true,
        start: { dateTime: startIso, timeZone: AKT_CALENDAR_TZ },
        end: { dateTime: endExclusive || nextIsoDate(startIso), timeZone: AKT_CALENDAR_TZ }
    };
}

/**
 * @param {object} item
 * @param {object} [stammdaten]
 * @param {{ attendees?: object[], sendInvites?: boolean }} [opts]
 */
export function buildAktOutlookEvent(item, stammdaten, opts) {
    const o = opts || {};
    const normalized = applyUeToDraft(item || {}).draft;
    const times = buildAktGraphTimes(normalized);
    const attendees = Array.isArray(o.attendees) ? o.attendees : [];
    const sendInvites = !!o.sendInvites && attendees.length > 0;
    const body = {
        subject: buildAktCalendarSubject(normalized, stammdaten),
        body: {
            contentType: 'text',
            content: buildAktCalendarBodyText(normalized, stammdaten)
        },
        isAllDay: times.isAllDay,
        start: times.start,
        end: times.end,
        location: normalized.ort ? { displayName: String(normalized.ort) } : undefined,
        showAs: 'busy',
        categories: ['Schulaktivität', typLabel(normalized.typ)],
        isReminderOn: true,
        reminderMinutesBeforeStart: times.isAllDay ? 720 : 60,
        responseRequested: sendInvites,
        isOnlineMeeting: false,
        allowNewTimeProposals: false
    };
    if (attendees.length) body.attendees = attendees;
    return body;
}

/**
 * Klasse → Graph Group-ID (classTeams / classGroupMatchByKey).
 * @param {string} klasseCode
 * @returns {{ groupId: string, displayName: string, mail: string, mailNickname: string, code: string }|null}
 */
export function lookupClassGroupLink(klasseCode) {
    const code = String(klasseCode || '').trim();
    if (!code) return null;
    if (/[,;~]/.test(code) || /(?:\d+[A-Za-z]+){2,}/.test(code)) return null;

    const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
    if (!api || typeof api.getContainer !== 'function') return null;

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
        if (!match) continue;
        return {
            groupId: t.graphGroupId ? String(t.graphGroupId).trim() : '',
            displayName: dn || code,
            mail: '',
            mailNickname: String(t.mailNickname || t.stableMailNickname || '').trim(),
            code: cc || want
        };
    }

    const setup = typeof api.getSetup === 'function' ? api.getSetup() : null;
    const map = (setup && setup.classGroupMatchByKey) || {};
    const entry = map[want] || map[code] || null;
    if (!entry || typeof entry !== 'object') return null;
    return {
        groupId: String(entry.groupId || entry.graphGroupId || '').trim(),
        displayName: String(entry.displayName || entry.name || code).trim(),
        mail: String(entry.mail || entry.email || '').trim().toLowerCase(),
        mailNickname: String(entry.mailNickname || entry.alias || '').trim(),
        code: want
    };
}

/**
 * @param {string} groupId
 * @param {string} token
 */
async function fetchGroupMail(groupId, token) {
    const gid = String(groupId || '').trim();
    if (!gid) return { mail: '', displayName: '', mailNickname: '' };
    const g = await graphJson(
        'GET',
        '/groups/' + encodeURIComponent(gid) + '?$select=id,displayName,mail,mailNickname',
        token
    );
    return {
        mail: String((g && g.mail) || '')
            .trim()
            .toLowerCase(),
        displayName: String((g && g.displayName) || '').trim(),
        mailNickname: String((g && g.mailNickname) || '').trim()
    };
}

function findStammdatenClass(sd, klasseCode) {
    const want = normCode(klasseCode);
    const list = (sd && Array.isArray(sd.classes) ? sd.classes : []) || [];
    for (let i = 0; i < list.length; i++) {
        const c = list[i];
        if (!c) continue;
        if (normCode(c.code) === want) return enrichClassWithM365Group(c);
        if (String(c.name || '').trim() === String(klasseCode || '').trim()) {
            return enrichClassWithM365Group(c);
        }
    }
    return enrichClassWithM365Group({ code: klasseCode, name: klasseCode });
}

function studentEmailsForKlasse(sd, klasseCode) {
    const want = String(klasseCode || '')
        .trim()
        .toLowerCase();
    const fromClass = findStammdatenClass(sd, klasseCode);
    /** @type {string[]} */
    const out = [];
    const seen = new Set();
    const pushEm = (em) => {
        const e = String(em || '')
            .trim()
            .toLowerCase();
        if (!isEmail(e) || seen.has(e)) return;
        seen.add(e);
        out.push(e);
    };
    (Array.isArray(fromClass.studentEmails) ? fromClass.studentEmails : []).forEach(pushEm);
    (Array.isArray(sd && sd.students) ? sd.students : []).forEach((st) => {
        const sk = String((st && st.klasse) || '')
            .trim()
            .toLowerCase();
        if (!sk || (sk !== want && sk !== String(fromClass.name || '').trim().toLowerCase())) return;
        pushEm(st && st.email);
    });
    return out;
}

/**
 * Einladungsziele: Klassen-M365-Gruppe (SP/Stammdaten/Teams) → sonst Schüler:innen,
 * plus Begleitung und explizite E-Mails im Teilnehmertext.
 * @param {object} item
 * @param {object} [stammdaten]
 * @returns {Promise<{ attendees: object[], resolved: {email:string,label:string,kind:string}[], warnings: string[] }>}
 */
export async function resolveAktInvitees(item, stammdaten) {
    const sd = stammdaten || {};
    const labels = labelMaps(sd);
    const teachers = Array.isArray(sd.teachers) ? sd.teachers : [];
    const teacherByCode = new Map(
        teachers
            .filter((t) => t && t.code)
            .map((t) => [String(t.code).toUpperCase(), t])
    );
    const seen = new Set();
    /** @type {{ email: string, label: string, kind: string }[]} */
    const resolved = [];
    /** @type {string[]} */
    const warnings = [];

    const push = (email, label, kind) => {
        const em = String(email || '')
            .trim()
            .toLowerCase();
        if (!isEmail(em) || seen.has(em)) return;
        seen.add(em);
        resolved.push({ email: em, label: label || em, kind: kind || 'person' });
    };

    const tok = await getGraphToken(AKT_INVITE_CALENDAR_SCOPES);
    const klassen = klasseCodesOf(item);
    for (let i = 0; i < klassen.length; i++) {
        const code = klassen[i];
        const klasseRow = findStammdatenClass(sd, code);
        const link = lookupClassGroupLink(code);
        let groupId = String(klasseRow.m365GroupId || (link && link.groupId) || '').trim();
        let mail = String(klasseRow.m365GroupMail || (link && link.mail) || '')
            .trim()
            .toLowerCase();
        let label =
            klasseRow.m365GroupName ||
            (link && link.displayName) ||
            labels.klasse[code] ||
            code;

        if (!mail && groupId) {
            try {
                const g = await fetchGroupMail(groupId, tok);
                mail = g.mail;
                if (g.displayName) label = g.displayName;
            } catch (e) {
                warnings.push(
                    'Gruppe für „' +
                        code +
                        '“ nicht lesbar: ' +
                        (e && e.message ? e.message : String(e))
                );
            }
        }

        if (mail) {
            push(mail, label, 'group');
            continue;
        }

        const students = studentEmailsForKlasse(sd, code);
        if (students.length) {
            students.forEach((em) => push(em, labels.klasse[code] || code, 'student'));
            warnings.push(
                'Klasse „' +
                    code +
                    '“: keine M365-Gruppen-Mail – ' +
                    students.length +
                    ' Schüler:innen einzeln eingeladen.'
            );
        } else {
            warnings.push(
                'Keine M365-Gruppe für Klasse „' +
                    code +
                    '“ verknüpft (SharePoint-Spalte M365-Gruppe / Klassen-Teams) und keine Schüler-E-Mails.'
            );
        }
    }

    const begleitung = parseCodeList((item && item.begleitungCodes) || (item && item.begleitung));
    begleitung.forEach((code) => {
        const t = teacherByCode.get(String(code).toUpperCase());
        const em = t && t.email ? String(t.email).trim().toLowerCase() : '';
        if (em) push(em, (t && t.name) || labels.lehrer[code] || code, 'teacher');
        else warnings.push('Begleitung „' + code + '“ ohne E-Mail in den Stammdaten.');
    });

    const teilText = String((item && item.teilnehmendeText) || '');
    const emailHits = teilText.match(/[^\s,;]+@[^\s,;]+/g) || [];
    emailHits.forEach((em) => push(em, em, 'person'));

    const attendees = resolved.map((r) => ({
        emailAddress: { address: r.email, name: r.label },
        type: 'required'
    }));

    return { attendees, resolved, warnings };
}

/**
 * @param {object} item
 * @param {{ accountEmail: string, stammdaten?: object, invite?: boolean }} opts
 */
export async function upsertAktToMyCalendar(item, opts) {
    const email = String((opts && opts.accountEmail) || '')
        .trim()
        .toLowerCase();
    if (!email) {
        throw new Error('Bitte anmelden – der Outlook-Kalender braucht Ihre Microsoft-Konto-E-Mail.');
    }
    const aktivitaetId = String((item && (item.aktivitaetId || item.itemId)) || '').trim();
    if (!aktivitaetId) throw new Error('Aktivitäts-ID fehlt.');
    if (String((item && item.status) || '').toLowerCase() !== 'genehmigt') {
        throw new Error('Nur genehmigte Aktivitäten können in den Kalender übernommen werden.');
    }

    const invite = !!(opts && opts.invite);
    let attendees = [];
    let inviteMeta = { resolved: [], warnings: [] };
    if (invite) {
        inviteMeta = await resolveAktInvitees(item, opts && opts.stammdaten);
        attendees = inviteMeta.attendees || [];
        if (!attendees.length) {
            const detail = (inviteMeta.warnings || []).slice(0, 3).join(' ');
            throw new Error(
                'Keine Einladungs-Adressen gefunden. ' +
                    (detail ||
                        'Bitte Klassen-Teams in den Stammdaten verknüpfen oder Begleitung mit E-Mail hinterlegen.')
            );
        }
    }

    const body = buildAktOutlookEvent(item, opts && opts.stammdaten, {
        attendees,
        sendInvites: invite
    });
    const scopes = invite ? AKT_INVITE_CALENDAR_SCOPES : AKT_USER_CALENDAR_SCOPES;
    const tok = await getGraphToken(scopes);
    let existingId = getAktPersonalCalendarEventId(aktivitaetId, email);

    if (existingId) {
        try {
            await graphJson(
                'PATCH',
                '/me/calendar/events/' + encodeURIComponent(existingId),
                tok,
                body,
                graphCalendarHeaders(invite)
            );
            return {
                eventId: existingId,
                created: false,
                invited: invite ? attendees.length : 0,
                resolved: inviteMeta.resolved,
                warnings: inviteMeta.warnings
            };
        } catch (err) {
            const msg = err && err.message ? String(err.message) : String(err);
            if (!/404|ErrorItemNotFound|not found/i.test(msg)) throw err;
            existingId = '';
        }
    }

    const created = await graphJson(
        'POST',
        '/me/calendar/events',
        tok,
        body,
        graphCalendarHeaders(invite)
    );
    const eventId = created && created.id != null ? String(created.id) : '';
    if (!eventId) throw new Error('Graph hat keine Event-ID zurückgegeben.');
    setAktPersonalCalendarEventId(aktivitaetId, email, eventId);
    return {
        eventId,
        created: true,
        invited: invite ? attendees.length : 0,
        resolved: inviteMeta.resolved,
        warnings: inviteMeta.warnings,
        webLink: created && created.webLink ? String(created.webLink) : ''
    };
}

/**
 * Einzelne Aktivität als .ics herunterladen.
 * @param {object} item
 * @param {object} [stammdaten]
 */
export function downloadAktIcs(item, stammdaten) {
    if (!item) throw new Error('Keine Aktivität gewählt.');
    const ics = buildIcs([item], stammdaten, 'Schulaktivität', { onlyApproved: false });
    const raw = String(item.titel || item.aktivitaetId || 'schulaktivitaet')
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9äöüß]+/gi, '-')
        .replace(/^-+|-+$/g, '')
        .slice(0, 48);
    downloadIcs(ics, (raw || 'schulaktivitaet') + '.ics');
}

/**
 * Kurzinfo für die UI: welche Ziele wären einladbar (ohne Graph-Call für Gruppen-Mail).
 * @param {object} item
 * @param {object} [stammdaten]
 */
export function previewAktInviteTargets(item, stammdaten) {
    const sd = stammdaten || {};
    const labels = labelMaps(sd);
    const teachers = Array.isArray(sd.teachers) ? sd.teachers : [];
    const teacherByCode = new Map(
        teachers.filter((t) => t && t.code).map((t) => [String(t.code).toUpperCase(), t])
    );
    /** @type {{ label: string, detail: string, ok: boolean }[]} */
    const rows = [];
    klasseCodesOf(item).forEach((code) => {
        const klasseRow = findStammdatenClass(sd, code);
        const link = lookupClassGroupLink(code);
        const mail = String(klasseRow.m365GroupMail || (link && link.mail) || '')
            .trim()
            .toLowerCase();
        const groupId = String(klasseRow.m365GroupId || (link && link.groupId) || '').trim();
        if (mail || groupId) {
            rows.push({
                label: labels.klasse[code] || code,
                detail: mail
                    ? mail
                    : groupId
                      ? 'M365-Gruppe verknüpft (Mail wird geladen)'
                      : 'ohne Mail',
                ok: true
            });
            return;
        }
        const students = studentEmailsForKlasse(sd, code);
        if (students.length) {
            rows.push({
                label: labels.klasse[code] || code,
                detail: students.length + ' Schüler:innen (einzeln, keine Gruppen-Mail)',
                ok: true
            });
        } else {
            rows.push({
                label: labels.klasse[code] || code,
                detail: 'Keine M365-Gruppe / keine Schüler-E-Mails',
                ok: false
            });
        }
    });
    parseCodeList((item && item.begleitungCodes) || (item && item.begleitung)).forEach((code) => {
        const t = teacherByCode.get(String(code).toUpperCase());
        const em = t && t.email ? String(t.email).trim() : '';
        rows.push({
            label: (t && t.name) || labels.lehrer[code] || code,
            detail: em || 'keine E-Mail',
            ok: !!em
        });
    });
    return rows;
}
