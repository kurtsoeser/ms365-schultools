/**
 * Graph: Dach-Team (Maturajahrgang) + privater Kanal je Arbeit.
 */
import {
    addChannelMember,
    createPrivateChannel,
    ensureTeamMembers,
    listChannelMembers,
    patchChannelDisplayName,
    removeChannelMemberByUserId
} from '../klassenchats/klassenchats-graph.js';
import { sanitizeTeamsChannelDisplayName } from '../klassenchats/klassenchats-logic.js';
import {
    buildCohortTeamDisplayName,
    buildCohortTeamMailNickname,
    buildPrivateChannelDisplayName,
    labelsForArbeitTyp
} from './diplomarbeiten-logic.js';
import { loadDiplomSettings } from './diplomarbeiten-settings.js';
import { getArbeit, getCohort, patchArbeit, upsertCohort } from './diplomarbeiten-store.js';

function G() {
    return typeof window !== 'undefined' ? window.ms365GraphUnifiedGroups : null;
}

async function token(scopes) {
    const api = G();
    if (!api || typeof api.getGraphToken !== 'function') {
        throw new Error('Graph-Modul nicht geladen.');
    }
    return api.getGraphToken(scopes);
}

async function resolveUserId(tok, email) {
    const api = G();
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return '';
    if (typeof api.resolveUserIdByEmail === 'function') {
        return (await api.resolveUserIdByEmail(tok, em)) || '';
    }
    try {
        const data = await api.graphJson(
            'GET',
            "/users?$filter=mail eq '" +
                em.replace(/'/g, "''") +
                "' or userPrincipalName eq '" +
                em.replace(/'/g, "''") +
                "'&$select=id&$top=1",
            tok,
            undefined
        );
        const u = data && data.value && data.value[0];
        return u && u.id ? String(u.id) : '';
    } catch {
        return '';
    }
}

const TEAM_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.ReadWrite.All',
    'https://graph.microsoft.com/Team.Create',
    'https://graph.microsoft.com/Channel.Create',
    'https://graph.microsoft.com/ChannelMember.ReadWrite.All',
    'https://graph.microsoft.com/Directory.Read.All'
];

/**
 * Dach-Team für Maturajahrgang anlegen und am Cohort speichern.
 * @param {string} cohortId
 * @param {{ onLog?: (m: string) => void }} [opts]
 */
export async function provisionCohortTeam(cohortId, opts) {
    const log = (m) => {
        if (opts && typeof opts.onLog === 'function') opts.onLog(m);
    };
    const cohort = getCohort(cohortId);
    if (!cohort) throw new Error('Jahrgang nicht gefunden.');
    if (cohort.teamId) {
        log('Team bereits verknüpft: ' + cohort.teamId);
        return { ok: true, cohort };
    }

    const settings = loadDiplomSettings();
    const L = labelsForArbeitTyp(settings.arbeitTyp);
    const displayName =
        cohort.displayName ||
        buildCohortTeamDisplayName({
            maturajahrgang: cohort.maturajahrgang,
            arbeitTyp: settings.arbeitTyp
        });
    const mailNickname =
        cohort.mailNickname ||
        buildCohortTeamMailNickname({
            maturajahrgang: cohort.maturajahrgang,
            arbeitTyp: settings.arbeitTyp
        });

    const api = G();
    if (!api) throw new Error('Graph-Modul nicht geladen.');
    const tok = await token(TEAM_SCOPES);
    log('Lege Dach-Team an: ' + displayName);
    const group = await api.createUnifiedGroup(
        tok,
        displayName,
        mailNickname,
        L.plural + ' Maturajahrgang ' + cohort.maturajahrgang + ' (MS365-Schul-Tools)'
    );
    const gid = group && group.id ? String(group.id) : '';
    if (!gid) throw new Error('Gruppe ohne ID.');
    log('Provisioniere Team …');
    await api.provisionTeamForGroup(tok, gid);

    let webUrl = '';
    try {
        const team = await api.graphJson(
            'GET',
            '/teams/' + encodeURIComponent(gid) + '?$select=id,webUrl',
            tok,
            undefined
        );
        webUrl = team && team.webUrl ? String(team.webUrl) : '';
    } catch {
        webUrl = '';
    }

    const saved = upsertCohort({
        ...cohort,
        displayName,
        mailNickname,
        teamId: gid,
        teamWebUrl: webUrl
    });
    return { ok: true, cohort: saved.item };
}

/**
 * Privaten Kanal für eine Arbeit anlegen.
 * @param {string} arbeitId
 * @param {{ onLog?: (m: string) => void }} [opts]
 */
export async function provisionArbeitChannel(arbeitId, opts) {
    const log = (m) => {
        if (opts && typeof opts.onLog === 'function') opts.onLog(m);
    };
    const arbeit = getArbeit(arbeitId);
    if (!arbeit) throw new Error('Arbeit nicht gefunden.');
    const cohort = getCohort(arbeit.cohortId);
    if (!cohort || !cohort.teamId) {
        throw new Error('Zuerst Dach-Team für den Maturajahrgang anlegen/verknüpfen.');
    }
    if (arbeit.channelId) {
        log('Kanal bereits vorhanden.');
        return { ok: true, arbeit };
    }

    const settings = loadDiplomSettings();
    const channelName = sanitizeTeamsChannelDisplayName(
        arbeit.channelName ||
            buildPrivateChannelDisplayName({
                topic: arbeit.topic,
                student: arbeit.student,
                arbeitTyp: settings.arbeitTyp
            })
    );
    if (!channelName) throw new Error('Kanalname fehlt oder ist ungültig.');

    const api = G();
    const tok = await token(TEAM_SCOPES);

    const ownerIds = [];
    const memberIds = [];
    if (arbeit.mentorEmail) {
        const uid = await resolveUserId(tok, arbeit.mentorEmail);
        if (uid) ownerIds.push(uid);
    }
    for (const em of arbeit.studentEmails || []) {
        const uid = await resolveUserId(tok, em);
        if (uid) memberIds.push(uid);
    }
    if (!ownerIds.length) {
        throw new Error('Betreuer-E-Mail muss im Tenant auflösbar sein (Kanal-Besitzer).');
    }

    const roster = Array.from(new Set(ownerIds.concat(memberIds)));
    log('Stelle Team-Mitgliedschaft sicher …');
    await ensureTeamMembers(api, tok, cohort.teamId, roster);

    log('Privater Kanal: ' + channelName);
    const ch = await createPrivateChannel(
        api,
        tok,
        cohort.teamId,
        channelName,
        { ownerUserIds: ownerIds, memberUserIds: memberIds },
        {
            description:
                (arbeit.topic || arbeit.student || 'Abschlussarbeit') +
                (arbeit.mentorName ? ' · Betreuung: ' + arbeit.mentorName : ''),
            onMembershipPending: async () => {
                await ensureTeamMembers(api, tok, cohort.teamId, roster);
            }
        }
    );
    const channelId = ch && ch.id ? String(ch.id) : '';
    const channelWebUrl = ch && ch.webUrl ? String(ch.webUrl) : '';
    if (!channelId) throw new Error('Kanal ohne ID.');

    const patched = patchArbeit(arbeitId, {
        channelId,
        channelName,
        channelWebUrl,
        status: 'aktiv'
    });
    return { ok: true, arbeit: patched.item };
}

/**
 * Bestehenden privaten Kanal aktualisieren: Name + Mitglieder nach SOLL.
 * @param {string} arbeitId
 * @param {{ onLog?: (m: string) => void }} [opts]
 */
export async function syncArbeitChannel(arbeitId, opts) {
    const log = (m) => {
        if (opts && typeof opts.onLog === 'function') opts.onLog(m);
    };
    const arbeit = getArbeit(arbeitId);
    if (!arbeit) throw new Error('Arbeit nicht gefunden.');
    const cohort = getCohort(arbeit.cohortId);
    if (!cohort || !cohort.teamId) throw new Error('Dach-Team fehlt.');
    if (!arbeit.channelId) throw new Error('Noch kein Kanal – bitte zuerst anlegen.');

    const channelName = sanitizeTeamsChannelDisplayName(arbeit.channelName || '');
    if (!channelName) throw new Error('Kanalname fehlt oder ist ungültig.');

    const api = G();
    const tok = await token(TEAM_SCOPES);

    const ownerIds = [];
    const memberIds = [];
    if (arbeit.mentorEmail) {
        const uid = await resolveUserId(tok, arbeit.mentorEmail);
        if (uid) ownerIds.push(uid);
    }
    for (const em of arbeit.studentEmails || []) {
        const uid = await resolveUserId(tok, em);
        if (uid) memberIds.push(uid);
    }
    if (!ownerIds.length) throw new Error('Betreuung muss im Tenant auflösbar sein.');

    const want = new Set(ownerIds.concat(memberIds));
    log('Team-Roster …');
    await ensureTeamMembers(api, tok, cohort.teamId, Array.from(want));

    log('Kanalnamen setzen: ' + channelName);
    await patchChannelDisplayName(api, tok, cohort.teamId, arbeit.channelId, channelName);

    log('Mitglieder abgleichen …');
    let existing = [];
    try {
        existing = await listChannelMembers(api, tok, cohort.teamId, arbeit.channelId);
    } catch (e) {
        throw new Error(
            'Kanalmitglieder lesen fehlgeschlagen: ' + (e && e.message ? e.message : String(e))
        );
    }
    const existingIds = new Set();
    existing.forEach((m) => {
        const uid = String((m && (m.userId || m.id)) || '').trim();
        if (uid) existingIds.add(uid);
    });

    for (const uid of ownerIds) {
        if (existingIds.has(uid)) continue;
        await addChannelMember(api, tok, cohort.teamId, arbeit.channelId, uid, true);
        existingIds.add(uid);
        log('+ Owner');
    }
    for (const uid of memberIds) {
        if (existingIds.has(uid) || ownerIds.indexOf(uid) !== -1) continue;
        await addChannelMember(api, tok, cohort.teamId, arbeit.channelId, uid, false);
        existingIds.add(uid);
        log('+ Mitglied');
    }
    for (const uid of Array.from(existingIds)) {
        if (want.has(uid)) continue;
        const ok = await removeChannelMemberByUserId(api, tok, cohort.teamId, arbeit.channelId, uid);
        if (ok) log('− Mitglied entfernt');
    }

    const patched = patchArbeit(arbeitId, {
        channelName,
        status: arbeit.status === 'geplant' ? 'aktiv' : arbeit.status
    });
    return { ok: true, arbeit: patched.item };
}

/**
 * Team-ID manuell am Jahrgang speichern (bestehendes Team).
 * @param {string} cohortId
 * @param {string} teamId
 * @param {string} [webUrl]
 * @param {{ displayName?: string, mailNickname?: string }} [meta]
 */
export function linkCohortTeam(cohortId, teamId, webUrl, meta) {
    const cohort = getCohort(cohortId);
    if (!cohort) throw new Error('Jahrgang nicht gefunden.');
    return upsertCohort({
        ...cohort,
        displayName: (meta && meta.displayName) || cohort.displayName,
        mailNickname: (meta && meta.mailNickname) || cohort.mailNickname,
        teamId: String(teamId || '').trim(),
        teamWebUrl: String(webUrl || '').trim()
    });
}

function groupHasTeam(g) {
    const opts = (g && g.resourceProvisioningOptions) || [];
    return Array.isArray(opts) && opts.some((x) => String(x).toLowerCase() === 'team');
}

/**
 * M365-Gruppen / Teams nach Anzeigename suchen (für Verknüpfung).
 * @param {string} query
 * @returns {Promise<Array<{ id: string, displayName: string, mailNickname: string, mail: string, hasTeam: boolean }>>}
 */
export async function searchTeamsGroups(query) {
    const q = String(query || '').trim();
    if (q.length < 2) return [];
    const api = G();
    if (!api) throw new Error('Graph-Modul nicht geladen.');
    const tok = await token([
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.Read.All',
        'https://graph.microsoft.com/Directory.Read.All'
    ]);
    const safe = q.replace(/'/g, "''");
    const select = 'id,displayName,mailNickname,mail,resourceProvisioningOptions,groupTypes';
    /** @type {object[]} */
    let rows = [];
    try {
        const data = await api.graphJson(
            'GET',
            "/groups?$filter=startswith(displayName,'" +
                safe +
                "')&$select=" +
                select +
                '&$top=25',
            tok,
            undefined,
            { ConsistencyLevel: 'eventual' }
        );
        rows = (data && data.value) || [];
    } catch {
        const data = await api.graphJson(
            'GET',
            "/groups?$filter=startswith(displayName,'" + safe + "')&$select=" + select + '&$top=25',
            tok,
            undefined
        );
        rows = (data && data.value) || [];
    }
    return rows
        .filter((g) => g && g.id)
        .filter((g) => {
            const types = Array.isArray(g.groupTypes) ? g.groupTypes : [];
            return types.indexOf('Unified') !== -1 || groupHasTeam(g);
        })
        .map((g) => ({
            id: String(g.id),
            displayName: String(g.displayName || ''),
            mailNickname: String(g.mailNickname || ''),
            mail: String(g.mail || ''),
            hasTeam: groupHasTeam(g)
        }));
}

/**
 * Bestehende M365-Gruppe verknüpfen und ggf. Teams-Funktionalität nachziehen.
 * @param {string} cohortId
 * @param {string} groupId
 * @param {{ onLog?: (m: string) => void, displayName?: string, mailNickname?: string }} [opts]
 */
export async function linkCohortTeamEnsured(cohortId, groupId, opts) {
    const log = (m) => {
        if (opts && typeof opts.onLog === 'function') opts.onLog(m);
    };
    const cohort = getCohort(cohortId);
    if (!cohort) throw new Error('Jahrgang nicht gefunden.');
    const gid = String(groupId || '').trim();
    if (!gid) throw new Error('Gruppen-ID fehlt.');

    const api = G();
    if (!api) throw new Error('Graph-Modul nicht geladen.');
    const tok = await token(TEAM_SCOPES);

    log('Lade Gruppe …');
    let g = null;
    try {
        g = await api.graphJson(
            'GET',
            '/groups/' +
                encodeURIComponent(gid) +
                '?$select=id,displayName,mailNickname,mail,resourceProvisioningOptions,groupTypes',
            tok,
            undefined
        );
    } catch (e) {
        throw new Error(
            'Gruppe nicht gefunden: ' + (e && e.message ? e.message : String(e || ''))
        );
    }

    const displayName =
        String((opts && opts.displayName) || (g && g.displayName) || cohort.displayName || '').trim();
    const mailNickname = String(
        (opts && opts.mailNickname) || (g && g.mailNickname) || cohort.mailNickname || ''
    ).trim();

    let hasTeam = groupHasTeam(g);
    if (!hasTeam) {
        log('Gruppe ohne Teams-Funktionalität – provisioniere Team …');
        await api.provisionTeamForGroup(tok, gid);
        hasTeam = true;
        log('Team bereit.');
    } else {
        log('Teams-Funktionalität bereits vorhanden.');
    }

    let webUrl = '';
    try {
        const team = await api.graphJson(
            'GET',
            '/teams/' + encodeURIComponent(gid) + '?$select=id,webUrl',
            tok,
            undefined
        );
        webUrl = team && team.webUrl ? String(team.webUrl) : '';
    } catch {
        webUrl = '';
    }

    const saved = linkCohortTeam(cohortId, gid, webUrl, { displayName, mailNickname });
    return { ok: true, cohort: saved.item, hasTeam };
}
