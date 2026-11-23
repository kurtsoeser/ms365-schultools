/**
 * Graph: Teams-Gruppenchats oder private Kanäle im Lehrerteam anlegen / syncen.
 * Transport: shared/graph-client.js (R22 / b11).
 */
import { graphJson as sharedGraphJson } from '../../shared/graph-client.js';

import {
    displayNameForTarget,
    kvEmailFromPlan,
    normalizeTargetKind,
    sanitizeTeamsChannelDisplayName,
    sanitizeTeamsChatTopic
} from './klassenchats-logic.js';

/**
 * @param {object} G window.ms365GraphUnifiedGroups
 * @param {string} token
 * @param {string} topic
 * @param {string[]} userIds – Graph User-IDs (inkl. angemeldeter User)
 * @param {{ preferredFirstUserId?: string }} [opts] – KV zuerst (nur Reihenfolge; Graph: alle = owner)
 */
export async function createGroupChat(G, token, topic, userIds, opts) {
    const ids = Array.from(new Set((userIds || []).filter(Boolean)));
    if (ids.length < 2) throw new Error('Mindestens 2 Benutzer für einen Gruppenchat nötig.');
    const preferred = String((opts && opts.preferredFirstUserId) || '').trim();
    const ordered =
        preferred && ids.indexOf(preferred) !== -1
            ? [preferred].concat(ids.filter((id) => id !== preferred))
            : ids;
    const safeTopic = sanitizeTeamsChatTopic(topic);
    // Graph erlaubt bei Gruppenchats nur roles owner|guest – kein „Mitglied“.
    const members = ordered.map((id) => ({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': 'https://graph.microsoft.com/v1.0/users(\'' + id + '\')'
    }));
    const body = {
        chatType: 'group',
        topic: safeTopic || undefined,
        members
    };
    return sharedGraphJson('POST', '/chats', token, body);
}

export async function patchChatTopic(G, token, chatId, topic) {
    const id = String(chatId || '').trim();
    if (!id) throw new Error('chatId fehlt.');
    return sharedGraphJson('PATCH', '/chats/' + encodeURIComponent(id), token, {
        topic: sanitizeTeamsChatTopic(topic)
    });
}

export async function listChatMembers(G, token, chatId) {
    const id = String(chatId || '').trim();
    if (!id) throw new Error('chatId fehlt.');
    const out = [];
    // Kein $select mit userId/email: liegen nur auf aadUserConversationMember, nicht auf conversationMember
    let next = '/chats/' + encodeURIComponent(id) + '/members';
    while (next) {
        const data = await sharedGraphJson('GET', next, token, undefined);
        (data && data.value ? data.value : []).forEach((m) => out.push(m));
        const link = data && data['@odata.nextLink'] ? String(data['@odata.nextLink']) : '';
        if (link.indexOf('https://graph.microsoft.com/v1.0') === 0) {
            next = link.slice('https://graph.microsoft.com/v1.0'.length);
        } else if (link.indexOf('/chats/') === 0 || link.indexOf('chats/') === 0) {
            next = link.startsWith('/') ? link : '/' + link;
        } else next = '';
    }
    return out;
}

export async function addChatMember(G, token, chatId, userId) {
    const cid = String(chatId || '').trim();
    const uid = String(userId || '').trim();
    if (!cid || !uid) throw new Error('chatId/userId fehlt.');
    return sharedGraphJson('POST', '/chats/' + encodeURIComponent(cid) + '/members', token, {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': 'https://graph.microsoft.com/v1.0/users(\'' + uid + '\')'
    });
}

/**
 * Entfernt einen User aus dem Gruppenchat (z. B. IT-Admin nach dem Anlegen).
 * @returns {Promise<boolean>}
 */
export async function removeChatMemberByUserId(G, token, chatId, userId) {
    const cid = String(chatId || '').trim();
    const uid = String(userId || '').trim();
    if (!cid || !uid) return false;
    let members = [];
    try {
        members = await listChatMembers(G, token, cid);
    } catch {
        return false;
    }
    const hit = members.find((m) => String(m.userId || '') === uid);
    if (!hit || !hit.id) return false;
    try {
        await sharedGraphJson(
            'DELETE',
            '/chats/' + encodeURIComponent(cid) + '/members/' + encodeURIComponent(String(hit.id)),
            token,
            undefined
        );
        return true;
    } catch (e) {
        if (isBenignMemberError(e)) return true;
        throw e;
    }
}

function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

function errorText(e) {
    return e && e.message ? String(e.message) : String(e || '');
}

/**
 * „Bereits Mitglied“ / Duplikat – kein echter Fehler beim Hinzufügen.
 * (Kein blankes „400“, sonst werden echte Ablehnungen verschluckt.)
 */
export function isBenignMemberError(e) {
    const msg = errorText(e);
    // Nur echte Duplikate – nicht „does not exist“ o. Ä.
    return (
        /already exist|already a member|added object references already|one or more added object references|conflict|duplicate|409\b/i.test(
            msg
        ) ||
        (/\b400\b/.test(msg) && /already exist|already a member|object references already/i.test(msg))
    );
}

/**
 * Teams hat die Gruppenmitgliedschaft noch nicht im Team-Roster
 * (typisch direkt nach POST /groups/.../members).
 */
export function isTeamMembershipPendingError(e) {
    const msg = errorText(e);
    return /parent roster|not part of the parent|users?\s+.+are not part|member of the (parent )?team|team member|not a member|must be.*member/i.test(
        msg
    );
}

/**
 * Kurze deutsche Meldung statt rohem Graph-JSON.
 */
export function formatKlassenchatsGraphError(e) {
    const msg = errorText(e);
    const rosterMatch = msg.match(
        /users?\s+((?:8:orgid:)?[0-9a-f-]{36}(?:,\s*(?:8:orgid:)?[0-9a-f-]{36})*)\s+are not part of the parent roster/i
    );
    if (rosterMatch || /parent roster|not part of the parent/i.test(msg)) {
        const ids = rosterMatch
            ? rosterMatch[1]
                  .split(/,\s*/)
                  .map((s) => s.replace(/^8:orgid:/i, '').trim())
                  .filter(Boolean)
            : [];
        const who = ids.length ? ' (User-ID: ' + ids.join(', ') + ')' : '';
        return (
            'Mindestens eine Lehrkraft ist (noch) nicht Mitglied des Lehrerteams' +
            who +
            '. Private Kanäle dürfen nur Team-Mitglieder enthalten. Bitte erneut versuchen – fehlende Personen werden zuerst ins Team aufgenommen (Teams braucht oft ein paar Sekunden).'
        );
    }
    if (/Forbidden|AccessDenied/i.test(msg) && /Channel|channel|CreateChannel/i.test(msg)) {
        return (
            'Kein Recht, den privaten Kanal anzulegen (Channel.Create / Team-Berechtigung). Details: ' +
            msg.slice(0, 280)
        );
    }
    // Rohes JSON kürzen, wenn vorhanden
    if (msg.length > 420 && msg.indexOf('{"code"') !== -1) {
        try {
            const j = JSON.parse(msg.replace(/^[^:{]*POST[^:]+:\s*/i, '').trim());
            const detail =
                (j.details && j.details[0] && j.details[0].message) ||
                (j.innerError && j.innerError.message) ||
                j.message ||
                msg;
            return String(detail).slice(0, 420);
        } catch {
            return msg.slice(0, 420);
        }
    }
    return msg;
}

/**
 * Löst E-Mails zu User-IDs auf; fehlende landen in missingEmails.
 */
export async function resolveMemberIds(G, token, emails) {
    const unique = Array.from(
        new Set((emails || []).map((e) => String(e || '').trim().toLowerCase()).filter(Boolean))
    );
    /** @type {Map<string, string>} */
    const idByEmail = new Map();
    const missingEmails = [];
    if (typeof G.resolveUsersByEmailsBulk === 'function') {
        const res = await G.resolveUsersByEmailsBulk(token, unique);
        const map = res && res.byEmail instanceof Map ? res.byEmail : new Map();
        unique.forEach((em) => {
            const u = map.get(em);
            if (u && u.id) idByEmail.set(em, String(u.id));
            else missingEmails.push(em);
        });
    } else {
        for (let i = 0; i < unique.length; i++) {
            const em = unique[i];
            try {
                const u = await G.resolveUserByEmail(token, em);
                if (u && u.id) idByEmail.set(em, String(u.id));
                else missingEmails.push(em);
            } catch {
                missingEmails.push(em);
            }
        }
    }
    return { idByEmail, missingEmails };
}

/**
 * Stellt sicher, dass User im Lehrerteam (M365-Gruppe) sind – Voraussetzung für private Kanäle.
 * @returns {Promise<{ added: number, addedUserIds: string[] }>}
 */
export async function ensureTeamMembers(G, token, teamId, userIds) {
    const tid = String(teamId || '').trim();
    if (!tid) throw new Error('teamId fehlt.');
    const ids = Array.from(new Set((userIds || []).filter(Boolean)));
    let added = 0;
    /** @type {string[]} */
    const addedUserIds = [];
    for (let i = 0; i < ids.length; i++) {
        const uid = ids[i];
        try {
            if (typeof G.graphAddMember === 'function') {
                await G.graphAddMember(token, tid, uid);
            } else {
                await sharedGraphJson(
                    'POST',
                    '/groups/' + encodeURIComponent(tid) + '/members/$ref',
                    token,
                    {
                        '@odata.id':
                            'https://graph.microsoft.com/v1.0/directoryObjects/' + encodeURIComponent(uid)
                    }
                );
            }
            added += 1;
            addedUserIds.push(uid);
        } catch (e) {
            if (!isBenignMemberError(e)) throw e;
        }
    }
    return { added, addedUserIds };
}

/**
 * Nach Gruppen-Add braucht Teams oft Zeit, bis der User im Team-Roster steht.
 * @param {number} [addedCount]
 */
export async function settleTeamRosterAfterAdds(addedCount) {
    const n = Number(addedCount) || 0;
    if (n <= 0) return;
    // Erste Sync oft 2–8s; etwas Puffer für Batch-Läufe
    await sleep(Math.min(12000, 3500 + n * 400));
}

/**
 * Alle geplanten E-Mails einmal ins Lehrerteam holen (vor Batch-Kanalanlage).
 * @returns {Promise<{ added: number, userIds: string[], missingEmails: string[] }>}
 */
export async function prepareTeacherTeamForChannels(G, token, teamId, emails, meUserId) {
    const { idByEmail, missingEmails } = await resolveMemberIds(G, token, emails);
    const userIds = Array.from(new Set(Array.from(idByEmail.values())));
    if (meUserId && userIds.indexOf(meUserId) === -1) userIds.push(meUserId);
    const teamEnsure = await ensureTeamMembers(G, token, teamId, userIds);
    await settleTeamRosterAfterAdds(teamEnsure.added);
    return { added: teamEnsure.added, userIds, missingEmails };
}

function channelMemberBody(userId, asOwner) {
    return {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: asOwner ? ['owner'] : [],
        'user@odata.bind': 'https://graph.microsoft.com/v1.0/users(\'' + userId + '\')'
    };
}

function memberHasOwnerRole(m) {
    return Array.isArray(m && m.roles) && m.roles.indexOf('owner') !== -1;
}

/**
 * Payload für privaten Kanal: explizite Owner + Mitglieder (kein „alle Owner“).
 * @param {string[]} ownerUserIds
 * @param {string[]} memberUserIds
 */
export function buildPrivateChannelMembersPayload(ownerUserIds, memberUserIds) {
    const owners = Array.from(new Set((ownerUserIds || []).filter(Boolean)));
    const membersOnly = Array.from(new Set((memberUserIds || []).filter(Boolean)));
    /** @type {ReturnType<typeof channelMemberBody>[]} */
    const members = [];
    const seen = new Set();
    owners.forEach((id) => {
        if (seen.has(id)) return;
        seen.add(id);
        members.push(channelMemberBody(id, true));
    });
    membersOnly.forEach((id) => {
        if (seen.has(id)) return;
        seen.add(id);
        members.push(channelMemberBody(id, false));
    });
    if (!members.length) throw new Error('Mindestens 1 Mitglied für einen privaten Kanal nötig.');
    if (!members.some(memberHasOwnerRole)) {
        members[0].roles = ['owner'];
    }
    return members.slice(0, 200);
}

/**
 * Gewünschte Owner-IDs: Klassenvorstand; Fallback erste Lehrkraft; zum Anlegen zusätzlich der angemeldete User.
 * @param {Map<string, string>} idByEmail
 * @param {object} plan
 * @param {string} meUserId
 * @param {{ forCreate?: boolean }} [opts]
 */
export function resolveChannelOwnerIds(idByEmail, plan, meUserId, opts) {
    const forCreate = !!(opts && opts.forCreate);
    const map = idByEmail instanceof Map ? idByEmail : new Map();
    const kvEmail = kvEmailFromPlan(plan);
    const kvId = kvEmail && map.get(kvEmail) ? String(map.get(kvEmail)) : '';
    const classIds = Array.from(new Set(Array.from(map.values()).filter(Boolean)));
    /** @type {string[]} */
    const owners = [];
    if (kvId) owners.push(kvId);
    else if (classIds[0]) owners.push(classIds[0]);
    const me = String(meUserId || '').trim();
    // Graph: Anlegender muss beim Create als Owner in members stehen
    if (forCreate && me && owners.indexOf(me) === -1) owners.unshift(me);
    return Array.from(new Set(owners.filter(Boolean)));
}

/**
 * @param {object} G
 * @param {string} token
 * @param {string} teamId
 * @param {string} displayName
 * @param {{ ownerUserIds: string[], memberUserIds: string[] }} membership
 * @param {{ description?: string, onMembershipPending?: (attempt: number) => Promise<void> }} [opts]
 */
export async function createPrivateChannel(G, token, teamId, displayName, membership, opts) {
    const tid = String(teamId || '').trim();
    const name = sanitizeTeamsChannelDisplayName(displayName);
    if (!tid) throw new Error('teamId fehlt.');
    if (!name) throw new Error('Kanalname fehlt.');

    const members = buildPrivateChannelMembersPayload(
        membership && membership.ownerUserIds,
        membership && membership.memberUserIds
    );

    const body = {
        '@odata.type': '#microsoft.graph.channel',
        membershipType: 'private',
        displayName: name,
        description: String((opts && opts.description) || '').trim() || undefined,
        members
    };

    const maxAttempts = 5;
    let lastErr = null;
    for (let attempt = 0; attempt < maxAttempts; attempt++) {
        try {
            return await sharedGraphJson(
                'POST',
                '/teams/' + encodeURIComponent(tid) + '/channels',
                token,
                body
            );
        } catch (e) {
            lastErr = e;
            if (attempt < maxAttempts - 1 && isTeamMembershipPendingError(e)) {
                if (opts && typeof opts.onMembershipPending === 'function') {
                    try {
                        await opts.onMembershipPending(attempt);
                    } catch {
                        /* ignore – Retry trotzdem */
                    }
                }
                await sleep(3000 + attempt * 2500);
                continue;
            }
            const friendly = new Error(formatKlassenchatsGraphError(e));
            friendly.cause = e;
            throw friendly;
        }
    }
    const friendly = new Error(formatKlassenchatsGraphError(lastErr));
    friendly.cause = lastErr;
    throw friendly;
}

export async function patchChannelMemberRoles(G, token, teamId, channelId, membershipId, asOwner) {
    const tid = String(teamId || '').trim();
    const cid = String(channelId || '').trim();
    const mid = String(membershipId || '').trim();
    if (!tid || !cid || !mid) throw new Error('teamId/channelId/membershipId fehlt.');
    return sharedGraphJson(
        'PATCH',
        '/teams/' +
            encodeURIComponent(tid) +
            '/channels/' +
            encodeURIComponent(cid) +
            '/members/' +
            encodeURIComponent(mid),
        token,
        {
            '@odata.type': '#microsoft.graph.aadUserConversationMember',
            roles: asOwner ? ['owner'] : []
        }
    );
}

export async function removeChannelMemberByUserId(G, token, teamId, channelId, userId) {
    const tid = String(teamId || '').trim();
    const cid = String(channelId || '').trim();
    const uid = String(userId || '').trim();
    if (!tid || !cid || !uid) return false;
    let members = [];
    try {
        members = await listChannelMembers(G, token, tid, cid);
    } catch {
        return false;
    }
    const hit = members.find((m) => String(m.userId || '') === uid);
    if (!hit || !hit.id) return false;
    try {
        await sharedGraphJson(
            'DELETE',
            '/teams/' +
                encodeURIComponent(tid) +
                '/channels/' +
                encodeURIComponent(cid) +
                '/members/' +
                encodeURIComponent(String(hit.id)),
            token,
            undefined
        );
        return true;
    } catch (e) {
        if (isBenignMemberError(e)) return true;
        throw e;
    }
}

/**
 * Stellt sicher: genau die gewünschten User (KV) sind Channel-Owner,
 * alle anderen Klassenmitglieder nur Mitglieder; IT-Admin ggf. entfernen.
 * @returns {Promise<{ ownersAligned: number, demoted: number, removed: number, verifiedOwnerIds: string[] }>}
 */
export async function alignPrivateChannelOwnership(G, token, teamId, channelId, desiredOwnerIds, classUserIds, meUserId) {
    const tid = String(teamId || '').trim();
    const cid = String(channelId || '').trim();
    const ownersWanted = Array.from(new Set((desiredOwnerIds || []).filter(Boolean)));
    const ownersWantedSet = new Set(ownersWanted);
    const classSet = new Set((classUserIds || []).filter(Boolean));
    const me = String(meUserId || '').trim();
    let ownersAligned = 0;
    let demoted = 0;
    let removed = 0;

    if (!ownersWanted.length) {
        throw new Error(
            'Kein Kanal-Owner ermittelbar (Klassenvorstand fehlt in den Stammdaten bzw. E-Mail nicht auflösbar).'
        );
    }

    async function loadByUser() {
        const members = await listChannelMembers(G, token, tid, cid);
        /** @type {Map<string, object>} */
        const byUser = new Map();
        members.forEach((m) => {
            const uid = String(m.userId || '');
            if (uid) byUser.set(uid, m);
        });
        return { members, byUser };
    }

    // Mehrere Versuche: Teams liefert Mitglieder/Rollen nach Create oft verzögert
    let lastVerifyErr = null;
    attemptLoop: for (let attempt = 0; attempt < 5; attempt++) {
        if (attempt > 0) await sleep(1500 + attempt * 1000);

        let { members, byUser } = await loadByUser();

        // 1) Gewünschte Owner: fehlende hinzufügen, sonst hochstufen
        for (let oi = 0; oi < ownersWanted.length; oi++) {
            const uid = ownersWanted[oi];
            let m = byUser.get(uid);
            if (!m) {
                try {
                    await addChannelMember(G, token, tid, cid, uid, true);
                    ownersAligned += 1;
                } catch (e) {
                    if (isBenignMemberError(e)) {
                        /* bereits drin */
                    } else if (attempt < 4 && isTeamMembershipPendingError(e)) {
                        lastVerifyErr = e instanceof Error ? e : new Error(String(e));
                        continue attemptLoop;
                    } else {
                        throw e;
                    }
                }
                ({ members, byUser } = await loadByUser());
                m = byUser.get(uid);
            }
            if (m && !memberHasOwnerRole(m)) {
                try {
                    await patchChannelMemberRoles(G, token, tid, cid, String(m.id), true);
                    ownersAligned += 1;
                    m.roles = ['owner'];
                } catch (e) {
                    if (!isBenignMemberError(e)) throw e;
                }
            }
        }

        ({ members, byUser } = await loadByUser());

        // Ohne mindestens einen gewünschten Owner nicht herabstufen (sonst hängt der falsche „last owner“)
        const desiredOwnerPresent = ownersWanted.some((uid) => {
            const m = byUser.get(uid);
            return !!(m && memberHasOwnerRole(m));
        });
        if (!desiredOwnerPresent) {
            lastVerifyErr = new Error(
                'Klassenvorstand konnte noch nicht als Kanal-Owner gesetzt werden (Versuch ' +
                    (attempt + 1) +
                    '/5).'
            );
            continue;
        }

        // 2) Alle anderen Owner → Mitglied
        for (let i = 0; i < members.length; i++) {
            const m = members[i];
            const uid = String(m.userId || '');
            if (!uid || !memberHasOwnerRole(m)) continue;
            if (ownersWantedSet.has(uid)) continue;
            const ownerCount = members.filter(memberHasOwnerRole).length;
            if (ownerCount <= 1) {
                lastVerifyErr = new Error(
                    'Kanal hat noch einen fremden alleinigen Owner – KV-Owner fehlt oder ist noch nicht wirksam.'
                );
                continue attemptLoop;
            }
            try {
                await patchChannelMemberRoles(G, token, tid, cid, String(m.id), false);
                demoted += 1;
                m.roles = [];
            } catch (e) {
                if (!isBenignMemberError(e)) throw e;
            }
        }

        // 3) IT-Admin entfernen, wenn nicht Klassenmitglied / nicht gewünschter Owner
        if (me && !classSet.has(me) && !ownersWantedSet.has(me)) {
            const ok = await removeChannelMemberByUserId(G, token, tid, cid, me);
            if (ok) removed += 1;
        }

        // 4) Verifizieren: genau die gewünschten Owner
        ({ members, byUser } = await loadByUser());
        const verifiedOwnerIds = members
            .filter(memberHasOwnerRole)
            .map((m) => String(m.userId || ''))
            .filter(Boolean);
        const missingDesired = ownersWanted.filter((uid) => verifiedOwnerIds.indexOf(uid) === -1);
        const unexpected = verifiedOwnerIds.filter((uid) => !ownersWantedSet.has(uid));

        if (!missingDesired.length && !unexpected.length) {
            return {
                ownersAligned,
                demoted,
                removed,
                verifiedOwnerIds
            };
        }

        lastVerifyErr = new Error(
            missingDesired.length
                ? 'Klassenvorstand ist noch nicht Kanal-Owner.'
                : 'Noch andere Kanal-Owner außer dem Klassenvorstand vorhanden.'
        );
    }

    throw lastVerifyErr || new Error('Kanal-Owner (Klassenvorstand) konnte nicht gesetzt werden.');
}

export async function patchChannelDisplayName(G, token, teamId, channelId, displayName) {
    const tid = String(teamId || '').trim();
    const cid = String(channelId || '').trim();
    const name = sanitizeTeamsChannelDisplayName(displayName);
    if (!tid || !cid) throw new Error('teamId/channelId fehlt.');
    if (!name) throw new Error('Kanalname fehlt.');
    return sharedGraphJson(
        'PATCH',
        '/teams/' + encodeURIComponent(tid) + '/channels/' + encodeURIComponent(cid),
        token,
        { displayName: name }
    );
}

export async function listChannelMembers(G, token, teamId, channelId) {
    const tid = String(teamId || '').trim();
    const cid = String(channelId || '').trim();
    if (!tid || !cid) throw new Error('teamId/channelId fehlt.');
    const out = [];
    // Kein $select=userId/email: Graph prüft $select gegen Basistyp conversationMember → BadRequest.
    // Ohne $select liefert aadUserConversationMember inkl. userId/roles/displayName.
    let next =
        '/teams/' + encodeURIComponent(tid) + '/channels/' + encodeURIComponent(cid) + '/members';
    while (next) {
        const data = await sharedGraphJson('GET', next, token, undefined);
        (data && data.value ? data.value : []).forEach((m) => out.push(m));
        const link = data && data['@odata.nextLink'] ? String(data['@odata.nextLink']) : '';
        if (link.indexOf('https://graph.microsoft.com/v1.0') === 0) {
            next = link.slice('https://graph.microsoft.com/v1.0'.length);
        } else if (link.indexOf('/teams/') === 0 || link.indexOf('teams/') === 0) {
            next = link.startsWith('/') ? link : '/' + link;
        } else next = '';
    }
    return out;
}

export async function addChannelMember(G, token, teamId, channelId, userId, asOwner) {
    const tid = String(teamId || '').trim();
    const cid = String(channelId || '').trim();
    const uid = String(userId || '').trim();
    if (!tid || !cid || !uid) throw new Error('teamId/channelId/userId fehlt.');
    return sharedGraphJson(
        'POST',
        '/teams/' +
            encodeURIComponent(tid) +
            '/channels/' +
            encodeURIComponent(cid) +
            '/members',
        token,
        channelMemberBody(uid, !!asOwner)
    );
}

/**
 * Erstellt oder syncronisiert einen Klassen-Gruppenchat.
 * Hinweis: Microsoft Graph erlaubt bei Gruppenchats nur roles owner|guest – alle Lehrkräfte
 * sind technisch Owner; KV wird zuerst gesetzt. Der IT-Admin wird nach dem Anlegen entfernt,
 * sofern er nicht selbst Klassenmitglied ist.
 * @returns {{ kind: 'groupChat', chatId: string, topic: string, webUrl: string, created: boolean, added: number, topicUpdated: boolean, adminRemoved: boolean, missingEmails: string[], memberEmails: string[] }}
 */
export async function provisionClassChat(G, token, plan, meUserId) {
    const emails = Array.isArray(plan.memberEmails) ? plan.memberEmails.slice() : [];
    const { idByEmail, missingEmails } = await resolveMemberIds(G, token, emails);
    const classUserIds = Array.from(new Set(Array.from(idByEmail.values())));
    const me = String(meUserId || '').trim();
    const kvEmail = kvEmailFromPlan(plan);
    const kvUserId = kvEmail && idByEmail.get(kvEmail) ? String(idByEmail.get(kvEmail)) : '';

    // Zum Anlegen muss der angemeldete User in members stehen
    const createIds = classUserIds.slice();
    if (me && createIds.indexOf(me) === -1) createIds.push(me);

    if (classUserIds.length < 2 && createIds.length < 2) {
        const err = new Error(
            'Zu wenige auflösbare Benutzer (mind. 2). Fehlend: ' + (missingEmails.join(', ') || '–')
        );
        err.missingEmails = missingEmails;
        throw err;
    }
    if (classUserIds.length < 2) {
        const err = new Error(
            'Zu wenige Klassen-Lehrkräfte (mind. 2). Fehlend: ' + (missingEmails.join(', ') || '–')
        );
        err.missingEmails = missingEmails;
        throw err;
    }

    let chatId = String(plan.chatId || '').trim();
    let created = false;
    let topicUpdated = false;
    let added = 0;
    let adminRemoved = false;
    let webUrl = '';
    const topic = displayNameForTarget(plan.topic, 'groupChat');

    if (!chatId) {
        const chat = await createGroupChat(G, token, topic, createIds, {
            preferredFirstUserId: kvUserId || undefined
        });
        chatId = chat && chat.id ? String(chat.id) : '';
        webUrl = chat && chat.webUrl ? String(chat.webUrl) : '';
        created = true;
        if (!chatId) throw new Error('Chat angelegt, aber keine ID zurückgegeben.');
        if (me && classUserIds.indexOf(me) === -1) {
            adminRemoved = await removeChatMemberByUserId(G, token, chatId, me);
        }
    } else {
        if (topic && sanitizeTeamsChatTopic(plan.existingTopic) !== topic) {
            await patchChatTopic(G, token, chatId, topic);
            topicUpdated = true;
        }
        let existingIds = new Set();
        try {
            const members = await listChatMembers(G, token, chatId);
            members.forEach((m) => {
                const uid = m.userId || m.id;
                if (uid) existingIds.add(String(uid));
            });
        } catch {
            existingIds = new Set();
        }
        for (let i = 0; i < classUserIds.length; i++) {
            const uid = classUserIds[i];
            if (existingIds.has(uid)) continue;
            try {
                await addChatMember(G, token, chatId, uid);
                added += 1;
                existingIds.add(uid);
            } catch (e) {
                if (!isBenignMemberError(e)) throw e;
            }
        }
        if (me && classUserIds.indexOf(me) === -1 && existingIds.has(me)) {
            adminRemoved = await removeChatMemberByUserId(G, token, chatId, me);
        }
        try {
            const chat = await sharedGraphJson(
                'GET',
                '/chats/' + encodeURIComponent(chatId) + '?$select=id,topic,webUrl',
                token,
                undefined
            );
            if (chat && chat.webUrl) webUrl = String(chat.webUrl);
        } catch {
            /* ignore */
        }
    }

    return {
        kind: 'groupChat',
        chatId,
        channelId: '',
        teamId: '',
        topic,
        webUrl,
        created,
        added,
        topicUpdated,
        adminRemoved,
        missingEmails,
        memberEmails: Array.from(idByEmail.keys())
    };
}

/**
 * Erstellt oder syncronisiert einen privaten Klassenkanal im Lehrerteam.
 * Klassenvorstand = Kanal-Owner, übrige Lehrkräfte = Mitglieder.
 * Der angemeldete IT-Admin muss beim Create kurz Owner sein und wird danach entfernt
 * (außer er ist selbst KV/Klassenmitglied ohne Owner-Rolle → dann Mitglied).
 * @returns {{ kind: 'privateChannel', chatId: string, channelId: string, teamId: string, topic: string, webUrl: string, created: boolean, added: number, topicUpdated: boolean, teamMembersAdded: number, ownersAligned: number, demoted: number, adminRemoved: number, missingEmails: string[], memberEmails: string[] }}
 */
export async function provisionClassPrivateChannel(G, token, plan, meUserId) {
    const teamId = String(plan.teamId || '').trim();
    if (!teamId) throw new Error('Lehrerteam (teamId) fehlt.');

    const emails = Array.isArray(plan.memberEmails) ? plan.memberEmails.slice() : [];
    const { idByEmail, missingEmails } = await resolveMemberIds(G, token, emails);
    const classUserIds = Array.from(new Set(Array.from(idByEmail.values())));
    const me = String(meUserId || '').trim();
    const kvEmail = kvEmailFromPlan(plan);
    const kvUserId = kvEmail && idByEmail.get(kvEmail) ? String(idByEmail.get(kvEmail)) : '';

    if (classUserIds.length < 2) {
        const err = new Error(
            'Zu wenige auflösbare Benutzer (mind. 2). Fehlend: ' + (missingEmails.join(', ') || '–')
        );
        err.missingEmails = missingEmails;
        throw err;
    }
    if (!kvEmail) {
        throw new Error(
            'Kein Klassenvorstand für Klasse „' +
                String(plan.klasse || '').trim() +
                '“ in den Stammdaten (headEmail). Ohne KV kann kein Kanal-Owner gesetzt werden.'
        );
    }
    if (!kvUserId) {
        throw new Error(
            'Klassenvorstand „' +
                kvEmail +
                '“ konnte in Microsoft 365 nicht aufgelöst werden – Kanal-Owner nicht setzbar.'
        );
    }

    const ensureIds = classUserIds.slice();
    if (me && ensureIds.indexOf(me) === -1) ensureIds.push(me);

    let teamEnsure = await ensureTeamMembers(G, token, teamId, ensureIds);
    if (!(plan && plan._teamRosterSettled) && teamEnsure.added > 0) {
        await settleTeamRosterAfterAdds(teamEnsure.added);
    }

    const ownerIdsForCreate = resolveChannelOwnerIds(idByEmail, plan, me, { forCreate: true });
    const ownerIdsFinal = resolveChannelOwnerIds(idByEmail, plan, me, { forCreate: false });
    // Sicherheit: Final-Owner muss der KV sein
    if (ownerIdsFinal.indexOf(kvUserId) === -1) {
        ownerIdsFinal.length = 0;
        ownerIdsFinal.push(kvUserId);
    }
    const memberIdsForCreate = classUserIds.filter((id) => ownerIdsForCreate.indexOf(id) === -1);

    let channelId = String(plan.channelId || '').trim();
    let created = false;
    let topicUpdated = false;
    let added = 0;
    let webUrl = '';
    let ownersAligned = 0;
    let demoted = 0;
    let adminRemoved = 0;
    const topic = displayNameForTarget(plan.topic, 'privateChannel');

    if (!channelId) {
        const ch = await createPrivateChannel(
            G,
            token,
            teamId,
            topic,
            { ownerUserIds: ownerIdsForCreate, memberUserIds: memberIdsForCreate },
            {
                description: 'KlassenChat ' + String(plan.klasse || '').trim(),
                onMembershipPending: async () => {
                    const again = await ensureTeamMembers(G, token, teamId, ensureIds);
                    teamEnsure = {
                        added: teamEnsure.added + again.added,
                        addedUserIds: (teamEnsure.addedUserIds || []).concat(again.addedUserIds || [])
                    };
                    await settleTeamRosterAfterAdds(Math.max(1, again.added || 1));
                }
            }
        );
        channelId = ch && ch.id ? String(ch.id) : '';
        webUrl = ch && ch.webUrl ? String(ch.webUrl) : '';
        created = true;
        if (!channelId) throw new Error('Kanal angelegt, aber keine ID zurückgegeben.');
    } else {
        if (
            topic &&
            sanitizeTeamsChannelDisplayName(plan.existingTopic) !== topic
        ) {
            await patchChannelDisplayName(G, token, teamId, channelId, topic);
            topicUpdated = true;
        }
        let existingIds = new Set();
        try {
            const members = await listChannelMembers(G, token, teamId, channelId);
            members.forEach((m) => {
                const uid = m.userId || m.id;
                if (uid) existingIds.add(String(uid));
            });
        } catch {
            existingIds = new Set();
        }
        for (let i = 0; i < classUserIds.length; i++) {
            const uid = classUserIds[i];
            if (existingIds.has(uid)) continue;
            const asOwner = ownerIdsFinal.indexOf(uid) !== -1;
            let ok = false;
            for (let attempt = 0; attempt < 3 && !ok; attempt++) {
                try {
                    await addChannelMember(G, token, teamId, channelId, uid, asOwner);
                    added += 1;
                    existingIds.add(uid);
                    ok = true;
                } catch (e) {
                    if (isBenignMemberError(e)) {
                        ok = true;
                        break;
                    }
                    if (attempt < 2 && isTeamMembershipPendingError(e)) {
                        await sleep(2000 + attempt * 1000);
                        continue;
                    }
                    throw e;
                }
            }
        }
        try {
            const ch = await sharedGraphJson(
                'GET',
                '/teams/' +
                    encodeURIComponent(teamId) +
                    '/channels/' +
                    encodeURIComponent(channelId) +
                    '?$select=id,displayName,webUrl',
                token,
                undefined
            );
            if (ch && ch.webUrl) webUrl = String(ch.webUrl);
        } catch {
            /* ignore */
        }
    }

    // Rollen bereinigen / nachziehen: KV = alleiniger Owner (auch beim erneuten Sync)
    try {
        const align = await alignPrivateChannelOwnership(
            G,
            token,
            teamId,
            channelId,
            [kvUserId],
            classUserIds,
            me
        );
        ownersAligned = align.ownersAligned;
        demoted = align.demoted;
        adminRemoved = align.removed;
    } catch (e) {
        throw new Error(
            formatKlassenchatsGraphError(e) +
                ' (Kanal angelegt/gefunden, aber Klassenvorstand konnte nicht als alleiniger Owner gesetzt werden.)'
        );
    }

    return {
        kind: 'privateChannel',
        chatId: '',
        channelId,
        teamId,
        topic,
        webUrl,
        created,
        added,
        topicUpdated,
        teamMembersAdded: teamEnsure.added,
        ownersAligned,
        demoted,
        adminRemoved,
        kvEmail,
        kvOwnerOk: true,
        missingEmails,
        memberEmails: Array.from(idByEmail.keys())
    };
}

/**
 * Dispatch: Gruppenchat oder privater Kanal je nach plan.kind / targetKind.
 */
export async function provisionClassResource(G, token, plan, meUserId) {
    const kind = normalizeTargetKind(plan && (plan.kind || plan.targetKind));
    if (kind === 'privateChannel') {
        return provisionClassPrivateChannel(G, token, plan, meUserId);
    }
    return provisionClassChat(G, token, plan, meUserId);
}
