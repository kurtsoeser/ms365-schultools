/**
 * Graph: Teams-Gruppenchats für Klassen anlegen / Mitglieder & Topic syncen.
 */

/**
 * @param {object} G window.ms365GraphUnifiedGroups
 * @param {string} token
 * @param {string} topic
 * @param {string[]} userIds – Graph User-IDs (inkl. angemeldeter User)
 */
export async function createGroupChat(G, token, topic, userIds) {
    const ids = Array.from(new Set((userIds || []).filter(Boolean)));
    if (ids.length < 2) throw new Error('Mindestens 2 Benutzer für einen Gruppenchat nötig.');
    const members = ids.map((id) => ({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'kevin.m@example.com': 'https://graph.microsoft.com/v1.0/users(\'' + id + '\')'
    }));
    const body = {
        chatType: 'group',
        topic: String(topic || '').trim() || undefined,
        members
    };
    return G.graphJson('POST', '/chats', token, body);
}

export async function patchChatTopic(G, token, chatId, topic) {
    const id = String(chatId || '').trim();
    if (!id) throw new Error('chatId fehlt.');
    return G.graphJson('PATCH', '/chats/' + encodeURIComponent(id), token, {
        topic: String(topic || '').trim()
    });
}

export async function listChatMembers(G, token, chatId) {
    const id = String(chatId || '').trim();
    if (!id) throw new Error('chatId fehlt.');
    const out = [];
    let next = '/chats/' + encodeURIComponent(id) + '/members?$select=id,displayName,roles,email,userId';
    while (next) {
        const data = await G.graphJson('GET', next, token, undefined);
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
    return G.graphJson('POST', '/chats/' + encodeURIComponent(cid) + '/members', token, {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'kevin.m@example.com': 'https://graph.microsoft.com/v1.0/users(\'' + uid + '\')'
    });
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
 * Erstellt oder syncronisiert einen Klassen-Chat.
 * @returns {{ chatId: string, topic: string, webUrl: string, created: boolean, added: number, topicUpdated: boolean, missingEmails: string[] }}
 */
export async function provisionClassChat(G, token, plan, meUserId) {
    const emails = Array.isArray(plan.memberEmails) ? plan.memberEmails.slice() : [];
    const { idByEmail, missingEmails } = await resolveMemberIds(G, token, emails);
    const userIds = Array.from(new Set(Array.from(idByEmail.values())));
    if (meUserId && userIds.indexOf(meUserId) === -1) userIds.push(meUserId);

    if (userIds.length < 2) {
        const err = new Error(
            'Zu wenige auflösbare Benutzer (mind. 2). Fehlend: ' + (missingEmails.join(', ') || '–')
        );
        err.missingEmails = missingEmails;
        throw err;
    }

    let chatId = String(plan.chatId || '').trim();
    let created = false;
    let topicUpdated = false;
    let added = 0;
    let webUrl = '';
    const topic = String(plan.topic || '').trim();

    if (!chatId) {
        const chat = await createGroupChat(G, token, topic, userIds);
        chatId = chat && chat.id ? String(chat.id) : '';
        webUrl = chat && chat.webUrl ? String(chat.webUrl) : '';
        created = true;
        if (!chatId) throw new Error('Chat angelegt, aber keine ID zurückgegeben.');
    } else {
        // Topic aktualisieren wenn nötig
        if (topic && plan.existingTopic !== topic) {
            await patchChatTopic(G, token, chatId, topic);
            topicUpdated = true;
        }
        // Fehlende Mitglieder nachziehen
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
        for (let i = 0; i < userIds.length; i++) {
            const uid = userIds[i];
            if (existingIds.has(uid)) continue;
            try {
                await addChatMember(G, token, chatId, uid);
                added += 1;
                existingIds.add(uid);
            } catch (e) {
                const msg = e && e.message ? String(e.message) : String(e);
                // bereits Mitglied / Konflikt → ignorieren
                if (!/already|conflict|exist|409|400/i.test(msg)) throw e;
            }
        }
        try {
            const chat = await G.graphJson(
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
        chatId,
        topic,
        webUrl,
        created,
        added,
        topicUpdated,
        missingEmails,
        memberEmails: Array.from(idByEmail.keys())
    };
}
