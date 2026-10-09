/**
 * KlassenChats: Belegung → Chat-Pläne, Namens-Pattern, Diff.
 */

/**
 * @typedef {{ type: string, value?: string }} NameToken
 * @typedef {{
 *   email: string,
 *   code: string,
 *   role: 'lehrer'|'kv',
 *   name?: string
 * }} ChatMember
 * @typedef {{
 *   klasse: string,
 *   topic: string,
 *   members: ChatMember[],
 *   memberEmails: string[],
 *   eligible: boolean,
 *   skipReason: string,
 *   chatId: string,
 *   existingTopic: string,
 *   status: 'neu'|'sync'|'ok'|'skip'
 * }} ClassChatPlan
 */

export function defaultChatNamePattern() {
    return [
        { type: 'yearPrefix' },
        { type: 'text', value: ' | ' },
        { type: 'klasse' },
        { type: 'text', value: ' | ' },
        { type: 'text', value: 'Lehrer' }
    ];
}

const FIELD_TOKEN_TYPES = new Set(['yearPrefix', 'klasse']);

export function normalizeChatNamePattern(pattern) {
    const arr = Array.isArray(pattern) ? pattern : [];
    const out = [];
    arr.forEach((p) => {
        if (!p || typeof p !== 'object') return;
        const type = String(p.type || '').trim();
        if (!type) return;
        if (type === 'text') out.push({ type: 'text', value: String(p.value ?? '') });
        else if (FIELD_TOKEN_TYPES.has(type)) out.push({ type });
    });
    return out.length ? out : defaultChatNamePattern();
}

/** Microsoft Graph: Doppelpunkt im group-chat topic ist verboten (max. 250 Zeichen). */
export function sanitizeTeamsChatTopic(topic) {
    let s = String(topic ?? '').trim();
    if (!s) return '';
    s = s.replace(/:/g, '·');
    if (s.length > 250) s = s.slice(0, 250);
    return s;
}

export function chatTokenLabel(t) {
    if (!t) return '';
    if (t.type === 'yearPrefix') return 'Schuljahr';
    if (t.type === 'klasse') return 'Klasse';
    if (t.type === 'text') {
        const v = String(t.value ?? '');
        return v === '' ? '(leer)' : v;
    }
    return String(t.type);
}

export function buildChatTopicFromPattern(pattern, ctx) {
    const parts = [];
    normalizeChatNamePattern(pattern).forEach((p) => {
        if (p.type === 'text') parts.push(String(p.value ?? ''));
        else if (p.type === 'yearPrefix') parts.push(String((ctx && ctx.yearPrefix) || ''));
        else if (p.type === 'klasse') parts.push(String((ctx && ctx.klasse) || ''));
    });
    return sanitizeTeamsChatTopic(parts.join('').trim());
}

export function calcYearPrefix(date) {
    const now = date instanceof Date ? date : new Date();
    const month = now.getMonth() + 1;
    const year = now.getFullYear();
    const sjYear = month >= 9 ? year : year - 1;
    return 'SJ' + String(sjYear).slice(-2);
}

function normStr(v) {
    return String(v ?? '').trim();
}

function normEmail(v) {
    return normStr(v).toLowerCase();
}

function normCode(v) {
    return normStr(v).toUpperCase();
}

function splitKlassenCell(raw) {
    const s = normStr(raw);
    if (!s) return [];
    return s.split(/[,;~]+/).map((c) => c.trim()).filter(Boolean);
}

/** Mehrklassen-Zelle (geteilt oder zusammengezogen wie 1AK1BK). */
export function isCombinedClassCell(raw) {
    const s = normStr(raw);
    if (!s) return false;
    if (/[,;~]/.test(s)) return true;
    return /(?:\d+[A-Za-z]+){2,}/.test(s);
}

/**
 * Index bekannter Klassencodes aus Stammdaten (für Auflösung von Mehrklassen-Zellen).
 * @param {Array<{ code?: string }>} classes
 */
export function buildKnownClassCodeIndex(classes) {
    /** @type {Map<string, string>} */
    const canonByUpper = new Map();
    (Array.isArray(classes) ? classes : []).forEach((c) => {
        const code = normStr(c && c.code);
        if (!code) return;
        const u = code.toUpperCase();
        if (!canonByUpper.has(u)) canonByUpper.set(u, code);
    });
    const codesSortedByLength = Array.from(canonByUpper.values()).sort(
        (a, b) => b.length - a.length || a.localeCompare(b, 'de')
    );
    return { canonByUpper, codesSortedByLength };
}

function greedyParseCombinedClass(rest, codesSortedByLength, canonByUpper) {
    let s = String(rest || '').replace(/\s+/g, '');
    const result = [];
    while (s.length > 0) {
        let match = '';
        codesSortedByLength.forEach((code) => {
            if (
                s.length >= code.length &&
                s.slice(0, code.length).toUpperCase() === code.toUpperCase()
            ) {
                if (code.length > match.length) match = code;
            }
        });
        if (!match) return [];
        result.push(canonByUpper.get(match.toUpperCase()) || match);
        s = s.slice(match.length);
    }
    return result;
}

/**
 * Belegungs-Klasse → ein oder mehrere Stammdaten-Klassencodes.
 * Geteilte / zusammengezogene Gruppen (1AK1BK1CK) werden auf echte Klassen aufgeteilt.
 * @param {string} klasseRaw
 * @param {ReturnType<typeof buildKnownClassCodeIndex>|Array<{ code?: string }>} classIndexOrClasses
 * @returns {string[]}
 */
export function expandKlasseToKnownCodes(klasseRaw, classIndexOrClasses) {
    const classIndex =
        classIndexOrClasses &&
        typeof classIndexOrClasses === 'object' &&
        classIndexOrClasses.canonByUpper instanceof Map
            ? classIndexOrClasses
            : buildKnownClassCodeIndex(classIndexOrClasses);

    const { canonByUpper, codesSortedByLength } = classIndex;
    const single = normStr(klasseRaw);
    if (!single) return [];

    if (!canonByUpper.size) {
        return splitKlassenCell(single).length ? splitKlassenCell(single) : [single];
    }

    const out = [];
    const seen = new Set();
    const add = (code) => {
        const c = normStr(code);
        if (!c) return;
        const u = c.toUpperCase();
        if (seen.has(u)) return;
        seen.add(u);
        out.push(canonByUpper.get(u) || c);
    };

    const parts = splitKlassenCell(single);
    const segments = parts.length ? parts : [single];
    segments.forEach((part) => {
        const p = part.replace(/\s+/g, '');
        if (!p) return;
        const upper = p.toUpperCase();
        if (canonByUpper.has(upper)) {
            add(canonByUpper.get(upper));
            return;
        }
        if (!isCombinedClassCell(p)) {
            add(p);
            return;
        }
        const parsed = greedyParseCombinedClass(p, codesSortedByLength, canonByUpper);
        if (parsed.length > 0) parsed.forEach(add);
        else add(p);
    });
    return out;
}

/**
 * KV-Map: Klassenkürzel → { email, name }
 * @param {Array<{ code?: string, headEmail?: string, headName?: string }>} classes
 */
export function buildKvByKlasse(classes) {
    /** @type {Map<string, { email: string, name: string }>} */
    const map = new Map();
    (Array.isArray(classes) ? classes : []).forEach((c) => {
        const code = normStr(c && c.code);
        if (!code) return;
        const email = normEmail(c.headEmail);
        if (!email || email.indexOf('@') === -1) return;
        map.set(code, { email, name: normStr(c.headName) });
        map.set(code.toUpperCase(), { email, name: normStr(c.headName) });
    });
    return map;
}

/**
 * @param {object|null} belegung – Snapshot aus getUnterrichtsbelegung()
 * @param {Map<string, { email: string, name: string }>} kvByKlasse
 * @param {{ yearPrefix?: string, namePattern?: NameToken[], existingByKlasse?: Map<string, { chatId?: string, topic?: string }>, knownClassCodes?: Array<{ code?: string }> }} [opts]
 * @returns {ClassChatPlan[]}
 */
export function buildClassChatPlans(belegung, kvByKlasse, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    const yearPrefix = normStr(o.yearPrefix) || calcYearPrefix();
    const pattern = normalizeChatNamePattern(o.namePattern);
    const existingByKlasse = o.existingByKlasse instanceof Map ? o.existingByKlasse : new Map();
    const classIndex = buildKnownClassCodeIndex(o.knownClassCodes || []);

    /** @type {Map<string, Map<string, ChatMember>>} */
    const byKlasse = new Map();

    const rows = belegung && Array.isArray(belegung.rows) ? belegung.rows : [];
    rows.forEach((r) => {
        const klasseRaw = normStr(r && r.klasse);
        if (!klasseRaw) return;
        const email = normEmail(r.lehrerEmail);
        if (!email || email.indexOf('@') === -1) return;
        const targets = expandKlasseToKnownCodes(klasseRaw, classIndex);
        targets.forEach((klasse) => {
            if (!klasse) return;
            if (!byKlasse.has(klasse)) byKlasse.set(klasse, new Map());
            const members = byKlasse.get(klasse);
            if (members.has(email)) return;
            members.set(email, {
                email,
                code: normCode(r.lehrerCode),
                role: 'lehrer'
            });
        });
    });

    const kvMap = kvByKlasse instanceof Map ? kvByKlasse : new Map();
    byKlasse.forEach((members, klasse) => {
        const kv =
            kvMap.get(klasse) ||
            kvMap.get(klasse.toUpperCase()) ||
            kvMap.get(normStr(klasse));
        if (!kv || !kv.email) return;
        if (members.has(kv.email)) {
            const cur = members.get(kv.email);
            if (cur.role !== 'kv') {
                members.set(kv.email, {
                    email: kv.email,
                    code: cur.code || '',
                    role: 'kv',
                    name: kv.name || cur.name
                });
            }
            return;
        }
        members.set(kv.email, {
            email: kv.email,
            code: '',
            role: 'kv',
            name: kv.name || ''
        });
    });

    // Klassen nur mit KV, ohne Belegungszeilen: trotzdem aufnehmen, wenn KV existiert
    // (meist <2 → skip). Nur wenn Belegung Klassen liefert oder wir explizit nur Belegung nutzen.
    // Aktuell: nur Klassen aus Belegung – KV wird ergänzt.

    const plans = [];
    let klassen = Array.from(byKlasse.keys());
    if (classIndex.canonByUpper.size) {
        klassen = klassen.filter((k) => classIndex.canonByUpper.has(k.toUpperCase()));
    }
    klassen.sort((a, b) => a.localeCompare(b, 'de'));
    klassen.forEach((klasse) => {
        const members = Array.from(byKlasse.get(klasse).values()).sort((a, b) =>
            a.email.localeCompare(b.email, 'de')
        );
        const memberEmails = members.map((m) => m.email);
        const topic = buildChatTopicFromPattern(pattern, { yearPrefix, klasse });
        const existing = existingByKlasse.get(klasse) || existingByKlasse.get(klasse.toUpperCase()) || null;
        const chatId = existing && existing.chatId ? String(existing.chatId) : '';
        const existingTopic = existing && existing.topic ? String(existing.topic) : '';

        let eligible = true;
        let skipReason = '';
        if (memberEmails.length < 2) {
            eligible = false;
            skipReason = 'Weniger als 2 Lehrkräfte mit E-Mail';
        } else if (!topic) {
            eligible = false;
            skipReason = 'Leerer Chat-Name (Pattern prüfen)';
        }

        let status = 'neu';
        if (!eligible) status = 'skip';
        else if (chatId) {
            status =
                existingTopic && sanitizeTeamsChatTopic(existingTopic) !== sanitizeTeamsChatTopic(topic)
                    ? 'sync'
                    : 'ok';
        }

        plans.push({
            klasse,
            topic,
            members,
            memberEmails,
            eligible,
            skipReason,
            chatId,
            existingTopic,
            status
        });
    });

    return plans;
}

export function summarizePlans(plans) {
    const list = Array.isArray(plans) ? plans : [];
    return {
        total: list.length,
        eligible: list.filter((p) => p.eligible).length,
        skip: list.filter((p) => !p.eligible).length,
        neu: list.filter((p) => p.status === 'neu').length,
        sync: list.filter((p) => p.status === 'sync' || p.status === 'ok').length,
        members: list.reduce((n, p) => n + (p.memberEmails ? p.memberEmails.length : 0), 0)
    };
}

/**
 * @param {object|null} raw
 */
export function normalizeClassChatsState(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const items = [];
    const seen = new Set();
    (Array.isArray(raw.items) ? raw.items : []).forEach((it) => {
        if (!it || typeof it !== 'object') return;
        const klasse = normStr(it.klasse);
        const chatId = normStr(it.chatId);
        if (!klasse || !chatId) return;
        const key = klasse.toUpperCase();
        if (seen.has(key)) return;
        seen.add(key);
        items.push({
            klasse,
            topic: normStr(it.topic),
            chatId,
            memberEmails: Array.isArray(it.memberEmails)
                ? it.memberEmails.map(normEmail).filter(Boolean)
                : [],
            webUrl: normStr(it.webUrl),
            createdAt: normStr(it.createdAt),
            lastSyncAt: normStr(it.lastSyncAt)
        });
    });
    if (!items.length && !normStr(raw.updatedAt) && !normStr(raw.yearPrefix)) return null;
    return {
        updatedAt: normStr(raw.updatedAt) || new Date().toISOString(),
        yearPrefix: normStr(raw.yearPrefix),
        namePattern: normalizeChatNamePattern(raw.namePattern),
        items
    };
}

export function existingByKlasseFromState(state) {
    const map = new Map();
    const s = normalizeClassChatsState(state);
    if (!s) return map;
    s.items.forEach((it) => {
        map.set(it.klasse, { chatId: it.chatId, topic: it.topic });
        map.set(it.klasse.toUpperCase(), { chatId: it.chatId, topic: it.topic });
    });
    return map;
}

/**
 * E-Mails aus Freitext (Zeilen, Komma, Semikolon).
 * @param {string} text
 * @returns {string[]}
 */
export function parseMemberEmailsText(text) {
    const raw = String(text ?? '');
    const parts = raw.split(/[\s,;]+/);
    const out = [];
    const seen = new Set();
    parts.forEach((p) => {
        const em = normEmail(p);
        if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

/**
 * Einzelnen Chat-Plan manuell (Klasse + E-Mails, KV optional ergänzen).
 * @param {{
 *   klasse: string,
 *   memberEmails?: string[],
 *   membersText?: string,
 *   includeKv?: boolean,
 *   yearPrefix?: string,
 *   namePattern?: NameToken[],
 *   existingByKlasse?: Map<string, { chatId?: string, topic?: string }>,
 *   kvByKlasse?: Map<string, { email: string, name: string }>,
 *   teacherDirectory?: Array<{ code?: string, email?: string, name?: string }>
 * }} input
 * @returns {ClassChatPlan}
 */
export function buildManualClassChatPlan(input) {
    const o = input && typeof input === 'object' ? input : {};
    const klasse = normStr(o.klasse);
    const yearPrefix = normStr(o.yearPrefix) || calcYearPrefix();
    const pattern = normalizeChatNamePattern(o.namePattern);
    const existingByKlasse = o.existingByKlasse instanceof Map ? o.existingByKlasse : new Map();
    const kvMap = o.kvByKlasse instanceof Map ? o.kvByKlasse : new Map();
    const includeKv = o.includeKv !== false;

    /** @type {Map<string, ChatMember>} */
    const byEmail = new Map();

    const dir = Array.isArray(o.teacherDirectory) ? o.teacherDirectory : [];
    const codeByEmail = new Map();
    const nameByEmail = new Map();
    dir.forEach((t) => {
        const em = normEmail(t && t.email);
        if (!em) return;
        if (t.code) codeByEmail.set(em, normCode(t.code));
        if (t.name) nameByEmail.set(em, normStr(t.name));
    });

    const fromList = Array.isArray(o.memberEmails) ? o.memberEmails : [];
    const fromText = parseMemberEmailsText(o.membersText || '');
    fromList.concat(fromText).forEach((raw) => {
        const email = normEmail(raw);
        if (!email || email.indexOf('@') === -1) return;
        if (byEmail.has(email)) return;
        byEmail.set(email, {
            email,
            code: codeByEmail.get(email) || '',
            role: 'lehrer',
            name: nameByEmail.get(email) || ''
        });
    });

    if (includeKv && klasse) {
        const kv = kvMap.get(klasse) || kvMap.get(klasse.toUpperCase()) || null;
        if (kv && kv.email) {
            if (byEmail.has(kv.email)) {
                const cur = byEmail.get(kv.email);
                byEmail.set(kv.email, {
                    email: kv.email,
                    code: cur.code || '',
                    role: 'kv',
                    name: kv.name || cur.name || ''
                });
            } else {
                byEmail.set(kv.email, {
                    email: kv.email,
                    code: '',
                    role: 'kv',
                    name: kv.name || ''
                });
            }
        }
    }

    const members = Array.from(byEmail.values()).sort((a, b) => a.email.localeCompare(b.email, 'de'));
    const memberEmails = members.map((m) => m.email);
    const topic = klasse ? buildChatTopicFromPattern(pattern, { yearPrefix, klasse }) : '';
    const existing = klasse
        ? existingByKlasse.get(klasse) || existingByKlasse.get(klasse.toUpperCase()) || null
        : null;
    const chatId = existing && existing.chatId ? String(existing.chatId) : '';
    const existingTopic = existing && existing.topic ? String(existing.topic) : '';

    let eligible = true;
    let skipReason = '';
    if (!klasse) {
        eligible = false;
        skipReason = 'Klasse fehlt';
    } else if (memberEmails.length < 2) {
        eligible = false;
        skipReason = 'Weniger als 2 Lehrkräfte mit E-Mail';
    } else if (!topic) {
        eligible = false;
        skipReason = 'Leerer Chat-Name (Pattern prüfen)';
    }

    let status = 'neu';
    if (!eligible) status = 'skip';
    else if (chatId) {
        status =
            existingTopic && sanitizeTeamsChatTopic(existingTopic) !== sanitizeTeamsChatTopic(topic)
                ? 'sync'
                : 'ok';
    }

    return {
        klasse,
        topic,
        members,
        memberEmails,
        eligible,
        skipReason,
        chatId,
        existingTopic,
        status,
        source: 'manual'
    };
}

/**
 * Upsert eines Chat-Eintrags in den classChats-State.
 */
export function upsertClassChatItem(state, item) {
    const base = normalizeClassChatsState(state) || {
        updatedAt: new Date().toISOString(),
        yearPrefix: '',
        namePattern: defaultChatNamePattern(),
        items: []
    };
    const klasse = normStr(item && item.klasse);
    const chatId = normStr(item && item.chatId);
    if (!klasse || !chatId) return base;
    const next = {
        klasse,
        topic: normStr(item.topic),
        chatId,
        memberEmails: Array.isArray(item.memberEmails)
            ? item.memberEmails.map(normEmail).filter(Boolean)
            : [],
        webUrl: normStr(item.webUrl),
        createdAt: normStr(item.createdAt) || new Date().toISOString(),
        lastSyncAt: normStr(item.lastSyncAt) || new Date().toISOString()
    };
    const items = base.items.slice();
    const idx = items.findIndex((x) => x.klasse.toUpperCase() === klasse.toUpperCase());
    if (idx >= 0) {
        next.createdAt = items[idx].createdAt || next.createdAt;
        items[idx] = next;
    } else items.push(next);
    return normalizeClassChatsState({
        ...base,
        updatedAt: new Date().toISOString(),
        yearPrefix: normStr(item.yearPrefix) || base.yearPrefix,
        namePattern: item.namePattern || base.namePattern,
        items
    });
}
