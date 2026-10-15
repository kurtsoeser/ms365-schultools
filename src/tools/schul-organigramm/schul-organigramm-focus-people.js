/**
 * Fokus-Personen für ARGE/Klasse/Unterrichtsteam: Stammdaten-Leitung + MS365-Mitglieder/Besitzer.
 */
import {
    NODE,
    isAreaNode,
    isArgeAreaNode,
    isArgeNode,
    isClassNode,
    isClassAreaNode
} from './schul-organigramm-logic.js';

function normStr(v) {
    return String(v ?? '').trim();
}

function normEmail(v) {
    return String(v ?? '')
        .trim()
        .toLowerCase();
}

function isKursteamNode(node) {
    return !!(node && (node.type === NODE.KURSTEAM || node.type === 'kursteam'));
}

/**
 * @param {object|null|undefined} node
 * @returns {string}
 */
export function leadSeatLabelForNode(node) {
    if (!node) return '';
    if (isArgeNode(node)) return 'ARGE-Leitung';
    if (isClassNode(node)) return 'Klassenvorstand';
    if (isKursteamNode(node)) return 'Besitzer';
    return '';
}

/**
 * Mehrfach-Funktion für Gruppenmitglieder (nicht Leitung).
 * @param {object|null|undefined} node
 * @returns {string}
 */
export function memberSeatLabelForNode(node) {
    if (!node) return '';
    if (isArgeNode(node)) return 'ARGE-Mitglieder';
    if (isClassNode(node)) return 'Schüler';
    return '';
}

/**
 * Bereich mit explizit verknüpfter M365-Gruppe (match/create + groupId).
 * @param {object|null|undefined} node
 */
export function areaHasLinkedM365Group(node) {
    if (!node || !isAreaNode(node)) return false;
    const m365 = node.m365 && typeof node.m365 === 'object' ? node.m365 : null;
    if (!m365) return false;
    const mode = String(m365.groupMode || '')
        .trim()
        .toLowerCase();
    if (mode === 'none') return false;
    if (mode !== 'match' && mode !== 'create') return false;
    return !!String(m365.groupId || '').trim();
}

/**
 * @param {object|null|undefined} node
 */
export function nodeUsesGraphFocusPeople(node) {
    if (!node) return false;
    // ARGE-/Klassen-Bereich: M365-Mitglieder/Besitzer (Übersicht + Zoom)
    if (isArgeAreaNode(node) || isClassAreaNode(node)) return true;
    // Freier Bereich mit verknüpfter Gruppe → Mitglieder wie bei ARGE darstellen
    if (isAreaNode(node)) return areaHasLinkedM365Group(node);
    return (
        node.type === NODE.ARGE ||
        node.type === 'arge' ||
        node.type === NODE.CLASS ||
        node.type === 'class' ||
        isKursteamNode(node)
    );
}

/**
 * fetchGroupMembers liefert oft `{ items, truncated }` statt eines Arrays.
 * @param {unknown} raw
 * @returns {object[]}
 */
export function normalizeGraphMemberList(raw) {
    if (Array.isArray(raw)) return raw;
    if (raw && typeof raw === 'object' && Array.isArray(/** @type {{ items?: unknown }} */ (raw).items)) {
        return /** @type {{ items: object[] }} */ (raw).items;
    }
    return [];
}

/**
 * @param {object} u Graph user
 */
export function graphUserToPerson(u, kind) {
    const email = normEmail((u && (u.mail || u.userPrincipalName)) || '');
    const name = normStr((u && u.displayName) || '');
    if (!email && !name) return null;
    return {
        kind: kind || 'member',
        name: name || email,
        email: email,
        label: name || email,
        movable: false,
        fromGraph: true
    };
}

/**
 * Leitungsslot sicherstellen (auch unbesetzt) und an den Anfang setzen.
 * @param {object[]} people
 * @param {object} node
 * @param {{ name?: string, email?: string }|null} [lead]
 */
export function ensureLeadSeat(people, node, lead) {
    const seatLabel = leadSeatLabelForNode(node);
    if (!seatLabel) return Array.isArray(people) ? people.slice() : [];
    const list = Array.isArray(people) ? people.slice() : [];
    const withoutHead = list.filter(function (p) {
        return normStr(p && p.kind).toLowerCase() !== 'head';
    });
    const name = normStr(lead && lead.name);
    const email = normEmail(lead && lead.email);
    withoutHead.unshift({
        kind: 'head',
        seatLabel: seatLabel,
        name: name,
        email: email,
        label: name || email || seatLabel + ' (offen)',
        vacant: !(name || email),
        movable: false,
        scope: isKursteamNode(node) ? 'teacher' : undefined
    });
    return withoutHead;
}

/**
 * Unterrichtsteam: Besitzer (= Lehrkräfte) + Mitglieder (= Schüler:innen).
 * @param {object} opts
 */
function mergeKursteamFocusPeople(opts) {
    const o = opts || {};
    const teachers = Array.isArray(o.teachers) ? o.teachers : [];
    const teacherByEmail = new Map();
    teachers.forEach(function (t) {
        const em = normEmail(t && t.email);
        if (em) teacherByEmail.set(em, t);
    });

    /** @type {Set<string>} */
    const seen = new Set();
    /** @type {object[]} */
    const owners = [];

    (Array.isArray(o.owners) ? o.owners : []).forEach(function (u) {
        const p = graphUserToPerson(u, 'head');
        if (!p) return;
        const em = p.email;
        if (em && seen.has(em)) return;
        if (em) seen.add(em);
        const t = em ? teacherByEmail.get(em) : null;
        if (t && normStr(t.name)) p.name = normStr(t.name);
        p.seatLabel = 'Besitzer';
        p.scope = 'teacher';
        p.kind = 'head';
        p.label = p.name || p.email || 'Besitzer';
        owners.push(p);
    });

    // Lokale Lehrkraft aus Belegung / Stammdaten, falls Graph keine Owner liefert
    if (!owners.length) {
        const leadName = normStr(o.localLead && o.localLead.name);
        const leadEmail = normEmail(o.localLead && o.localLead.email);
        if (leadName || leadEmail) {
            if (leadEmail) seen.add(leadEmail);
            owners.push({
                kind: 'head',
                seatLabel: 'Besitzer',
                name: leadName,
                email: leadEmail,
                label: leadName || leadEmail || 'Besitzer (offen)',
                vacant: !(leadName || leadEmail),
                movable: false,
                scope: 'teacher'
            });
        } else {
            owners.push({
                kind: 'head',
                seatLabel: 'Besitzer',
                name: '',
                email: '',
                label: 'Besitzer (offen)',
                vacant: true,
                movable: false,
                scope: 'teacher'
            });
        }
    }

    /** @type {object[]} */
    const members = [];
    (Array.isArray(o.members) ? o.members : []).forEach(function (u) {
        const p = graphUserToPerson(u, 'student');
        if (!p || !p.email || seen.has(p.email)) return;
        seen.add(p.email);
        // Lehrkräfte in der Mitgliederliste nicht als Schüler zeigen
        if (teacherByEmail.has(p.email)) return;
        members.push(p);
    });

    (Array.isArray(o.localPeople) ? o.localPeople : []).forEach(function (p) {
        if (!p) return;
        if (normStr(p.kind).toLowerCase() === 'head') return;
        const em = normEmail(p.email);
        if (em && seen.has(em)) return;
        if (em) seen.add(em);
        members.push(
            Object.assign({}, p, {
                kind: 'student',
                scope: 'student',
                movable: false
            })
        );
    });

    members.sort(function (a, b) {
        return normStr(a.name || a.email).localeCompare(normStr(b.name || b.email), 'de');
    });

    return owners.concat(members);
}

/**
 * Lokale + Graph-Mitglieder zusammenführen; Besitzer → Leitung wenn Stammdaten leer.
 * @param {{
 *   localPeople?: object[],
 *   members?: object[],
 *   owners?: object[],
 *   node: object,
 *   teachers?: object[],
 *   localLead?: { name?: string, email?: string }|null
 * }} opts
 */
export function mergeGraphFocusPeople(opts) {
    const o = opts || {};
    const node = o.node;
    if (isKursteamNode(node)) {
        return mergeKursteamFocusPeople(o);
    }

    const seatLabel = leadSeatLabelForNode(node);
    const memberSeat = memberSeatLabelForNode(node);
    const teachers = Array.isArray(o.teachers) ? o.teachers : [];
    const teacherByEmail = new Map();
    teachers.forEach(function (t) {
        const em = normEmail(t && t.email);
        if (em) teacherByEmail.set(em, t);
    });

    const memberList = normalizeGraphMemberList(o.members);
    const ownerList = Array.isArray(o.owners) ? o.owners : [];

    let leadName = normStr(o.localLead && o.localLead.name);
    let leadEmail = normEmail(o.localLead && o.localLead.email);
    if (!leadEmail && !leadName && ownerList.length) {
        for (let i = 0; i < ownerList.length; i += 1) {
            const em = normEmail(ownerList[i] && (ownerList[i].mail || ownerList[i].userPrincipalName));
            if (!em) continue;
            const t = teacherByEmail.get(em);
            leadEmail = em;
            leadName = (t && normStr(t.name)) || normStr(ownerList[i].displayName) || em;
            break;
        }
    }

    const seen = new Set();
    if (leadEmail) seen.add(leadEmail);

    /** @type {object[]} */
    const members = [];
    const memberKind =
        node && (node.type === NODE.CLASS || node.type === 'class' || isClassAreaNode(node))
            ? 'student'
            : 'member';

    memberList.forEach(function (u) {
        const p = graphUserToPerson(u, memberKind);
        if (!p || !p.email || seen.has(p.email)) return;
        seen.add(p.email);
        const t = teacherByEmail.get(p.email);
        if (t && normStr(t.name)) p.name = normStr(t.name);
        if (memberSeat) p.seatLabel = memberSeat;
        members.push(p);
    });

    // Weitere Besitzer (nicht Leitung) auch anzeigen, falls nicht in members
    ownerList.forEach(function (u) {
        const p = graphUserToPerson(u, 'member');
        if (!p || !p.email || seen.has(p.email)) return;
        seen.add(p.email);
        const t = teacherByEmail.get(p.email);
        if (t && normStr(t.name)) p.name = normStr(t.name);
        p.seatLabel = memberSeat || p.seatLabel || 'Besitzer';
        members.push(p);
    });

    // Lokale Personen (ohne head) ergänzen, die nicht in Graph sind
    (Array.isArray(o.localPeople) ? o.localPeople : []).forEach(function (p) {
        if (!p) return;
        if (normStr(p.kind).toLowerCase() === 'head') return;
        const em = normEmail(p.email);
        if (em && seen.has(em)) return;
        if (em) seen.add(em);
        const next = Object.assign({}, p, {
            kind: p.kind || 'member',
            movable: false
        });
        if (memberSeat && !normStr(next.seatLabel)) next.seatLabel = memberSeat;
        members.push(next);
    });

    members.sort(function (a, b) {
        return normStr(a.name || a.email).localeCompare(normStr(b.name || b.email), 'de');
    });

    if (!seatLabel) return members;
    return ensureLeadSeat(members, node, { name: leadName, email: leadEmail });
}
