/**
 * Personen-Fokus: alle Mitgliedschaften / Funktionen einer Person auf der Landkarte.
 */
import {
    NODE,
    childrenOf,
    isAreaNode,
    isArgeNode,
    isClassNode,
    normalizeSchulOrganigramm
} from './schul-organigramm-logic.js';
import { listPeopleForSchulOrgNode } from './schul-organigramm-edit.js';

function normStr(v) {
    return String(v ?? '').trim();
}

function normEmail(v) {
    return String(v ?? '')
        .trim()
        .toLowerCase();
}

/**
 * Stabiler Schlüssel für Person (E-Mail bevorzugt).
 * @param {{ email?: string, name?: string }|null|undefined} person
 */
export function schulOrgPersonFocusKey(person) {
    const em = normEmail(person && person.email);
    if (em) return 'e:' + em;
    const name = normStr(person && person.name).toLowerCase();
    if (name) return 'n:' + name;
    return '';
}

/**
 * @param {{ email?: string, name?: string }|null|undefined} a
 * @param {{ email?: string, name?: string }|null|undefined} b
 */
export function schulOrgPersonFocusMatches(a, b) {
    const ka = schulOrgPersonFocusKey(a);
    const kb = schulOrgPersonFocusKey(b);
    if (!ka || !kb) return false;
    if (ka === kb) return true;
    // Fallback: gleiche E-Mail auch wenn Keys unterschiedlich gebaut wurden
    const ea = normEmail(a && a.email);
    const eb = normEmail(b && b.email);
    if (ea && eb && ea === eb) return true;
    return false;
}

/**
 * @param {object} node
 */
function membershipKindLabel(node, personHit) {
    const t = String((node && node.type) || '').toLowerCase();
    const seat = normStr(personHit && personHit.seatLabel);
    if (seat) return seat;
    if (t === NODE.ROLE || t === 'role') return 'Rolle';
    if (isArgeNode(node)) {
        const kind = String((personHit && personHit.kind) || '').toLowerCase();
        if (kind === 'teacher' || kind === 'member') return 'ARGE-Mitglied';
        return 'ARGE-Leitung';
    }
    if (isClassNode(node) || t === NODE.CLASS || t === 'class') {
        const kind = String((personHit && personHit.kind) || '').toLowerCase();
        if (kind === 'student') return 'Schüler:in';
        return 'Klassenvorstand';
    }
    if (t === NODE.KURSTEAM || t === 'kursteam') return 'Unterrichtsteam';
    if (t === NODE.COUNCIL || t === 'council') return 'Schülervertretung';
    if (t === NODE.SGA || t === 'sga') return 'SGA';
    if (isAreaNode(node)) {
        const kind = String((personHit && personHit.kind) || '').toLowerCase();
        if (kind === 'teacher') return 'Lehrkraft';
        if (kind === 'sga') return 'SGA';
        if (kind === 'council') return 'SV';
        return 'Bereichsmitglied';
    }
    return 'Mitgliedschaft';
}

/**
 * @param {object} layout
 * @param {string} nodeId
 * @returns {string[]}
 */
function ancestorPathIds(layout, nodeId) {
    const byId = new Map();
    (layout.nodes || []).forEach(function (n) {
        if (n && n.id) byId.set(String(n.id), n);
    });
    /** @type {string[]} */
    const path = [];
    let cur = byId.get(String(nodeId || ''));
    let guard = 0;
    while (cur && guard++ < 40) {
        path.push(String(cur.id));
        const pid = cur.parentId != null ? String(cur.parentId) : '';
        if (!pid) break;
        cur = byId.get(pid);
    }
    return path;
}

/**
 * @param {object} layout
 * @param {string} nodeId
 * @returns {string[]}
 */
function ancestorPathLabels(layout, nodeId) {
    const byId = new Map();
    (layout.nodes || []).forEach(function (n) {
        if (n && n.id) byId.set(String(n.id), n);
    });
    /** @type {string[]} */
    const labels = [];
    let cur = byId.get(String(nodeId || ''));
    let guard = 0;
    while (cur && guard++ < 40) {
        const lab = normStr(cur.label);
        if (lab) labels.unshift(lab);
        const pid = cur.parentId != null ? String(cur.parentId) : '';
        if (!pid) break;
        cur = byId.get(pid);
    }
    return labels;
}

/**
 * Verzeichnis aller Personen aus den Stammdaten (für Picker).
 * @param {object} settings
 * @returns {{ name: string, email: string, sources: string[] }[]}
 */
export function listSchulOrgDirectoryPeople(settings) {
    const s = settings || {};
    /** @type {Map<string, { name: string, email: string, sources: Set<string> }>} */
    const map = new Map();

    function add(name, email, source) {
        const em = normEmail(email);
        const nm = normStr(name);
        if (!em && !nm) return;
        const key = em ? 'e:' + em : 'n:' + nm.toLowerCase();
        let row = map.get(key);
        if (!row) {
            row = { name: nm, email: em, sources: new Set() };
            map.set(key, row);
        }
        if (nm && (!row.name || row.name.length < nm.length)) row.name = nm;
        if (em && !row.email) row.email = em;
        if (source) row.sources.add(source);
    }

    (Array.isArray(s.admin) ? s.admin : []).forEach(function (p) {
        add(p && p.name, p && p.email, 'Verwaltung');
    });
    (Array.isArray(s.teachers) ? s.teachers : []).forEach(function (p) {
        add(p && p.name, p && p.email, 'Lehrer');
    });
    (Array.isArray(s.students) ? s.students : []).forEach(function (p) {
        add(p && p.name, p && p.email, 'Schüler');
    });
    (Array.isArray(s.studentCouncil) ? s.studentCouncil : []).forEach(function (p) {
        add(p && p.name, p && p.email, 'SV');
    });
    (Array.isArray(s.sga) ? s.sga : []).forEach(function (p) {
        add(p && p.name, p && p.email, 'SGA');
    });
    (Array.isArray(s.adminAudienceMemberships) ? s.adminAudienceMemberships : []).forEach(
        function (m) {
            add(m && m.name, m && m.email, 'Bereich');
        }
    );
    (Array.isArray(s.arges) ? s.arges : []).forEach(function (a) {
        add(a && a.headName, a && a.headEmail, 'ARGE');
    });
    (Array.isArray(s.classes) ? s.classes : []).forEach(function (c) {
        add(c && c.headName, c && c.headEmail, 'Klasse');
    });

    return Array.from(map.values())
        .map(function (r) {
            return {
                name: r.name,
                email: r.email,
                sources: Array.from(r.sources).sort(function (a, b) {
                    return a.localeCompare(b, 'de');
                })
            };
        })
        .sort(function (a, b) {
            return (a.name || a.email).localeCompare(b.name || b.email, 'de');
        });
}

/**
 * Alle Knoten-Mitgliedschaften einer Person + Highlight-Mengen.
 * @param {object} settings
 * @param {unknown} layout
 * @param {{ email?: string, name?: string }|null|undefined} person
 * @returns {{
 *   person: { name: string, email: string },
 *   memberships: Array<{
 *     nodeId: string,
 *     nodeLabel: string,
 *     nodeType: string,
 *     kindLabel: string,
 *     personKind: string,
 *     seatLabel: string,
 *     pathLabels: string[],
 *     pathText: string
 *   }>,
 *   hitIds: string[],
 *   pathIds: string[]
 * }}
 */
export function collectSchulOrgPersonMemberships(settings, layout, person) {
    const focus = {
        name: normStr(person && person.name),
        email: normEmail(person && person.email)
    };
    const org = normalizeSchulOrganigramm(layout || (settings && settings.schulOrganigramm));
    /** @type {ReturnType<typeof collectSchulOrgPersonMemberships>['memberships']} */
    const memberships = [];
    /** @type {Set<string>} */
    const hitIds = new Set();
    /** @type {Set<string>} */
    const pathIds = new Set();

    if (!schulOrgPersonFocusKey(focus)) {
        return {
            person: focus,
            memberships: [],
            hitIds: [],
            pathIds: []
        };
    }

    (org.nodes || []).forEach(function (node) {
        if (!node || !node.id) return;
        // Lehrer-/Schüler-Sammelbereiche bewusst mitnehmen (Übersicht „gehört zu Lehrkräfte“)
        const people = listPeopleForSchulOrgNode(settings, node);
        const hit = people.find(function (p) {
            return !p.vacant && schulOrgPersonFocusMatches(focus, p);
        });
        if (!hit) return;
        hitIds.add(String(node.id));
        ancestorPathIds(org, node.id).forEach(function (id) {
            pathIds.add(id);
        });
        const pathLabels = ancestorPathLabels(org, node.id);
        const kindLabel = membershipKindLabel(node, hit);
        memberships.push({
            nodeId: String(node.id),
            nodeLabel: normStr(node.label) || String(node.id),
            nodeType: String(node.type || ''),
            kindLabel: kindLabel,
            personKind: String(hit.kind || ''),
            seatLabel: normStr(hit.seatLabel),
            pathLabels: pathLabels,
            pathText: pathLabels.join(' › ')
        });
    });

    memberships.sort(function (a, b) {
        const ta = a.pathText.localeCompare(b.pathText, 'de');
        if (ta) return ta;
        return a.nodeLabel.localeCompare(b.nodeLabel, 'de');
    });

    return {
        person: focus,
        memberships: memberships,
        hitIds: Array.from(hitIds),
        pathIds: Array.from(pathIds)
    };
}

/**
 * @param {object} settings
 * @param {string} query
 * @param {number} [limit]
 */
export function searchSchulOrgDirectoryPeople(settings, query, limit) {
    const q = normStr(query).toLowerCase();
    const max = typeof limit === 'number' && limit > 0 ? limit : 40;
    const all = listSchulOrgDirectoryPeople(settings);
    if (!q) return all.slice(0, max);
    return all
        .filter(function (p) {
            const hay = ((p.name || '') + ' ' + (p.email || '') + ' ' + (p.sources || []).join(' '))
                .toLowerCase();
            return hay.indexOf(q) !== -1;
        })
        .slice(0, max);
}

export const api = {
    schulOrgPersonFocusKey: schulOrgPersonFocusKey,
    schulOrgPersonFocusMatches: schulOrgPersonFocusMatches,
    listSchulOrgDirectoryPeople: listSchulOrgDirectoryPeople,
    searchSchulOrgDirectoryPeople: searchSchulOrgDirectoryPeople,
    collectSchulOrgPersonMemberships: collectSchulOrgPersonMemberships
};

export default api;
