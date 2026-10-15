/**
 * Stammdaten-Verwaltung: Baum aus Schul-Organigramm (Bereiche → Unterbereiche → Funktionen → Personen).
 * Quelle der Wahrheit für die Struktur: core.schulOrganigramm (+ admin / adminRoles).
 */
import {
    audienceIdForArea,
    childrenOf,
    isAreaNode,
    listSchulOrgRoleParentAreas,
    normalizeSchulOrganigramm,
    NODE
} from '../tools/schul-organigramm/schul-organigramm-logic.js';
import { normalizeAdminAudienceMemberships } from './administration-audience-groups.js';

function normStr(v) {
    return String(v ?? '').trim();
}

function normEmail(v) {
    return normStr(v).toLowerCase();
}

/**
 * @param {object} settings
 * @param {{ code?: string, name?: string }|null|undefined} role
 */
function peopleForAdminRole(settings, role) {
    if (!role) return [];
    const name = normStr(role.name).toLowerCase();
    const code = normStr(role.code).toLowerCase();
    return (Array.isArray(settings.admin) ? settings.admin : []).filter(function (p) {
        const r = normStr(p && p.role).toLowerCase();
        const dk = normStr(p && p.defaultKey).toLowerCase();
        const hasIdentity = normStr(p && p.name) || (normEmail(p && p.email).indexOf('@') !== -1);
        if (!hasIdentity) return false;
        return (name && (r === name || dk === name)) || (code && (r === code || dk === code));
    });
}

/**
 * @param {object} settings
 * @param {string} code
 */
function findAdminRole(settings, code) {
    const want = normStr(code).toUpperCase();
    if (!want) return null;
    return (
        (Array.isArray(settings.adminRoles) ? settings.adminRoles : []).find(function (r) {
            return normStr(r && r.code).toUpperCase() === want;
        }) || null
    );
}

/**
 * Direkte Bereichspersonen (Audience + Area-Slots), ohne Rollenbesetzungen.
 * @param {object} settings
 * @param {object} areaNode
 * @param {object} layout
 */
function directAreaPeople(settings, areaNode, layout) {
    /** @type {Array<{ name: string, email: string }>} */
    const out = [];
    const seen = new Set();
    const audId = audienceIdForArea(areaNode);
    if (audId) {
        normalizeAdminAudienceMemberships(settings.adminAudienceMemberships).forEach(function (m) {
            if (String(m.groupId) !== audId) return;
            const em = normEmail(m.email);
            if (!em || seen.has(em)) return;
            seen.add(em);
            out.push({ name: normStr(m.name), email: em });
        });
    }
    childrenOf(layout, areaNode.id).forEach(function (ch) {
        if (!ch || ch.type !== NODE.PERSON || normStr(ch.refKind) !== 'adminPerson') return;
        const em = normEmail(ch.refKey);
        if (!em || seen.has(em)) return;
        seen.add(em);
        out.push({ name: normStr(ch.label), email: em });
    });
    return out.sort(function (a, b) {
        return (a.name || a.email).localeCompare(b.name || b.email, 'de');
    });
}

/**
 * Verwaltungs-Baum aus Organigramm + Stammdaten-Rollen/Personen.
 * @param {object} settings
 * @returns {{
 *   roots: Array<object>,
 *   counts: { areas: number, roles: number, people: number },
 *   source: 'organigramm' | 'empty'
 * }}
 */
export function buildVerwaltungStammdatenTree(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const layout = normalizeSchulOrganigramm(s.schulOrganigramm);
    const parentAreas = listSchulOrgRoleParentAreas(layout);
    /** @type {Map<string, object>} */
    const areaById = new Map();
    parentAreas.forEach(function (a) {
        if (a && a.id) areaById.set(String(a.id), a);
    });

    /** @type {Set<string>} */
    const onRoles = new Set();
    let roleCount = 0;
    let peopleCount = 0;

    /**
     * @param {object} area
     * @param {number} depth
     */
    function buildAreaNode(area, depth) {
        const areaId = String(area.id);
        /** @type {object[]} */
        const children = [];

        // Unterbereiche (nur Verwaltungs-/Rollen-fähige)
        childrenOf(layout, areaId)
            .filter(function (ch) {
                return isAreaNode(ch) && areaById.has(String(ch.id));
            })
            .sort(function (a, b) {
                return (a.order || 0) - (b.order || 0) ||
                    normStr(a.label).localeCompare(normStr(b.label), 'de');
            })
            .forEach(function (ch) {
                children.push(buildAreaNode(ch, depth + 1));
            });

        // Funktionen / Rollen
        childrenOf(layout, areaId)
            .filter(function (ch) {
                const t = normStr(ch && ch.type).toLowerCase();
                return t === NODE.ROLE || t === 'role';
            })
            .sort(function (a, b) {
                return (a.order || 0) - (b.order || 0) ||
                    normStr(a.label).localeCompare(normStr(b.label), 'de');
            })
            .forEach(function (roleNode) {
                roleCount += 1;
                const role = findAdminRole(s, roleNode.refKey) || {
                    code: normStr(roleNode.refKey),
                    name: normStr(roleNode.label) || normStr(roleNode.refKey)
                };
                const people = peopleForAdminRole(s, role).map(function (p) {
                    const em = normEmail(p.email);
                    if (em) onRoles.add(em);
                    peopleCount += 1;
                    return {
                        kind: 'person',
                        id: 'person:' + (em || normStr(p.name)),
                        label: normStr(p.name) || em || 'Person',
                        email: em,
                        depth: depth + 2,
                        source: 'role',
                        sourceRoleId: String(roleNode.id),
                        sourceRoleCode: normStr(role.code) || normStr(roleNode.refKey),
                        children: []
                    };
                });
                children.push({
                    kind: 'role',
                    id: String(roleNode.id),
                    label: normStr(roleNode.label) || normStr(role.name) || normStr(role.code),
                    code: normStr(role.code),
                    depth: depth + 1,
                    children: people
                });
            });

        // Direkte Bereichspersonen (nicht schon auf Rollen)
        directAreaPeople(s, area, layout).forEach(function (p) {
            if (p.email && onRoles.has(p.email)) return;
            peopleCount += 1;
            children.push({
                kind: 'person',
                id: 'area-person:' + areaId + ':' + p.email,
                label: p.name || p.email,
                email: p.email,
                depth: depth + 1,
                source: 'area',
                sourceAreaId: areaId,
                children: []
            });
        });

        const audId = audienceIdForArea(area);
        return {
            kind: 'area',
            id: areaId,
            label: normStr(area.label) || areaId,
            depth: depth,
            audienceId: audId || '',
            children: children
        };
    }

    const roots = parentAreas
        .filter(function (a) {
            const pid = a.parentId == null || a.parentId === '' ? '' : String(a.parentId);
            return !pid || !areaById.has(pid);
        })
        .sort(function (a, b) {
            return (a.order || 0) - (b.order || 0) ||
                normStr(a.label).localeCompare(normStr(b.label), 'de');
        })
        .map(function (a) {
            return buildAreaNode(a, 0);
        });

    return {
        roots: roots,
        counts: {
            areas: parentAreas.length,
            roles: roleCount,
            people: peopleCount
        },
        source: parentAreas.length ? 'organigramm' : 'empty'
    };
}

export default {
    buildVerwaltungStammdatenTree: buildVerwaltungStammdatenTree
};
