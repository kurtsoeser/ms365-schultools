/**
 * Organigramm-Datenmodell + Layout-Logik (Rollen/Personen).
 * Keine Graph-Aufrufe – reine Datenlogik.
 *
 * Persistenz (kanonisch): ms365-schooltool-data-v2.core.verwaltungOrganigramm
 * Spiegel: ms365-tenant-settings-v1.verwaltungOrganigramm
 * Rollen/Personen: core.administration | adminRoles | admin
 * Browser-Backup & IT-Sicherungsbibliothek: voller v2-Container.
 *
 * Layout: { version: 1, nodes: OrgNode[] }
 * OrgNode: { id, type: 'group'|'role'|'sammel', label, parentId, order, roleCode?, … }
 */

import {
    ADMIN_TIER_SCHULLEITUNG,
    ADMIN_TIER_VERWALTUNG,
    inferAdminTierForRole
} from '../../shared/administration-audience-logic.js';

export const ORG_NODE_GROUP = 'group';
export const ORG_NODE_ROLE = 'role';
export const ORG_NODE_SAMMEL = 'sammel';

function normStr(v) {
    return String(v ?? '').trim();
}

function normEmail(v) {
    return String(v ?? '')
        .trim()
        .toLowerCase();
}

function normCode(v) {
    return normStr(v).toUpperCase();
}

export function newOrgNodeId(prefix) {
    const p = normStr(prefix) || 'og';
    return p + '-' + Math.random().toString(36).slice(2, 10);
}

/**
 * @param {unknown} raw
 * @returns {{ version: number, nodes: object[] }}
 */
export function normalizeOrganigrammLayout(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const nodesIn = Array.isArray(o.nodes) ? o.nodes : Array.isArray(o.groups) ? o.groups : [];
    const seen = new Set();
    const nodes = [];
    nodesIn.forEach(function (n, idx) {
        if (!n || typeof n !== 'object') return;
        let id = normStr(n.id);
        if (!id || seen.has(id)) id = newOrgNodeId('og');
        seen.add(id);
        const type = normStr(n.type).toLowerCase();
        const kind =
            type === ORG_NODE_ROLE || type === ORG_NODE_SAMMEL || type === ORG_NODE_GROUP
                ? type
                : ORG_NODE_GROUP;
        const parentId = n.parentId == null || n.parentId === '' ? null : normStr(n.parentId);
        const node = {
            id: id,
            type: kind,
            label: normStr(n.label || n.name) || (kind === ORG_NODE_ROLE ? 'Rolle' : 'Gruppe'),
            parentId: parentId,
            order: typeof n.order === 'number' ? n.order : idx
        };
        if (kind === ORG_NODE_ROLE) {
            node.roleCode = normCode(n.roleCode || n.code);
            node.tone = normStr(n.tone) || '';
        }
        if (kind === ORG_NODE_SAMMEL) {
            node.audienceGroupId = normStr(n.audienceGroupId);
            node.m365GroupId = normStr(n.m365GroupId);
            node.m365GroupLabel = normStr(n.m365GroupLabel);
            node.m365GroupMail = normEmail(n.m365GroupMail);
            node.tone = normStr(n.tone) || 'sammel';
        }
        if (kind === ORG_NODE_GROUP) {
            node.tone = normStr(n.tone) || 'custom';
        }
        nodes.push(node);
    });
    nodes.sort(function (a, b) {
        return (a.order || 0) - (b.order || 0);
    });
    return { version: 1, nodes: nodes };
}

/**
 * @param {{ name?: string, code?: string, tier?: string, seatMode?: string }} role
 * @param {Array<{ name?: string, email?: string, role?: string }>} people
 */
export function organigramRoleNode(role, people) {
    const list = Array.isArray(people) ? people : [];
    return {
        code: normCode(role && role.code),
        name: normStr((role && role.name) || (role && role.code) || 'Rolle'),
        tier: inferAdminTierForRole(role || {}),
        seatMode: normStr(role && role.seatMode) || 'multi',
        people: list.map(function (p) {
            return {
                name: normStr(p && p.name),
                email: normEmail(p && p.email)
            };
        }),
        vacant: list.length === 0
    };
}

/**
 * Layout aus bestehenden Rollen erzeugen (zwei Sammel-Abschnitte).
 * @param {object[]} roles
 */
export function seedOrganigrammFromRoles(roles) {
    const nodes = [];
    const leadId = newOrgNodeId('sec');
    const staffId = newOrgNodeId('sec');
    nodes.push({
        id: leadId,
        type: ORG_NODE_GROUP,
        label: 'Schulleitung',
        parentId: null,
        order: 0,
        tone: 'lead'
    });
    nodes.push({
        id: staffId,
        type: ORG_NODE_GROUP,
        label: 'Verwaltung (Personal)',
        parentId: null,
        order: 1,
        tone: 'staff'
    });
    let iLead = 0;
    let iStaff = 0;
    (Array.isArray(roles) ? roles : []).forEach(function (role) {
        const code = normCode(role && role.code);
        const name = normStr((role && role.name) || code);
        if (!code && !name) return;
        const tier = inferAdminTierForRole(role || {});
        const parentId = tier === ADMIN_TIER_SCHULLEITUNG ? leadId : staffId;
        const order = tier === ADMIN_TIER_SCHULLEITUNG ? iLead++ : iStaff++;
        nodes.push({
            id: newOrgNodeId('role'),
            type: ORG_NODE_ROLE,
            label: name || code,
            parentId: parentId,
            order: order,
            roleCode: code || name.toUpperCase().replace(/\s+/g, '').slice(0, 24),
            tone: tier === ADMIN_TIER_SCHULLEITUNG ? 'lead' : 'staff'
        });
    });
    return normalizeOrganigrammLayout({ version: 1, nodes: nodes });
}

/**
 * Fehlende Rollen-Knoten ergänzen; verwaiste Role-Knoten behalten (ohne Löschen).
 * @param {{ nodes?: object[] }|null} layout
 * @param {object[]} roles
 */
export function syncOrganigrammWithRoles(layout, roles) {
    const base = normalizeOrganigrammLayout(layout);
    const byCode = new Map();
    base.nodes.forEach(function (n) {
        if (n.type === ORG_NODE_ROLE && n.roleCode) byCode.set(normCode(n.roleCode), n);
    });
    let rootLead = base.nodes.find(function (n) {
        return n.type === ORG_NODE_GROUP && !n.parentId && (n.tone === 'lead' || /schulleitung/i.test(n.label));
    });
    let rootStaff = base.nodes.find(function (n) {
        return n.type === ORG_NODE_GROUP && !n.parentId && (n.tone === 'staff' || /verwaltung/i.test(n.label));
    });
    if (!rootLead) {
        rootLead = {
            id: newOrgNodeId('sec'),
            type: ORG_NODE_GROUP,
            label: 'Schulleitung',
            parentId: null,
            order: 0,
            tone: 'lead'
        };
        base.nodes.push(rootLead);
    }
    if (!rootStaff) {
        rootStaff = {
            id: newOrgNodeId('sec'),
            type: ORG_NODE_GROUP,
            label: 'Verwaltung (Personal)',
            parentId: null,
            order: 1,
            tone: 'staff'
        };
        base.nodes.push(rootStaff);
    }
    let maxOrder = base.nodes.reduce(function (m, n) {
        return Math.max(m, n.order || 0);
    }, 0);
    (Array.isArray(roles) ? roles : []).forEach(function (role) {
        const code = normCode(role && role.code);
        if (!code) return;
        const existing = byCode.get(code);
        if (existing) {
            const name = normStr(role && role.name);
            if (name) existing.label = name;
            return;
        }
        const tier = inferAdminTierForRole(role || {});
        maxOrder += 1;
        const node = {
            id: newOrgNodeId('role'),
            type: ORG_NODE_ROLE,
            label: normStr((role && role.name) || code),
            parentId: tier === ADMIN_TIER_SCHULLEITUNG ? rootLead.id : rootStaff.id,
            order: maxOrder,
            roleCode: code,
            tone: tier === ADMIN_TIER_SCHULLEITUNG ? 'lead' : 'staff'
        };
        base.nodes.push(node);
        byCode.set(code, node);
    });
    return normalizeOrganigrammLayout(base);
}

/**
 * Knotenbeschriftung ändern (Gruppe / Sammelgruppe / Rolle im Layout).
 * @param {{ nodes?: object[] }|null} layout
 * @param {string} nodeId
 * @param {string} label
 */
export function updateOrgNodeLabel(layout, nodeId, label) {
    const base = normalizeOrganigrammLayout(layout);
    const id = String(nodeId || '');
    const nextLabel = normStr(label);
    if (!id || !nextLabel) return base;
    return normalizeOrganigrammLayout({
        version: base.version,
        nodes: base.nodes.map(function (n) {
            if (n.id !== id) return n;
            return Object.assign({}, n, { label: nextLabel });
        })
    });
}

/**
 * @param {{ nodes: object[] }} layout
 * @param {string|null} parentId
 */
export function childrenOf(layout, parentId) {
    const pid = parentId == null || parentId === '' ? null : String(parentId);
    return normalizeOrganigrammLayout(layout)
        .nodes.filter(function (n) {
            const np = n.parentId == null || n.parentId === '' ? null : String(n.parentId);
            return np === pid;
        })
        .sort(function (a, b) {
            return (a.order || 0) - (b.order || 0);
        });
}

/**
 * Person einer anderen Rolle zuordnen (admin-Zeilen).
 * @param {object[]} rows
 * @param {string} emailOrKey – E-Mail oder Index als "idx:N"
 * @param {string} toRoleName
 * @param {string} [toRoleCode]
 */
export function movePersonToRole(rows, emailOrKey, toRoleName, toRoleCode) {
    const list = Array.isArray(rows) ? rows.slice() : [];
    const targetRole = normStr(toRoleName) || normStr(toRoleCode);
    if (!targetRole) return { rows: list, moved: false };
    const key = String(emailOrKey || '');
    let idx = -1;
    if (key.indexOf('idx:') === 0) {
        idx = parseInt(key.slice(4), 10);
    } else {
        const em = normEmail(key);
        idx = list.findIndex(function (r) {
            return normEmail(r && r.email) === em && em.indexOf('@') !== -1;
        });
        if (idx < 0) {
            idx = list.findIndex(function (r, i) {
                return String(i) === key;
            });
        }
    }
    if (idx < 0 || !list[idx]) return { rows: list, moved: false };
    const next = Object.assign({}, list[idx], {
        role: targetRole,
        defaultKey: targetRole
    });
    list[idx] = next;
    return { rows: list, moved: true, person: next };
}

/**
 * Klassische Abschnitts-Ansicht (ohne gespeichertes Layout).
 */
export function buildVerwaltungOrganigramm(roles, peopleForRole, opts) {
    const q = String((opts && opts.filter) || '')
        .trim()
        .toLowerCase();
    const getPeople = typeof peopleForRole === 'function' ? peopleForRole : function () {
        return [];
    };
    const schulleitung = [];
    const verwaltung = [];
    let peopleCount = 0;
    let vacantCount = 0;

    (Array.isArray(roles) ? roles : []).forEach(function (role) {
        const people = getPeople(role) || [];
        const node = organigramRoleNode(role, people);
        if (q) {
            const blob = [node.name, node.code]
                .concat(
                    node.people.map(function (p) {
                        return (p.name || '') + ' ' + (p.email || '');
                    })
                )
                .join(' ')
                .toLowerCase();
            if (blob.indexOf(q) === -1) return;
        }
        peopleCount += node.people.length;
        if (node.vacant) vacantCount += 1;
        if (node.tier === ADMIN_TIER_SCHULLEITUNG) schulleitung.push(node);
        else verwaltung.push(node);
    });

    function byName(a, b) {
        return String(a.name || '').localeCompare(String(b.name || ''), 'de', { sensitivity: 'base' });
    }
    schulleitung.sort(byName);
    verwaltung.sort(byName);

    const sections = [];
    if (schulleitung.length) {
        sections.push({
            id: ADMIN_TIER_SCHULLEITUNG,
            label: 'Schulleitung',
            tone: 'lead',
            roles: schulleitung
        });
    }
    if (verwaltung.length) {
        sections.push({
            id: ADMIN_TIER_VERWALTUNG,
            label: 'Verwaltung (Personal)',
            tone: 'staff',
            roles: verwaltung
        });
    }

    return {
        sections: sections,
        stats: {
            roles: schulleitung.length + verwaltung.length,
            people: peopleCount,
            vacant: vacantCount
        }
    };
}

/**
 * @param {{ nodes: object[] }} layout
 * @param {object} node
 * @param {string|null} newParentId
 * @param {number} [newOrder]
 */
export function moveNode(layout, nodeId, newParentId, newOrder) {
    const base = normalizeOrganigrammLayout(layout);
    const id = normStr(nodeId);
    const node = base.nodes.find(function (n) {
        return n.id === id;
    });
    if (!node) return base;
    if (newParentId && String(newParentId) === id) return base;
    // Zyklus verhindern
    if (newParentId) {
        let walk = base.nodes.find(function (n) {
            return n.id === String(newParentId);
        });
        let guard = 0;
        while (walk && guard++ < 50) {
            if (walk.id === id) return base;
            walk = base.nodes.find(function (n) {
                return n.id === walk.parentId;
            });
        }
    }
    node.parentId = newParentId == null || newParentId === '' ? null : String(newParentId);
    if (typeof newOrder === 'number') node.order = newOrder;
    return normalizeOrganigrammLayout(base);
}

export function addOrgNode(layout, node) {
    const base = normalizeOrganigrammLayout(layout);
    const next = normalizeOrganigrammLayout({
        nodes: base.nodes.concat([Object.assign({ id: newOrgNodeId('og'), order: base.nodes.length }, node)])
    });
    return next;
}

export function removeOrgNode(layout, nodeId, opts) {
    const cascade = !opts || opts.cascade !== false;
    const base = normalizeOrganigrammLayout(layout);
    const id = normStr(nodeId);
    if (!id) return base;
    const removeIds = new Set([id]);
    if (cascade) {
        let changed = true;
        while (changed) {
            changed = false;
            base.nodes.forEach(function (n) {
                if (n.parentId && removeIds.has(String(n.parentId)) && !removeIds.has(n.id)) {
                    removeIds.add(n.id);
                    changed = true;
                }
            });
        }
    }
    return normalizeOrganigrammLayout({
        nodes: base.nodes.filter(function (n) {
            return !removeIds.has(n.id);
        })
    });
}

const api = {
    ORG_NODE_GROUP: ORG_NODE_GROUP,
    ORG_NODE_ROLE: ORG_NODE_ROLE,
    ORG_NODE_SAMMEL: ORG_NODE_SAMMEL,
    newOrgNodeId: newOrgNodeId,
    normalizeOrganigrammLayout: normalizeOrganigrammLayout,
    organigramRoleNode: organigramRoleNode,
    seedOrganigrammFromRoles: seedOrganigrammFromRoles,
    syncOrganigrammWithRoles: syncOrganigrammWithRoles,
    childrenOf: childrenOf,
    movePersonToRole: movePersonToRole,
    moveNode: moveNode,
    addOrgNode: addOrgNode,
    removeOrgNode: removeOrgNode,
    updateOrgNodeLabel: updateOrgNodeLabel,
    buildVerwaltungOrganigramm: buildVerwaltungOrganigramm
};

if (typeof window !== 'undefined') {
    window.ms365VerwaltungOrganigramm = api;
}

export default api;
