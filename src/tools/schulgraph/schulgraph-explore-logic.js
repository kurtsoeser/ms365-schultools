/**
 * Graph-Zusammenführung & Layout für M365-Expansion (Klick → Nachbarschaft).
 */

/**
 * @param {{ nodes: object[], edges: object[] }} base
 * @param {{ nodes: object[], edges: object[] }} extra
 */
export function mergeGraphLayers(base, extra) {
    const nodes = new Map();
    (base.nodes || []).forEach((n) => {
        if (n && n.id) nodes.set(n.id, n);
    });
    (extra.nodes || []).forEach((n) => {
        if (n && n.id) nodes.set(n.id, n);
    });
    const edgeKeys = new Set();
    const edges = [];
    function push(e) {
        if (!e || !e.source || !e.target) return;
        const key = [e.kind, e.source, e.target].join('|');
        if (edgeKeys.has(key)) return;
        edgeKeys.add(key);
        edges.push(e);
    }
    (base.edges || []).forEach(push);
    (extra.edges || []).forEach(push);
    return { nodes: [...nodes.values()], edges };
}

/**
 * @param {Map<string, { x: number, y: number }>} positions
 * @param {{ x: number, y: number }} center
 * @param {string[]} nodeIds
 * @param {number} [radius]
 */
export function placeNodesAround(positions, center, nodeIds, radius) {
    const r = radius || 130;
    const cx = center.x;
    const cy = center.y;
    const list = nodeIds || [];
    list.forEach((id, i) => {
        const angle = (i / Math.max(list.length, 1)) * Math.PI * 2 - Math.PI / 2;
        positions.set(id, {
            x: cx + Math.cos(angle) * r,
            y: cy + Math.sin(angle) * r
        });
    });
}

/**
 * @param {string} classCode
 * @param {object} setup
 */
export function resolveClassGraphGroupId(classCode, setup) {
    const code = String(classCode || '')
        .trim()
        .toUpperCase();
    if (!code) return '';
    const links = setup && Array.isArray(setup.catalogLinks) ? setup.catalogLinks : [];
    const hit = links.find((L) => String(L.kind || '') === 'class' && String(L.code || '').trim().toUpperCase() === code);
    if (hit && hit.graphGroupId) return String(hit.graphGroupId).trim();
    const match = setup && setup.classGroupMatchByKey && typeof setup.classGroupMatchByKey === 'object'
        ? setup.classGroupMatchByKey[code]
        : null;
    if (match && match.id) return String(match.id).trim();
    if (match && match.graphGroupId) return String(match.graphGroupId).trim();
    return '';
}

/**
 * Stammdaten-Fallback für Klasse ohne Graph-Gruppe.
 * @param {string} classCode
 * @param {object} settings
 * @param {object} yearBucket
 */
export function stammdatenMembersForClass(classCode, settings, yearBucket) {
    const code = String(classCode || '')
        .trim()
        .toUpperCase();
    const studentsYear = Array.isArray(yearBucket.students) ? yearBucket.students : [];
    const studentsCore = Array.isArray(settings.students) ? settings.students : [];
    const students = studentsYear.length ? studentsYear : studentsCore;
    const rows = [];
    students.forEach((s) => {
        const k = String(s.klasse || s.class || '')
            .trim()
            .toUpperCase();
        if (k !== code) return;
        const email = String(s.email || '')
            .trim()
            .toLowerCase();
        const name = String(s.name || email || 'Schüler:in').trim();
        if (!name && !email) return;
        rows.push({ name, email, source: 'stammdaten', role: 'student' });
    });
    return rows;
}

export function entraUserNodeId(graphUserId) {
    return `entraUser:${String(graphUserId || '').trim()}`;
}

export function graphGroupNodeId(graphGroupId) {
    return `m365group:${String(graphGroupId || '').trim()}`;
}
