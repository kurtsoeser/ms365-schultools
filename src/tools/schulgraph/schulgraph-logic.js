/**
 * Stammdaten → Graph (Knoten/Kanten) für die interaktive Karte.
 * Rein funktional – ohne DOM, gut testbar.
 */
import { DEFAULT_GRAPH_OPTIONS } from './schulgraph-schema.js';
import {
    normalizeAspects,
    aspectsFromLegacyLayers,
    detectPreset,
    DEFAULT_ASPECTS,
    aspectsForPreset
} from './schulgraph-aspects.js';

export { aspectsForPreset, detectPreset, GRAPH_PRESETS, ASPECT_CATALOG } from './schulgraph-aspects.js';

function normCode(v) {
    return String(v ?? '')
        .trim()
        .toUpperCase();
}

function normClassKey(v) {
    return String(v ?? '')
        .trim()
        .toUpperCase();
}

function normEmail(v) {
    return String(v ?? '')
        .trim()
        .toLowerCase();
}

/**
 * @param {object} raw
 * @returns {Required<typeof DEFAULT_GRAPH_OPTIONS>}
 */
export function normalizeGraphOptions(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const layersIn = o.layers && typeof o.layers === 'object' ? o.layers : {};
    const base = DEFAULT_GRAPH_OPTIONS;

    let aspects = normalizeAspects(o.aspects || DEFAULT_ASPECTS);
    if (!o.aspects && (o.layers || layersIn.org !== undefined)) {
        aspects = aspectsFromLegacyLayers(layersIn);
    }

    let preset = String(o.preset || '').trim();
    if (!preset || preset === 'custom') {
        preset = detectPreset(aspects);
    }

    return {
        preset,
        aspects,
        layers: {
            org: aspects.schulorganisation || aspects.klassenvorstand || aspects.fachgruppen,
            teaching: aspects.unterricht,
            people: aspects.schueler || aspects.eltern,
            m365: aspects.microsoft365
        },
        klasseFilter: normClassKey(o.klasseFilter || ''),
        peopleAutoOffThreshold: Number.isFinite(Number(o.peopleAutoOffThreshold))
            ? Number(o.peopleAutoOffThreshold)
            : base.peopleAutoOffThreshold,
        maxStudents: Number.isFinite(Number(o.maxStudents)) ? Number(o.maxStudents) : base.maxStudents,
        maxGuardians: Number.isFinite(Number(o.maxGuardians)) ? Number(o.maxGuardians) : base.maxGuardians,
        includeSchoolHub: o.includeSchoolHub !== false
    };
}

/**
 * @typedef {{ id: string, kind: string, label: string, sublabel?: string, meta?: object }} GraphNode
 * @typedef {{ id: string, source: string, target: string, kind: string, label?: string }} GraphEdge
 */

/**
 * @param {object} params
 * @param {object|null} params.settings ms365TenantSettingsLoad()
 * @param {object|null} params.yearBucket app-data years.* bucket
 * @param {object|null} params.setup app-data setup (catalogLinks)
 * @param {object} [params.options]
 * @returns {{ nodes: GraphNode[], edges: GraphEdge[], stats: object, warnings: string[] }}
 */
export function buildSchulGraph(params) {
    const settings = params && params.settings && typeof params.settings === 'object' ? params.settings : {};
    const yearBucket = params && params.yearBucket && typeof params.yearBucket === 'object' ? params.yearBucket : {};
    const setup = params && params.setup && typeof params.setup === 'object' ? params.setup : {};
    const options = normalizeGraphOptions(params && params.options);
    const asp = options.aspects;

    /** @type {Map<string, GraphNode>} */
    const nodes = new Map();
    /** @type {GraphEdge[]} */
    const edges = [];
    const warnings = [];
    const edgeKeys = new Set();

    function addNode(node) {
        if (!node || !node.id) return;
        if (!nodes.has(node.id)) nodes.set(node.id, node);
    }

    function addEdge(edge) {
        if (!edge || !edge.source || !edge.target || edge.source === edge.target) return;
        const key = [edge.kind, edge.source, edge.target].join('|');
        if (edgeKeys.has(key)) return;
        edgeKeys.add(key);
        if (!edge.id) edge.id = key;
        edges.push(edge);
    }

    const classes = (Array.isArray(settings.classes) ? settings.classes : []).filter(Boolean);
    const subjects = (Array.isArray(settings.subjects) ? settings.subjects : []).filter(Boolean);
    const arges = (Array.isArray(settings.arges) ? settings.arges : []).filter(Boolean);
    const teachers = (Array.isArray(settings.teachers) ? settings.teachers : []).filter(Boolean);
    const studentsCore = Array.isArray(settings.students) ? settings.students : [];
    const studentsYear = Array.isArray(yearBucket.students) ? yearBucket.students : [];
    const students = studentsYear.length ? studentsYear : studentsCore;
    const guardians = Array.isArray(yearBucket.guardians) ? yearBucket.guardians : [];
    const ub = yearBucket.unterrichtsbelegung && Array.isArray(yearBucket.unterrichtsbelegung.rows)
        ? yearBucket.unterrichtsbelegung.rows
        : [];
    const catalogLinks = Array.isArray(setup.catalogLinks) ? setup.catalogLinks : [];

    const klasseFilter = options.klasseFilter;
    const filteredClasses = klasseFilter
        ? classes.filter((c) => normClassKey(c.code) === klasseFilter)
        : classes;

    let peopleLayer = asp.schueler;
    if (!klasseFilter && students.length > options.peopleAutoOffThreshold) {
        peopleLayer = false;
        if (asp.schueler) {
            warnings.push(
                `Schüler:innen (${students.length}) – Aspekt „Schüler:innen“ deaktiviert. Klasse filtern oder kleinere Datenmenge wählen.`
            );
        }
    }

    const showSchoolHub =
        options.includeSchoolHub &&
        (asp.schulorganisation || asp.microsoft365 || asp.klassenvorstand || asp.fachgruppen);
    const needTeachers = asp.klassenvorstand || asp.unterricht;

    const schoolName = String(settings.schoolName || 'Schule').trim() || 'Schule';
    if (showSchoolHub) {
        addNode({
            id: 'school:root',
            kind: 'school',
            label: schoolName,
            sublabel: String(settings.domain || '').trim() || undefined
        });
    }

    /** @type {Map<string, string>} code → teacher node id */
    const teacherByCode = new Map();
    /** @type {Map<string, string>} email → teacher node id */
    const teacherByEmail = new Map();

    if (needTeachers) {
        teachers.forEach((t) => {
            const code = normCode(t.code);
            if (!code) return;
            const id = `teacher:${code}`;
            teacherByCode.set(code, id);
            const em = normEmail(t.email);
            if (em) teacherByEmail.set(em, id);
            addNode({
                id,
                kind: 'teacher',
                label: String(t.name || code).trim() || code,
                sublabel: code,
                meta: { code, email: em || undefined }
            });
        });
    }

    if (asp.schulorganisation || asp.klassenvorstand || asp.schueler) {
        filteredClasses.forEach((c) => {
            const code = normClassKey(c.code);
            if (!code) return;
            const id = `class:${code}`;
            addNode({
                id,
                kind: 'class',
                label: String(c.name || code).trim() || code,
                sublabel: c.year ? `JG ${c.year}` : code,
                meta: { code, year: c.year || '' }
            });
            if (showSchoolHub && asp.schulorganisation) {
                addEdge({ kind: 'belongs_to', source: id, target: 'school:root' });
            }
            if (asp.klassenvorstand) {
                const headEmail = normEmail(c.headEmail);
                if (headEmail && teacherByEmail.has(headEmail)) {
                    addEdge({
                        kind: 'kv_of',
                        source: teacherByEmail.get(headEmail),
                        target: id,
                        label: 'KV'
                    });
                }
            }
        });
    }

    if (asp.schulorganisation || asp.fachgruppen) {
        if (asp.schulorganisation) {
            subjects.forEach((s) => {
                const code = normCode(s.code);
                if (!code) return;
                const id = `subject:${code}`;
                addNode({
                    id,
                    kind: 'subject',
                    label: String(s.name || code).trim() || code,
                    sublabel: code
                });
                if (showSchoolHub) {
                    addEdge({ kind: 'belongs_to', source: id, target: 'school:root' });
                }
            });
        }

        arges.forEach((a) => {
            const code = normCode(a.code);
            if (!code) return;
            const id = `arge:${code}`;
            if (asp.schulorganisation) {
                addNode({
                    id,
                    kind: 'arge',
                    label: String(a.name || code).trim() || code,
                    sublabel: code
                });
                if (showSchoolHub) {
                    addEdge({ kind: 'belongs_to', source: id, target: 'school:root' });
                }
            }
            if (!asp.fachgruppen) return;
            const subjList = Array.isArray(a.subjects) ? a.subjects : [];
            subjList.forEach((sc) => {
                const sCode = normCode(sc);
                if (!sCode) return;
                const sId = `subject:${sCode}`;
                if (!nodes.has(sId)) {
                    addNode({ id: sId, kind: 'subject', label: sCode, sublabel: sCode });
                }
                if (!nodes.has(id)) {
                    addNode({
                        id,
                        kind: 'arge',
                        label: String(a.name || code).trim() || code,
                        sublabel: code
                    });
                }
                addEdge({ kind: 'subject_in_arge', source: sId, target: id });
            });
        });
    }

    if (asp.unterricht && ub.length) {
        ub.forEach((row) => {
            const klasse = normClassKey(row.klasse);
            const fachRaw = String(row.fach || '').trim();
            const fachCode = normCode(fachRaw);
            const classId = klasse ? `class:${klasse}` : '';
            if (klasseFilter && klasse !== klasseFilter) return;

            let teacherId = '';
            const lc = normCode(row.lehrerCode);
            if (lc && teacherByCode.has(lc)) teacherId = teacherByCode.get(lc);
            else {
                const em = normEmail(row.lehrerEmail);
                if (em && teacherByEmail.has(em)) teacherId = teacherByEmail.get(em);
            }

            if (klasse && classId && !nodes.has(classId)) {
                addNode({ id: classId, kind: 'class', label: klasse, sublabel: klasse });
            }

            let subjectId = '';
            if (fachCode) {
                subjectId = `subject:${fachCode}`;
                if (!nodes.has(subjectId)) {
                    addNode({ id: subjectId, kind: 'subject', label: fachRaw || fachCode, sublabel: fachCode });
                }
            } else if (fachRaw) {
                subjectId = `subject:${normCode(fachRaw.split(/\s+/)[0] || fachRaw)}`;
            }

            if (teacherId && classId) {
                addEdge({ kind: 'teaches', source: teacherId, target: classId, label: row.gruppe || undefined });
            }
            if (teacherId && subjectId) {
                addEdge({ kind: 'teaches', source: teacherId, target: subjectId });
            }
            if (classId && subjectId) {
                addEdge({ kind: 'class_subject', source: classId, target: subjectId });
            }
        });
    }

    if (peopleLayer) {
        const guardianById = new Map();
        guardians.forEach((g) => {
            if (!g || !g.id) return;
            guardianById.set(g.id, g);
        });

        let studentCount = 0;
        students.forEach((s) => {
            if (studentCount >= options.maxStudents) return;
            const klasse = normClassKey(s.klasse);
            if (klasseFilter && klasse !== klasseFilter) return;
            if (!klasse && !s.name && !s.email) return;

            const sid = String(s.id || [klasse, s.email || s.name].join('|')).trim();
            if (!sid) return;
            const id = sid.startsWith('student:') ? sid : `student:${sid}`;
            studentCount += 1;
            addNode({
                id,
                kind: 'student',
                label: String(s.name || s.email || 'Schüler:in').trim(),
                sublabel: klasse || undefined,
                meta: { klasse, email: normEmail(s.email) || undefined }
            });
            if (klasse) {
                const classId = `class:${klasse}`;
                if (!nodes.has(classId)) {
                    addNode({ id: classId, kind: 'class', label: klasse, sublabel: klasse });
                }
                addEdge({ kind: 'in_class', source: id, target: classId });
            }

            let gCount = 0;
            (Array.isArray(s.guardianIds) ? s.guardianIds : []).forEach((gid) => {
                if (!asp.eltern) return;
                if (gCount >= 4) return;
                const g = guardianById.get(gid);
                if (!g) return;
                const gidNode = `guardian:${g.id}`;
                if (!nodes.has(gidNode)) {
                    if (nodes.size > options.maxGuardians + 500) return;
                    addNode({
                        id: gidNode,
                        kind: 'guardian',
                        label: String(g.name || g.email || 'Eltern').trim(),
                        sublabel: normEmail(g.email) || undefined
                    });
                }
                addEdge({ kind: 'guardian_of', source: gidNode, target: id });
                gCount += 1;
            });
        });
        if (studentCount >= options.maxStudents) {
            warnings.push(`Anzeige auf ${options.maxStudents} Schüler:innen begrenzt – Klasse filtern für Details.`);
        }
    }

    if (asp.microsoft365 && catalogLinks.length) {
        catalogLinks.forEach((link, idx) => {
            const kind = String(link.kind || '').trim();
            const code = normCode(link.code);
            if (!code && kind !== 'schueler' && kind !== 'lehrer' && kind !== 'verwaltung') return;
            const gId = String(link.graphGroupId || link.mailNickname || idx).trim();
            if (!gId) return;
            const groupNodeId = `m365group:${gId}`;
            addNode({
                id: groupNodeId,
                kind: 'm365group',
                label: String(link.displayName || link.mailNickname || 'Gruppe').trim(),
                sublabel: String(link.mailNickname || '').trim() || undefined,
                meta: { catalogKind: kind, graphGroupId: link.graphGroupId || '' }
            });

            let targetId = '';
            if (kind === 'subject') targetId = `subject:${code}`;
            else if (kind === 'arge') targetId = `arge:${code}`;
            else if (kind === 'class') targetId = `class:${normClassKey(code)}`;
            else if (kind === 'cohort' || kind === 'eltern') targetId = showSchoolHub ? 'school:root' : '';

            if (targetId && nodes.has(targetId)) {
                addEdge({ kind: 'm365_link', source: targetId, target: groupNodeId });
            } else if (kind === 'schueler' || kind === 'lehrer' || kind === 'verwaltung') {
                if (showSchoolHub) {
                    addEdge({ kind: 'm365_link', source: 'school:root', target: groupNodeId });
                }
            }
        });
    }

    const nodeList = [...nodes.values()];
    const stats = {
        nodes: nodeList.length,
        edges: edges.length,
        byKind: nodeList.reduce((acc, n) => {
            acc[n.kind] = (acc[n.kind] || 0) + 1;
            return acc;
        }, /** @type {Record<string, number>} */ ({})),
        studentsTotal: students.length,
        peopleLayer,
        preset: options.preset,
        aspectsActive: options.aspects
    };

    return { nodes: nodeList, edges, stats, warnings };
}

/**
 * Nachbarschaft für Fokus / Highlight.
 * @param {string} nodeId
 * @param {GraphEdge[]} edges
 * @returns {Set<string>}
 */
export function neighborhood(nodeId, edges) {
    const set = new Set([nodeId]);
    (edges || []).forEach((e) => {
        if (e.source === nodeId) set.add(e.target);
        if (e.target === nodeId) set.add(e.source);
    });
    return set;
}

/**
 * @param {GraphNode[]} nodes
 * @param {GraphEdge[]} edges
 * @param {{ width?: number, height?: number, seed?: number }} [opts]
 * @returns {Map<string, { x: number, y: number }>}
 */
export function layoutClusterGraph(nodes, edges, opts) {
    const width = opts && opts.width ? opts.width : 960;
    const height = opts && opts.height ? opts.height : 640;
    const positions = new Map();

    const byKind = /** @type {Record<string, GraphNode[]>} */ ({});
    (nodes || []).forEach((n) => {
        if (!byKind[n.kind]) byKind[n.kind] = [];
        byKind[n.kind].push(n);
    });

    const kindOrder = ['school', 'class', 'teacher', 'subject', 'arge', 'student', 'guardian', 'm365group', 'entraUser'];
    const cx = width / 2;
    const cy = height / 2;
    let ring = 0;

    if (byKind.school && byKind.school.length) {
        byKind.school.forEach((n) => positions.set(n.id, { x: cx, y: cy }));
    }

    kindOrder.forEach((kind) => {
        if (kind === 'school') return;
        const list = byKind[kind] || [];
        if (!list.length) return;
        ring += 1;
        const radius = 90 + ring * 78;
        list.forEach((n, i) => {
            const angle = (i / list.length) * Math.PI * 2 - Math.PI / 2;
            positions.set(n.id, {
                x: cx + Math.cos(angle) * radius,
                y: cy + Math.sin(angle) * radius
            });
        });
    });

    // Leichte Feder-Anziehung entlang Kanten (few iterations, no deps)
    const ids = [...positions.keys()];
    const adj = new Map();
    (edges || []).forEach((e) => {
        if (!adj.has(e.source)) adj.set(e.source, []);
        if (!adj.has(e.target)) adj.set(e.target, []);
        adj.get(e.source).push(e.target);
        adj.get(e.target).push(e.source);
    });

    for (let iter = 0; iter < 40; iter += 1) {
        ids.forEach((id) => {
            const p = positions.get(id);
            if (!p) return;
            const neighbors = adj.get(id) || [];
            neighbors.forEach((nid) => {
                const q = positions.get(nid);
                if (!q) return;
                const dx = q.x - p.x;
                const dy = q.y - p.y;
                p.x += dx * 0.04;
                p.y += dy * 0.04;
            });
        });
        // Abstoßung
        for (let i = 0; i < ids.length; i += 1) {
            for (let j = i + 1; j < ids.length; j += 1) {
                const a = positions.get(ids[i]);
                const b = positions.get(ids[j]);
                if (!a || !b) continue;
                let dx = b.x - a.x;
                let dy = b.y - a.y;
                let dist = Math.hypot(dx, dy) || 1;
                if (dist < 48) {
                    const push = ((48 - dist) / dist) * 0.5;
                    dx *= push;
                    dy *= push;
                    a.x -= dx;
                    a.y -= dy;
                    b.x += dx;
                    b.y += dy;
                }
            }
        }
    }

    return positions;
}
