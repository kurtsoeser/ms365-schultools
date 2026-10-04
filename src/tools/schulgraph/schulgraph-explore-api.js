/**
 * M365-Graph: Nachbarschaft beim Knoten-Klick (Mitgliedschaften / Mitglieder).
 */
import {
    entraUserNodeId,
    graphGroupNodeId,
    resolveClassGraphGroupId,
    stammdatenMembersForClass
} from './schulgraph-explore-logic.js';

function graphApi() {
    const g = typeof window !== 'undefined' ? window.ms365GraphUnifiedGroups : null;
    if (!g || typeof g.getGraphToken !== 'function') {
        throw new Error('Graph-Modul nicht geladen. Seite neu laden und anmelden.');
    }
    return g;
}

function personLabel(u) {
    const g = graphApi();
    if (typeof g.personLabel === 'function') return g.personLabel(u);
    return String(u?.displayName || u?.mail || u?.userPrincipalName || 'Person').trim();
}

function normEmail(v) {
    return String(v ?? '')
        .trim()
        .toLowerCase();
}

/**
 * @param {object} node
 */
export function emailFromNode(node) {
    if (!node) return '';
    const meta = node.meta && typeof node.meta === 'object' ? node.meta : {};
    if (meta.email) return normEmail(meta.email);
    if (node.kind === 'teacher' || node.kind === 'student' || node.kind === 'guardian') {
        const sub = normEmail(node.sublabel);
        if (sub.indexOf('@') !== -1) return sub;
    }
    return '';
}

async function fetchUserGroups(token, userId) {
    const g = graphApi();
    const path =
        '/users/' +
        encodeURIComponent(userId) +
        '/memberOf/microsoft.graph.group?$select=id,displayName,mail,mailNickname,groupTypes,resourceProvisioningOptions&$top=999';
    if (typeof g.fetchAllPagesSimple === 'function') {
        return g.fetchAllPagesSimple(token, path, 800);
    }
    const data = await g.graphJson('GET', path, token, undefined);
    return Array.isArray(data.value) ? data.value : [];
}

async function resolveGroupIdForNode(node) {
    const meta = node.meta && typeof node.meta === 'object' ? node.meta : {};
    if (meta.graphGroupId) return String(meta.graphGroupId).trim();
    const raw = String(node.id || '').replace(/^m365group:/, '').trim();
    if (/^[0-9a-f-]{36}$/i.test(raw)) return raw;
    const g = graphApi();
    const token = await g.getGraphToken();
    const q = String(node.label || meta.mailNickname || raw || '').trim();
    if (!q) throw new Error('Gruppen-ID unbekannt – in Einrichtung verknüpfen.');
    if (typeof g.searchUnifiedGroups === 'function') {
        const hits = await g.searchUnifiedGroups(token, q, 8);
        const exact = (hits || []).find(
            (h) =>
                String(h.mailNickname || '').toLowerCase() === String(node.sublabel || '').toLowerCase() ||
                String(h.displayName || '').toLowerCase() === String(node.label || '').toLowerCase()
        );
        if (exact && exact.id) return exact.id;
        if (hits && hits[0] && hits[0].id) return hits[0].id;
    }
    throw new Error('Gruppe in M365 nicht gefunden.');
}

/**
 * @param {object} node
 * @param {object} ctx
 * @returns {Promise<{ nodes: object[], edges: object[], summary: string, warnings: string[] }>}
 */
export async function expandNodeNeighborhood(node, ctx) {
    const settings = ctx.settings || {};
    const yearBucket = ctx.yearBucket || {};
    const setup = ctx.setup || {};
    const catalogIds = ctx.catalogGroupIds || new Set();

    if (node.kind === 'm365group') {
        return expandGroupMembers(node);
    }
    if (node.kind === 'class') {
        return expandClass(node, settings, yearBucket, setup);
    }
    if (node.kind === 'teacher' || node.kind === 'student' || node.kind === 'guardian' || node.kind === 'entraUser') {
        return expandPerson(node, catalogIds);
    }

    return {
        nodes: [],
        edges: [],
        summary: 'Für diesen Knotentyp gibt es keine M365-Mitgliedschafts-Ansicht.',
        warnings: []
    };
}

async function expandPerson(node, catalogIds) {
    const g = graphApi();
    const token = await g.getGraphToken();
    let userId = '';
    if (node.kind === 'entraUser') {
        userId = String(node.id || '').replace(/^entraUser:/, '');
    } else {
        const email = emailFromNode(node);
        if (!email) {
            return {
                nodes: [],
                edges: [],
                summary: 'Keine E-Mail am Knoten – M365-Gruppen können nicht geladen werden.',
                warnings: ['Stammdaten um E-Mail ergänzen oder Person in Personen-Verwaltung prüfen.']
            };
        }
        const u = await g.resolveUserByEmail(token, email);
        if (!u || !u.id) {
            return {
                nodes: [],
                edges: [],
                summary: 'Benutzer in M365 nicht gefunden.',
                warnings: [email]
            };
        }
        userId = u.id;
    }

    const groups = await fetchUserGroups(token, userId);
    const sorted = (groups || []).slice().sort((a, b) => personLabel(a).localeCompare(personLabel(b), 'de'));
    const nodes = [];
    const edges = [];
    let schoolCount = 0;

    sorted.forEach((gr) => {
        if (!gr || !gr.id) return;
        const id = graphGroupNodeId(gr.id);
        const inCatalog = catalogIds.has(String(gr.id).toLowerCase());
        if (inCatalog) schoolCount += 1;
        nodes.push({
            id,
            kind: 'm365group',
            label: String(gr.displayName || gr.mailNickname || 'Gruppe').trim(),
            sublabel: String(gr.mailNickname || gr.mail || '').trim() || undefined,
            meta: { graphGroupId: gr.id, fromExplore: true, inCatalog }
        });
        edges.push({ kind: 'member_of', source: node.id, target: id, label: 'Mitglied' });
    });

    const summary =
        sorted.length === 0
            ? 'Keine Gruppenmitgliedschaften in M365.'
            : `${sorted.length} Gruppe(n)${schoolCount ? `, davon ${schoolCount} in Schul-Einrichtung` : ''}.`;

    return {
        nodes,
        edges,
        summary,
        warnings: sorted.length > 120 ? ['Sehr viele Gruppen – Darstellung im Graph kann unübersichtlich werden.'] : []
    };
}

async function expandGroupMembers(node) {
    const g = graphApi();
    const token = await g.getGraphToken();
    const groupId = await resolveGroupIdForNode(node);
    const res = await g.fetchGroupMembers(token, groupId);
    const items = res && Array.isArray(res.items) ? res.items : [];
    const groupNodeId =
        node.id && String(node.id).startsWith('m365group:') ? node.id : graphGroupNodeId(groupId);
    const nodes = [];
    const edges = [];

    items.forEach((u) => {
        if (!u || !u.id) return;
        const id = entraUserNodeId(u.id);
        nodes.push({
            id,
            kind: 'entraUser',
            label: personLabel(u),
            sublabel: normEmail(u.mail || u.userPrincipalName) || undefined,
            meta: { graphUserId: u.id, fromExplore: true }
        });
        edges.push({ kind: 'group_member', source: id, target: groupNodeId, label: 'Mitglied' });
    });

    let summary = `${items.length} Mitglied(er) in „${node.label}“.`;
    if (res && res.truncated) {
        summary += ' (Liste gekürzt – sehr große Gruppe.)';
    }

    return {
        nodes,
        edges,
        summary,
        warnings: res && res.truncated ? ['Nicht alle Mitglieder geladen (Graph-Limit).'] : []
    };
}

async function expandClass(node, settings, yearBucket, setup) {
    const code = String(node.meta?.code || node.sublabel || node.label || '')
        .trim()
        .toUpperCase();
    const groupId = resolveClassGraphGroupId(code, setup);
    const warnings = [];

    if (groupId) {
        const syntheticGroup = {
            id: graphGroupNodeId(groupId),
            kind: 'm365group',
            label: node.label || code,
            meta: { graphGroupId: groupId }
        };
        const fromGraph = await expandGroupMembers(syntheticGroup);
        fromGraph.summary = `Klasse ${code}: ${fromGraph.summary}`;
        return fromGraph;
    }

    const local = stammdatenMembersForClass(code, settings, yearBucket);
    if (!local.length) {
        return {
            nodes: [],
            edges: [],
            summary: `Klasse ${code}: keine Schüler:innen in Stammdaten und keine M365-Klassengruppe verknüpft.`,
            warnings: ['In Einrichtung catalogLinks (Klasse) oder Klassen-Teams zuordnen.']
        };
    }

    warnings.push('Keine M365-Klassengruppe – Mitglieder aus lokalen Stammdaten (nicht live aus Entra).');
    const nodes = [];
    const edges = [];
    local.forEach((row, idx) => {
        const id = row.email ? `student:local-${normEmail(row.email)}` : `student:local-${code}-${idx}`;
        nodes.push({
            id,
            kind: 'student',
            label: row.name,
            sublabel: row.email || code,
            meta: { email: row.email, klasse: code, fromStammdaten: true }
        });
        edges.push({ kind: 'in_class', source: id, target: node.id });
    });

    return {
        nodes,
        edges,
        summary: `Klasse ${code}: ${local.length} Schüler:in(nen) aus Stammdaten.`,
        warnings
    };
}

/**
 * @param {object} setup
 * @returns {Set<string>}
 */
export function catalogGroupIdSet(setup) {
    const set = new Set();
    const links = setup && Array.isArray(setup.catalogLinks) ? setup.catalogLinks : [];
    links.forEach((L) => {
        const id = String(L.graphGroupId || '').trim().toLowerCase();
        if (id) set.add(id);
    });
    return set;
}
