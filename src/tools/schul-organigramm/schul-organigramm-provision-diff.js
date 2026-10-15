/**
 * S4/S5: Offene M365-Anlagen (Batch) + Diff Landkarte ↔ Listen (± Graph).
 * Spec: docs/schul-struktur-kanon.md
 */
import {
    NODE,
    isAreaNode,
    isBuiltinSchulOrganigrammArea,
    normalizeSchulOrganigramm,
    normalizeSchulOrgGroupKind,
    schulOrgGroupKindLabel,
    suggestSchulOrgMailNickname,
    setSchulOrganigrammAreaM365,
    syncSchulOrganigrammWithStammdaten
} from './schul-organigramm-logic.js';

function normStr(v) {
    return String(v ?? '').trim();
}

function normCode(v) {
    return normStr(v).toUpperCase();
}

function normEmail(v) {
    return normStr(v).toLowerCase();
}

/**
 * Ob eine Area für Batch-Provisioning in Frage kommt.
 * @param {object} node
 */
export function areaNeedsM365Provision(node) {
    if (!node || !isAreaNode(node)) return false;
    const m365 = node.m365 && typeof node.m365 === 'object' ? node.m365 : {};
    const mode = normStr(m365.groupMode) || 'none';
    const gid = normStr(m365.groupId);
    const kind = normalizeSchulOrgGroupKind(m365.groupKind || 'unified');
    // Teams-Gruppenchats werden nicht über die Entra-Gruppen-Provision angelegt
    if (kind === 'groupChat') return false;
    if (mode === 'create' && !gid) return true;
    return false;
}

/**
 * Areas ohne Group-Bindung (Kandidaten zum Vormerken).
 * Builtin-Sammelbereiche (Lehrer/Schüler) ausnehmen – die laufen über setup.matched.
 * @param {object} node
 */
export function areaIsUnboundProvisionCandidate(node) {
    if (!node || !isAreaNode(node)) return false;
    if (isBuiltinSchulOrganigrammArea(node)) return false;
    const refKind = normStr(node.refKind).toLowerCase();
    if (refKind === 'sammelgruppe' || refKind === 'jahrgang') return false;
    const m365 = node.m365 && typeof node.m365 === 'object' ? node.m365 : {};
    const mode = normStr(m365.groupMode) || 'none';
    const gid = normStr(m365.groupId);
    if (gid) return false;
    if (mode === 'match' || mode === 'create') return mode === 'create';
    // unbound custom / audience areas
    return true;
}

/**
 * @param {object} node
 * @returns {{
 *   nodeId: string,
 *   label: string,
 *   groupKind: string,
 *   groupKindLabel: string,
 *   displayName: string,
 *   mailNickname: string,
 *   reason: string
 * }}
 */
export function provisionJobFromArea(node) {
    const m365 = node && node.m365 && typeof node.m365 === 'object' ? node.m365 : {};
    const displayName = normStr(node && node.label) || 'Bereich';
    const kind = normalizeSchulOrgGroupKind(m365.groupKind || 'unified');
    const nick = suggestSchulOrgMailNickname(
        normStr(m365.groupMail) || normStr(m365.groupLabel) || displayName
    );
    return {
        nodeId: String(node.id),
        label: displayName,
        groupKind: kind,
        groupKindLabel: schulOrgGroupKindLabel(kind),
        displayName: displayName,
        mailNickname: nick,
        reason:
            normStr(m365.groupMode) === 'create'
                ? 'Als „Neu anlegen“ markiert'
                : 'Noch keine M365-Gruppe verknüpft'
    };
}

/**
 * Offene Batch-Jobs aus dem Layout.
 * @param {unknown} layout
 * @param {{ includeUnboundCandidates?: boolean }} [opts]
 */
export function listOpenSchulOrgProvisions(layout, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    const base = normalizeSchulOrganigramm(layout);
    /** @type {ReturnType<typeof provisionJobFromArea>[]} */
    const out = [];
    const seen = new Set();
    base.nodes.forEach(function (n) {
        if (!areaNeedsM365Provision(n) && !(o.includeUnboundCandidates && areaIsUnboundProvisionCandidate(n))) {
            return;
        }
        if (!areaNeedsM365Provision(n) && !areaIsUnboundProvisionCandidate(n)) return;
        if (seen.has(n.id)) return;
        seen.add(n.id);
        out.push(provisionJobFromArea(n));
    });
    out.sort(function (a, b) {
        return a.label.localeCompare(b.label, 'de');
    });
    return out;
}

/**
 * Areas mit groupMode=create vormerken (für Batch).
 * @param {unknown} layout
 * @param {string[]} areaIds
 * @param {string} [groupKind]
 */
export function markSchulOrgAreasForProvision(layout, areaIds, groupKind) {
    let base = normalizeSchulOrganigramm(layout);
    const ids = Array.isArray(areaIds) ? areaIds.map(String) : [];
    let kind = normalizeSchulOrgGroupKind(groupKind || 'unified');
    if (kind === 'groupChat') kind = 'unified';
    ids.forEach(function (id) {
        const node = base.nodes.find(function (n) {
            return n.id === id && isAreaNode(n);
        });
        if (!node) return;
        const prev = node.m365 && typeof node.m365 === 'object' ? node.m365 : {};
        let nextKind = normalizeSchulOrgGroupKind(prev.groupKind || kind);
        if (nextKind === 'groupChat') nextKind = kind;
        base = setSchulOrganigrammAreaM365(base, id, {
            groupMode: 'create',
            groupKind: nextKind,
            groupId: null,
            groupLabel: prev.groupLabel || node.label || null,
            groupMail: prev.groupMail || null,
            mailboxMode: prev.mailboxMode || 'none',
            mailboxUpn: prev.mailboxUpn || null,
            mailboxLabel: prev.mailboxLabel || null,
            mailboxId: prev.mailboxId || null
        });
    });
    return normalizeSchulOrganigramm(base);
}

/**
 * Struktur-Diff: Landkarte ↔ Stammdaten-Listen (ohne Graph).
 * @param {object} settings
 * @param {unknown} [layout]
 */
export function collectSchulOrgStructureDiff(settings, layout) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const base = normalizeSchulOrganigramm(layout != null ? layout : s.schulOrganigramm);
    const byRef = new Map();
    base.nodes.forEach(function (n) {
        const rk = normStr(n.refKind);
        const key = normStr(n.refKey);
        if (!rk || rk === 'none' || !key) return;
        byRef.set(rk + ':' + key.toUpperCase(), n);
    });

    /** @type {{ kind: string, key: string, label: string, detail: string }[]} */
    const missingInLayout = [];
    /** @type {{ kind: string, key: string, label: string, nodeId: string, detail: string }[]} */
    const orphanInLayout = [];

    (Array.isArray(s.adminRoles) ? s.adminRoles : []).forEach(function (r) {
        const code = normCode(r && r.code);
        if (!code) return;
        if (!byRef.has('adminRole:' + code)) {
            missingInLayout.push({
                kind: 'adminRole',
                key: code,
                label: normStr(r.name) || code,
                detail: 'Rolle in Stammdaten, fehlt auf der Landkarte'
            });
        }
    });
    (Array.isArray(s.arges) ? s.arges : []).forEach(function (a) {
        const code = normCode(a && a.code);
        if (!code) return;
        if (!byRef.has('arge:' + code)) {
            missingInLayout.push({
                kind: 'arge',
                key: code,
                label: normStr(a.name) || code,
                detail: 'ARGE in Stammdaten, fehlt auf der Landkarte'
            });
        }
    });
    (Array.isArray(s.classes) ? s.classes : []).forEach(function (c) {
        const code = normCode(c && c.code) || normCode(c && c.name);
        if (!code) return;
        if (!byRef.has('class:' + code)) {
            missingInLayout.push({
                kind: 'class',
                key: code,
                label: normStr(c.name) || code,
                detail: 'Klasse in Stammdaten, fehlt auf der Landkarte'
            });
        }
    });

    base.nodes.forEach(function (n) {
        const rk = normStr(n.refKind);
        const key = normStr(n.refKey);
        if (!rk || rk === 'none' || !key) return;
        if (rk === 'adminRole') {
            const hit = (Array.isArray(s.adminRoles) ? s.adminRoles : []).some(function (r) {
                return normCode(r && r.code) === normCode(key);
            });
            if (!hit) {
                orphanInLayout.push({
                    kind: rk,
                    key: key,
                    label: normStr(n.label) || key,
                    nodeId: n.id,
                    detail: 'Landkarte-Rolle ohne Stammdaten-Eintrag'
                });
            }
        } else if (rk === 'arge') {
            const hit = (Array.isArray(s.arges) ? s.arges : []).some(function (a) {
                return normCode(a && a.code) === normCode(key);
            });
            if (!hit) {
                orphanInLayout.push({
                    kind: rk,
                    key: key,
                    label: normStr(n.label) || key,
                    nodeId: n.id,
                    detail: 'Landkarte-ARGE ohne Stammdaten-Eintrag'
                });
            }
        } else if (rk === 'class') {
            const hit = (Array.isArray(s.classes) ? s.classes : []).some(function (c) {
                return (
                    normCode(c && c.code) === normCode(key) ||
                    normCode(c && c.name) === normCode(key)
                );
            });
            if (!hit) {
                orphanInLayout.push({
                    kind: rk,
                    key: key,
                    label: normStr(n.label) || key,
                    nodeId: n.id,
                    detail: 'Landkarte-Klasse ohne Stammdaten-Eintrag'
                });
            }
        }
    });

    const pendingCreate = listOpenSchulOrgProvisions(base, { includeUnboundCandidates: false });
    const unboundCandidates = listOpenSchulOrgProvisions(base, {
        includeUnboundCandidates: true
    }).filter(function (j) {
        return !pendingCreate.some(function (p) {
            return p.nodeId === j.nodeId;
        });
    });

    const boundAreas = base.nodes.filter(function (n) {
        if (!isAreaNode(n) || !n.m365) return false;
        return !!normStr(n.m365.groupId);
    });

    return {
        missingInLayout: missingInLayout,
        orphanInLayout: orphanInLayout,
        pendingCreate: pendingCreate,
        unboundCandidates: unboundCandidates,
        boundAreaCount: boundAreas.length,
        areaCount: base.nodes.filter(isAreaNode).length,
        counts: {
            missingInLayout: missingInLayout.length,
            orphanInLayout: orphanInLayout.length,
            pendingCreate: pendingCreate.length,
            unboundCandidates: unboundCandidates.length,
            boundAreas: boundAreas.length
        }
    };
}

/**
 * Layout mit Stammdaten synchronisieren (fehlende Knoten ergänzen).
 * @param {object} settings
 */
export function healSchulOrgStructureFromLists(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const result = syncSchulOrganigrammWithStammdaten(s.schulOrganigramm, s);
    return {
        layout: result.layout,
        added: result.added || 0,
        removed: result.removed || 0,
        healed: !!result.healed,
        labelsUpdated: result.labelsUpdated || 0
    };
}

/**
 * Eine Area in M365 anlegen und Layout aktualisieren.
 * @param {unknown} layout
 * @param {object} job provisionJobFromArea (+ optional memberEmails, ownerEmail)
 * @param {{
 *   token: string,
 *   createGroupByKind: (token: string, kind: string, name: string, nick: string, desc: string) => Promise<object>,
 *   ensureOwners?: (token: string, groupId: string, ownerEmails: string[]) => Promise<void>,
 *   syncEmailsToGroup?: (token: string, groupId: string, emails: string[], label?: string) => Promise<object>,
 *   sleep?: (ms: number) => Promise<void>
 * }} deps
 */
export async function provisionSchulOrgAreaJob(layout, job, deps) {
    const d = deps || {};
    if (!job || !job.nodeId) throw new Error('Provisioning-Job fehlt.');
    if (!d.token) throw new Error('Graph-Token fehlt.');
    if (typeof d.createGroupByKind !== 'function') {
        throw new Error('createGroupByKind nicht verfügbar.');
    }
    const created = await d.createGroupByKind(
        d.token,
        job.groupKind || 'unified',
        job.displayName || job.label,
        job.mailNickname,
        'Schul-Organigramm: ' + (job.displayName || job.label)
    );
    const gid = created && created.id ? String(created.id) : '';
    if (!gid) throw new Error('Gruppe angelegt, aber keine ID zurückgegeben.');

    const memberEmails = Array.isArray(job.memberEmails)
        ? job.memberEmails.map(function (e) {
              return normEmail(e);
          }).filter(Boolean)
        : [];
    const ownerEmail = normEmail(job.ownerEmail) || memberEmails[0] || '';
    if (typeof d.sleep === 'function') await d.sleep(1200);
    if (typeof d.ensureOwners === 'function') {
        await d.ensureOwners(d.token, gid, ownerEmail ? [ownerEmail] : []);
    }
    if (memberEmails.length && typeof d.syncEmailsToGroup === 'function') {
        await d.syncEmailsToGroup(
            d.token,
            gid,
            memberEmails,
            job.displayName || job.label || ''
        );
    }

    const prevNode = normalizeSchulOrganigramm(layout).nodes.find(function (n) {
        return n.id === job.nodeId;
    });
    const prevM365 = prevNode && prevNode.m365 && typeof prevNode.m365 === 'object' ? prevNode.m365 : {};
    const next = setSchulOrganigrammAreaM365(layout, job.nodeId, {
        groupMode: 'match',
        groupKind: job.groupKind || 'unified',
        groupId: gid,
        groupLabel: (created && created.displayName) || job.displayName || job.label,
        groupMail: normEmail((created && created.mail) || job.mailNickname),
        mailboxMode: prevM365.mailboxMode || 'none',
        mailboxUpn: prevM365.mailboxUpn || null,
        mailboxLabel: prevM365.mailboxLabel || null,
        mailboxId: prevM365.mailboxId || null
    });
    return {
        layout: next,
        groupId: gid,
        groupLabel: (created && created.displayName) || job.displayName || job.label,
        memberCount: memberEmails.length,
        ownerEmail: ownerEmail || null
    };
}

/**
 * Batch: Jobs nacheinander anlegen.
 * @param {unknown} layout
 * @param {ReturnType<typeof provisionJobFromArea>[]} jobs
 * @param {object} deps
 * @param {(ev: { index: number, total: number, job: object, ok: boolean, error?: string, groupId?: string }) => void} [onProgress]
 */
export async function runSchulOrgProvisionBatch(layout, jobs, deps, onProgress) {
    let current = normalizeSchulOrganigramm(layout);
    const list = Array.isArray(jobs) ? jobs : [];
    /** @type {{ nodeId: string, ok: boolean, error?: string, groupId?: string }[]} */
    const results = [];
    for (let i = 0; i < list.length; i++) {
        const job = list[i];
        try {
            const done = await provisionSchulOrgAreaJob(current, job, deps);
            current = done.layout;
            results.push({ nodeId: job.nodeId, ok: true, groupId: done.groupId });
            if (typeof onProgress === 'function') {
                onProgress({
                    index: i,
                    total: list.length,
                    job: job,
                    ok: true,
                    groupId: done.groupId
                });
            }
        } catch (e) {
            const msg = e && e.message ? String(e.message) : String(e);
            results.push({ nodeId: job.nodeId, ok: false, error: msg });
            if (typeof onProgress === 'function') {
                onProgress({
                    index: i,
                    total: list.length,
                    job: job,
                    ok: false,
                    error: msg
                });
            }
        }
    }
    const okCount = results.filter(function (r) {
        return r.ok;
    }).length;
    return {
        layout: current,
        results: results,
        okCount: okCount,
        failCount: results.length - okCount
    };
}

/**
 * Graph-Check: gebundene Areas gegen fetchGroup.
 * @param {unknown} layout
 * @param {{
 *   token: string,
 *   fetchGroup: (token: string, id: string) => Promise<object|null>
 * }} deps
 * @param {(ev: object) => void} [onProgress]
 */
export async function checkSchulOrgBoundGroups(layout, deps, onProgress) {
    const d = deps || {};
    const base = normalizeSchulOrganigramm(layout);
    if (!d.token || typeof d.fetchGroup !== 'function') {
        return { ok: [], missing: [], errors: [], skipped: true };
    }
    const bound = base.nodes.filter(function (n) {
        return isAreaNode(n) && n.m365 && normStr(n.m365.groupId);
    });
    /** @type {object[]} */
    const ok = [];
    /** @type {object[]} */
    const missing = [];
    /** @type {object[]} */
    const errors = [];
    for (let i = 0; i < bound.length; i++) {
        const n = bound[i];
        const gid = normStr(n.m365.groupId);
        try {
            const g = await d.fetchGroup(d.token, gid);
            if (g && g.id) {
                ok.push({
                    nodeId: n.id,
                    label: n.label,
                    groupId: gid,
                    groupLabel: g.displayName || n.m365.groupLabel || ''
                });
            } else {
                missing.push({
                    nodeId: n.id,
                    label: n.label,
                    groupId: gid,
                    detail: 'Gruppe in Entra nicht gefunden'
                });
            }
        } catch (e) {
            const msg = e && e.message ? String(e.message) : String(e);
            if (/404|not found|does not exist/i.test(msg)) {
                missing.push({
                    nodeId: n.id,
                    label: n.label,
                    groupId: gid,
                    detail: 'Gruppe fehlt in Entra'
                });
            } else {
                errors.push({
                    nodeId: n.id,
                    label: n.label,
                    groupId: gid,
                    detail: msg
                });
            }
        }
        if (typeof onProgress === 'function') {
            onProgress({ index: i, total: bound.length, nodeId: n.id });
        }
    }
    return { ok: ok, missing: missing, errors: errors, skipped: false };
}

export default {
    areaNeedsM365Provision,
    areaIsUnboundProvisionCandidate,
    provisionJobFromArea,
    listOpenSchulOrgProvisions,
    markSchulOrgAreasForProvision,
    collectSchulOrgStructureDiff,
    healSchulOrgStructureFromLists,
    provisionSchulOrgAreaJob,
    runSchulOrgProvisionBatch,
    checkSchulOrgBoundGroups
};
