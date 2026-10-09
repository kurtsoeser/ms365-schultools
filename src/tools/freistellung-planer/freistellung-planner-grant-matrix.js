/**
 * Planer-Zugriff: Zeilen = Gruppe oder Person, Spalten = Planer-Rollen (Direktion / KV / Schüler).
 */
import { normalizePlannerUsers } from './freistellung-planer-direktion-users.js';

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** @typedef {'group'|'user'} FrPrincipalType */
/** @typedef {'admin'|'direktion'|'kv'|'schueler'} FrPlannerRoleKey */

/**
 * @typedef {{
 *   principalType: FrPrincipalType,
 *   groupId: string,
 *   groupLabel: string,
 *   mail: string,
 *   displayName: string,
 *   roles: { admin: boolean, direktion: boolean, kv: boolean, schueler: boolean }
 * }} FrPlannerGrantRow
 */

export const FR_PLANNER_GRANT_COLS = [
    { key: 'admin', label: 'Admin', hint: 'Planer-Administration (IT / Technik)' },
    { key: 'direktion', label: 'Direktion', hint: 'Gesamtübersicht, Berichte' },
    { key: 'kv', label: 'Klassenvorstand', hint: 'Eigene Klassen, KV-Ansicht' },
    { key: 'schueler', label: 'Schüler', hint: 'Eigene Anträge' }
];

function emptyRoles() {
    return { admin: false, direktion: false, kv: false, schueler: false };
}

function normalizeExtraGroups(raw) {
    const arr = Array.isArray(raw) ? raw : [];
    const out = [];
    const seen = new Set();
    arr.forEach(function (entry) {
        const o = entry && typeof entry === 'object' ? entry : {};
        const groupId = String(o.groupId || o.id || '').trim();
        if (!groupId || !GUID_RE.test(groupId)) return;
        const key = groupId.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push({
            groupId,
            groupLabel: String(o.groupLabel || o.label || o.group || o.name || '').trim()
        });
    });
    return out;
}

function normMail(v) {
    return String(v || '').trim().toLowerCase();
}

/**
 * @param {unknown} raw
 * @returns {FrPlannerGrantRow[]}
 */
export function normalizePlannerGrantRows(raw) {
    const list = Array.isArray(raw) ? raw : [];
    /** @type {FrPlannerGrantRow[]} */
    const out = [];
    list.forEach(function (item) {
        if (!item || typeof item !== 'object') return;
        const principalType = item.principalType === 'user' ? 'user' : 'group';
        const groupId = String(item.groupId || item.id || '').trim();
        const groupLabel = String(item.groupLabel || item.label || '').trim();
        const mail = normMail(item.mail);
        const displayName = String(item.displayName || item.name || '').trim();
        const r = item.roles && typeof item.roles === 'object' ? item.roles : {};
        const roles = {
            admin: !!r.admin,
            direktion: !!r.direktion,
            kv: !!r.kv,
            schueler: !!r.schueler
        };
        const hasRole = roles.admin || roles.direktion || roles.kv || roles.schueler;
        if (principalType === 'user') {
            if (!mail && !displayName) return;
            if (!hasRole) return;
            out.push({
                principalType: 'user',
                groupId: '',
                groupLabel: '',
                mail,
                displayName: displayName || mail,
                roles
            });
            return;
        }
        if (!groupId && !groupLabel) return;
        if (!hasRole) return;
        out.push({
            principalType: 'group',
            groupId,
            groupLabel,
            mail: '',
            displayName: '',
            roles
        });
    });
    return mergeDuplicateGrantRows(out);
}

/**
 * @param {FrPlannerGrantRow[]} rows
 */
export function mergeDuplicateGrantRows(rows) {
    /** @type {FrPlannerGrantRow[]} */
    const groups = [];
    /** @type {Map<string, FrPlannerGrantRow>} */
    const users = new Map();
    (rows || []).forEach(function (row) {
        if (row.principalType === 'user') {
            const key = normMail(row.mail) || String(row.displayName || '').toLowerCase();
            if (!key) return;
            const hit = users.get(key);
            if (!hit) {
                users.set(key, { ...row, roles: { ...row.roles } });
                return;
            }
            FR_PLANNER_GRANT_COLS.forEach(function (col) {
                if (row.roles[col.key]) hit.roles[col.key] = true;
            });
            if (!hit.displayName && row.displayName) hit.displayName = row.displayName;
            return;
        }
        const gid = groupIdKey(row.groupId, row.groupLabel);
        const hit = groups.find(function (g) {
            return groupIdKey(g.groupId, g.groupLabel) === gid;
        });
        if (!hit) {
            groups.push({ ...row, roles: { ...row.roles } });
            return;
        }
        FR_PLANNER_GRANT_COLS.forEach(function (col) {
            if (row.roles[col.key]) hit.roles[col.key] = true;
        });
        if (!hit.groupId && row.groupId) hit.groupId = row.groupId;
        if (!hit.groupLabel && row.groupLabel) hit.groupLabel = row.groupLabel;
    });
    return groups.concat([...users.values()]);
}

function groupIdKey(id, label) {
    const g = String(id || '').trim().toLowerCase();
    if (g && GUID_RE.test(g)) return 'id:' + g;
    return 'lbl:' + String(label || '').trim().toLowerCase();
}

function pushGroupRow(out, groupId, groupLabel, roleKey) {
    const roles = emptyRoles();
    roles[roleKey] = true;
    out.push({
        principalType: 'group',
        groupId: String(groupId || '').trim(),
        groupLabel: String(groupLabel || '').trim(),
        mail: '',
        displayName: '',
        roles
    });
}

function pushUserRow(out, user, roleKey) {
    const mail = normMail(user && user.mail);
    if (!mail) return;
    const roles = emptyRoles();
    roles[roleKey] = true;
    out.push({
        principalType: 'user',
        groupId: '',
        groupLabel: '',
        mail,
        displayName: String((user && user.displayName) || mail).trim(),
        roles
    });
}

/**
 * @param {Record<string, unknown>} cfg
 * @returns {FrPlannerGrantRow[]}
 */
export function migrateLegacyFreistellungPerms(cfg) {
    const c = cfg && typeof cfg === 'object' ? cfg : {};
    /** @type {FrPlannerGrantRow[]} */
    const out = [];
    if (c.groupDirektionId || c.groupDirektion) {
        pushGroupRow(out, c.groupDirektionId, c.groupDirektion, 'direktion');
    }
    normalizeExtraGroups(c.direktionGroups).forEach(function (g) {
        pushGroupRow(out, g.groupId, g.groupLabel, 'direktion');
    });
    normalizePlannerUsers(c.direktionUsers).forEach(function (u) {
        pushUserRow(out, u, 'direktion');
    });

    if (c.groupKvId || c.groupKv) {
        pushGroupRow(out, c.groupKvId, c.groupKv, 'kv');
    }
    normalizeExtraGroups(c.kvGroups).forEach(function (g) {
        pushGroupRow(out, g.groupId, g.groupLabel, 'kv');
    });
    normalizePlannerUsers(c.kvUsers).forEach(function (u) {
        pushUserRow(out, u, 'kv');
    });

    if (c.groupSchuelerId || c.groupSchueler) {
        pushGroupRow(out, c.groupSchuelerId, c.groupSchueler, 'schueler');
    }
    normalizeExtraGroups(c.schuelerGroups).forEach(function (g) {
        pushGroupRow(out, g.groupId, g.groupLabel, 'schueler');
    });
    normalizePlannerUsers(c.schuelerUsers).forEach(function (u) {
        pushUserRow(out, u, 'schueler');
    });

    return mergeDuplicateGrantRows(out);
}

/**
 * @param {FrPlannerGrantRow[]} grantRows
 */
export function compileGrantRowsToLegacyFields(grantRows) {
    const rows = normalizePlannerGrantRows(grantRows);
    /** @type {Record<FrPlannerRoleKey, { groupId: string, groupLabel: string }[]>} */
    const groupLists = { admin: [], direktion: [], kv: [], schueler: [] };
    /** @type {Record<FrPlannerRoleKey, ReturnType<typeof normalizePlannerUsers>>} */
    const userLists = { admin: [], direktion: [], kv: [], schueler: [] };

    rows.forEach(function (row) {
        FR_PLANNER_GRANT_COLS.forEach(function (col) {
            if (!row.roles[col.key]) return;
            if (row.principalType === 'user') {
                userLists[col.key] = normalizePlannerUsers([
                    ...userLists[col.key],
                    { mail: row.mail, displayName: row.displayName }
                ]);
                return;
            }
            groupLists[col.key].push({
                groupId: row.groupId,
                groupLabel: row.groupLabel
            });
        });
    });

    function splitPrimary(list) {
        const norm = [];
        const seen = new Set();
        (list || []).forEach(function (g) {
            const groupId = String(g.groupId || '').trim();
            const groupLabel = String(g.groupLabel || '').trim();
            const key = groupId ? groupId.toLowerCase() : 'lbl:' + groupLabel.toLowerCase();
            if (!groupId && !groupLabel) return;
            if (seen.has(key)) return;
            seen.add(key);
            norm.push({ groupId, groupLabel });
        });
        const first = norm[0] || { groupId: '', groupLabel: '' };
        const rest = normalizeExtraGroups(norm.slice(1));
        return { first, rest };
    }

    const adm = splitPrimary(groupLists.admin);
    const dir = splitPrimary(groupLists.direktion);
    const kv = splitPrimary(groupLists.kv);
    const sch = splitPrimary(groupLists.schueler);

    return {
        groupAdmin: adm.first.groupLabel,
        groupAdminId: adm.first.groupId,
        adminGroups: adm.rest,
        adminUsers: userLists.admin,
        groupDirektion: dir.first.groupLabel,
        groupDirektionId: dir.first.groupId,
        groupKv: kv.first.groupLabel,
        groupKvId: kv.first.groupId,
        groupSchueler: sch.first.groupLabel,
        groupSchuelerId: sch.first.groupId,
        direktionGroups: dir.rest,
        kvGroups: kv.rest,
        schuelerGroups: sch.rest,
        direktionUsers: userLists.direktion,
        kvUsers: userLists.kv,
        schuelerUsers: userLists.schueler,
        plannerGrantRows: rows
    };
}
