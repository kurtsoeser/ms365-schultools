/**
 * Entra-Gruppen für Freistellungs-Planer-Rollen (Schüler, KV, Direktion).
 * Getrennt von Schularbeiten – gleiche Gruppen können eingetragen werden.
 */
import { overlayStammdatenAudienceOnPermissions } from '../../shared/school-admin-groups.js';
import {
    FR_STAMMDATEN_GROUP_ROLES,
    stripStammdatenGroupFieldsFromPatch
} from '../../shared/planner-stammdaten-audience-ui.js';
import {
    normalizeAllowedJahrgang,
    normalizeJahrgangGroups
} from './freistellung-planer-jahrgang-scope.js';
import { normalizePlannerUsers } from './freistellung-planer-direktion-users.js';
import {
    normalizePlannerGrantRows,
    migrateLegacyFreistellungPerms,
    compileGrantRowsToLegacyFields
} from './freistellung-planner-grant-matrix.js';

export const PERMS_STORAGE_KEY = 'ms365-freistellung-perms-v1';
export const SETUP_STORAGE_KEY = 'ms365-freistellung-setup-v1';

const DEFAULT_GROUPS = {
    groupAdmin: '',
    groupAdminId: '',
    groupDirektion: '',
    groupDirektionId: '',
    groupKv: '',
    groupKvId: '',
    groupSchueler: '',
    groupSchuelerId: ''
};

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * @param {unknown} raw
 * @returns {{ code: string, groupId: string, name: string }[]}
 */
/**
 * Klassenliste für Schüler-Dropdown (vom IT in den Planer veröffentlicht).
 * @param {unknown} raw
 * @returns {{ code: string, name: string }[]}
 */
export function normalizeClassCatalog(raw) {
    const arr = Array.isArray(raw) ? raw : [];
    const out = [];
    const seen = new Set();
    arr.forEach((entry) => {
        if (!entry || typeof entry !== 'object') return;
        const code = String(entry.code || entry.c || entry.name || '').trim();
        if (!code || seen.has(code.toLowerCase())) return;
        seen.add(code.toLowerCase());
        const headEmail = String(
            entry.headEmail || entry.klassenvorstandEmail || entry.kvEmail || entry.h || ''
        )
            .trim()
            .toLowerCase();
        const headName = String(entry.headName || entry.klassenvorstandName || entry.hn || '').trim();
        out.push({
            code,
            name: String(entry.name || entry.n || code).trim(),
            headEmail,
            headName
        });
    });
    return out;
}

export function normalizeClassTeamLinks(raw) {
    const arr = Array.isArray(raw) ? raw : [];
    const out = [];
    const seen = new Set();
    arr.forEach((entry) => {
        if (!entry || typeof entry !== 'object') return;
        const code = String(entry.code || entry.c || '').trim();
        const groupId = String(entry.groupId || entry.g || entry.graphGroupId || '').trim();
        if (!code || !GUID_RE.test(groupId)) return;
        const key = code.toLowerCase() + '|' + groupId.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push({
            code,
            groupId,
            name: String(entry.name || entry.n || code).trim()
        });
    });
    return out;
}

/**
 * Weitere Entra-Gruppen (Direktion / KV / Schüler), neben der Haupt-gruppe*Id.
 * @param {unknown} raw
 * @returns {Array<{ groupId: string, groupLabel: string }>}
 */
export function normalizePlannerExtraEntraGroups(raw) {
    const arr = Array.isArray(raw) ? raw : [];
    const out = [];
    const seen = new Set();
    arr.forEach((entry) => {
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

/** @deprecated Alias – gleiche Struktur wie direktionGroups / kvGroups */
export function normalizeSchuelerGroups(raw) {
    return normalizePlannerExtraEntraGroups(raw);
}

/**
 * @param {object|null|undefined} raw
 */
export function normalizePermissionsConfig(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    return {
        groupAdmin: String(r.groupAdmin || DEFAULT_GROUPS.groupAdmin).trim(),
        groupAdminId: String(r.groupAdminId || DEFAULT_GROUPS.groupAdminId).trim(),
        adminGroups: normalizePlannerExtraEntraGroups(r.adminGroups),
        adminUsers: normalizePlannerUsers(r.adminUsers),
        groupDirektion: String(r.groupDirektion || DEFAULT_GROUPS.groupDirektion).trim(),
        groupDirektionId: String(r.groupDirektionId || DEFAULT_GROUPS.groupDirektionId).trim(),
        groupKv: String(r.groupKv || DEFAULT_GROUPS.groupKv).trim(),
        groupKvId: String(r.groupKvId || DEFAULT_GROUPS.groupKvId).trim(),
        groupSchueler: String(r.groupSchueler || DEFAULT_GROUPS.groupSchueler).trim(),
        groupSchuelerId: String(r.groupSchuelerId || DEFAULT_GROUPS.groupSchuelerId).trim(),
        direktionGroups: normalizePlannerExtraEntraGroups(r.direktionGroups),
        kvGroups: normalizePlannerExtraEntraGroups(r.kvGroups),
        schuelerGroups: normalizePlannerExtraEntraGroups(r.schuelerGroups),
        direktionUsers: normalizePlannerUsers(r.direktionUsers),
        kvUsers: normalizePlannerUsers(r.kvUsers),
        schuelerUsers: normalizePlannerUsers(r.schuelerUsers),
        skipPerms: !!r.skipPerms,
        allowedJahrgang: normalizeAllowedJahrgang(r.allowedJahrgang),
        jahrgangGroups: normalizeJahrgangGroups(r.jahrgangGroups),
        classTeamLinks: normalizeClassTeamLinks(r.classTeamLinks),
        classCatalog: normalizeClassCatalog(r.classCatalog),
        plannerGrantRows: normalizePlannerGrantRows(r.plannerGrantRows)
    };
}

function resolvePlannerGrantRows(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    let rows = normalizePlannerGrantRows(r.plannerGrantRows);
    if (!rows.length) rows = migrateLegacyFreistellungPerms(r);
    return rows;
}

/** Zeilen für Setup-Matrix. */
export function plannerGrantRowsForUi(cfg) {
    const c = normalizePermissionsConfig(cfg || loadPermissionsConfig());
    const rows = resolvePlannerGrantRows(c);
    if (rows.length) return rows;
    return [];
}

function readSetupPlannerGroups() {
    try {
        const setup = JSON.parse(localStorage.getItem(SETUP_STORAGE_KEY) || '{}');
        if (setup && setup.plannerGroups && typeof setup.plannerGroups === 'object') {
            return normalizePermissionsConfig(setup.plannerGroups);
        }
    } catch {
        /* ignore */
    }
    return normalizePermissionsConfig({});
}

function mergePermissionLayers(base, overlay) {
    const b = normalizePermissionsConfig(base);
    const o = normalizePermissionsConfig(overlay);
    const out = { ...b };
    Object.keys(o).forEach((k) => {
        if (Array.isArray(o[k])) {
            if (o[k].length) out[k] = o[k];
            return;
        }
        if (String(o[k] || '').trim()) out[k] = o[k];
    });
    return normalizePermissionsConfig(out);
}

export function loadPermissionsConfig() {
    const fromSetup = readSetupPlannerGroups();
    try {
        const raw = JSON.parse(localStorage.getItem(PERMS_STORAGE_KEY) || '{}');
        return mergePermissionLayers(fromSetup, raw);
    } catch {
        return fromSetup;
    }
}

/** Schüler-Sammelgruppe aus Stammdaten, wenn gesetzt. */
export function loadEffectivePermissionsConfig() {
    return normalizePermissionsConfig(overlayStammdatenAudienceOnPermissions(loadPermissionsConfig()));
}

/**
 * Gruppen auch im Setup-JSON ablegen (Backup / gleicher Browser).
 * @param {ReturnType<typeof normalizePermissionsConfig>} config
 */
export function persistPlannerGroupsToSetup(config) {
    const groups = normalizePermissionsConfig(config);
    try {
        const setup = JSON.parse(localStorage.getItem(SETUP_STORAGE_KEY) || '{}') || {};
        setup.plannerGroups = groups;
        localStorage.setItem(SETUP_STORAGE_KEY, JSON.stringify(setup));
    } catch {
        /* ignore */
    }
    return groups;
}

/**
 * @param {object} patch
 */
export function savePermissionsConfig(patch) {
    const merged = { ...loadPermissionsConfig(), ...(patch || {}) };
    let rows =
        patch && patch.plannerGrantRows != null
            ? normalizePlannerGrantRows(patch.plannerGrantRows)
            : resolvePlannerGrantRows(merged);
    const compiled = compileGrantRowsToLegacyFields(rows);
    const mergedWithRows = {
        ...merged,
        ...compiled,
        plannerGrantRows: compiled.plannerGrantRows
    };
    const next = normalizePermissionsConfig(
        stripStammdatenGroupFieldsFromPatch(mergedWithRows, FR_STAMMDATEN_GROUP_ROLES)
    );
    try {
        localStorage.setItem(PERMS_STORAGE_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    persistPlannerGroupsToSetup(next);
    return next;
}

/**
 * @param {ReturnType<typeof normalizePermissionsConfig>|null|undefined} [config]
 */
export function entraGroupsConfigured(config) {
    const c = normalizePermissionsConfig(config || loadEffectivePermissionsConfig());
    return !!(
        c.groupAdminId ||
        c.groupDirektionId ||
        c.groupKvId ||
        c.groupSchuelerId ||
        (c.adminGroups && c.adminGroups.length) ||
        c.adminUsers.length ||
        (c.direktionGroups && c.direktionGroups.length) ||
        (c.kvGroups && c.kvGroups.length) ||
        (c.schuelerGroups && c.schuelerGroups.length) ||
        c.direktionUsers.length ||
        c.kvUsers.length ||
        c.schuelerUsers.length ||
        (c.plannerGrantRows && c.plannerGrantRows.length)
    );
}
