/**
 * Entra-Gruppen für Freistellungs-Planer-Rollen (Schüler, KV, Direktion).
 * Getrennt von Schularbeiten – gleiche Gruppen können eingetragen werden.
 */
import { normalizePlannerUsers } from './freistellung-planer-direktion-users.js';

export const PERMS_STORAGE_KEY = 'ms365-freistellung-perms-v1';
export const SETUP_STORAGE_KEY = 'ms365-freistellung-setup-v1';

const DEFAULT_GROUPS = {
    groupDirektion: '',
    groupDirektionId: '',
    groupKv: '',
    groupKvId: '',
    groupSchueler: '',
    groupSchuelerId: ''
};

/**
 * @param {object|null|undefined} raw
 */
export function normalizePermissionsConfig(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    return {
        groupDirektion: String(r.groupDirektion || DEFAULT_GROUPS.groupDirektion).trim(),
        groupDirektionId: String(r.groupDirektionId || DEFAULT_GROUPS.groupDirektionId).trim(),
        groupKv: String(r.groupKv || DEFAULT_GROUPS.groupKv).trim(),
        groupKvId: String(r.groupKvId || DEFAULT_GROUPS.groupKvId).trim(),
        groupSchueler: String(r.groupSchueler || DEFAULT_GROUPS.groupSchueler).trim(),
        groupSchuelerId: String(r.groupSchuelerId || DEFAULT_GROUPS.groupSchuelerId).trim(),
        direktionUsers: normalizePlannerUsers(r.direktionUsers),
        kvUsers: normalizePlannerUsers(r.kvUsers),
        schuelerUsers: normalizePlannerUsers(r.schuelerUsers),
        skipPerms: !!r.skipPerms
    };
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
    const next = normalizePermissionsConfig({ ...loadPermissionsConfig(), ...(patch || {}) });
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
    const c = normalizePermissionsConfig(config || loadPermissionsConfig());
    return !!(
        c.groupDirektionId ||
        c.groupKvId ||
        c.groupSchuelerId ||
        c.direktionUsers.length ||
        c.kvUsers.length ||
        c.schuelerUsers.length
    );
}
