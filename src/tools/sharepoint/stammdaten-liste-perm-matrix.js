/**
 * Bearbeitbare Berechtigungs-Matrix für Stammdaten-Listen (Schritt 4).
 */
import { LIST_TYPE_KEYS, STAMMDATEN_LIST_PERM_PROFILES } from './stammdaten-liste-permissions.js';

/** @typedef {'read'|'contribute'|'fullControl'} PermLevel */

export const PERM_LEVEL_OPTIONS = [
    { value: 'fullControl', label: 'Vollzugriff' },
    { value: 'contribute', label: 'Beitragen' },
    { value: 'read', label: 'Lesen' }
];

/** Spalten in der UI (Zeilen = Entra-Gruppen). */
export const STAMMDATEN_PERM_MATRIX_COLS = [
    { key: 'schueler', label: 'Schülerinnen' },
    { key: 'faecher', label: 'Fächer' },
    { key: 'fachgruppen', label: 'Fachgruppen' },
    { key: 'arges', label: 'ARGEs' },
    { key: 'klassen', label: 'Klassen' },
    { key: 'lehrer', label: 'Lehrerinnen' }
];

/** @deprecated Nur Migration / Tests – altes Layout Listen×Rollen */
export const STAMMDATEN_PERM_MATRIX_ROWS = [
    { keys: ['schueler'], label: 'Schülerinnen' },
    { keys: ['faecher', 'fachgruppen'], label: 'Fächer / Fachgruppen' },
    { keys: ['arges'], label: 'ARGEs' },
    { keys: ['klassen'], label: 'Klassen' }
];

const ROLES = ['admin', 'lehrer', 'schueler'];

/**
 * @param {unknown} v
 * @param {PermLevel} fallback
 * @returns {PermLevel}
 */
function normalizeLevel(v, fallback) {
    if (v === 'read' || v === 'contribute' || v === 'fullControl') return v;
    return fallback;
}

/**
 * @param {unknown} v
 * @param {PermLevel|null} fallback
 * @returns {PermLevel|null}
 */
function normalizeNullableLevel(v, fallback) {
    if (v === null || v === undefined || v === '' || v === 'none' || v === '—') return null;
    return normalizeLevel(v, fallback || 'read');
}

/**
 * @param {Record<string, unknown>|null|undefined} raw
 * @returns {Record<string, { admin: PermLevel, lehrer: PermLevel|null, schueler: PermLevel|null }>}
 */
export function normalizeListPermProfiles(raw) {
    const out = {};
    const src = raw && typeof raw === 'object' ? raw : {};
    LIST_TYPE_KEYS.forEach(function (key) {
        const def = STAMMDATEN_LIST_PERM_PROFILES[key];
        if (!def) return;
        const p = src[key];
        if (p && typeof p === 'object') {
            out[key] = {
                admin: normalizeLevel(p.admin, def.admin),
                lehrer: normalizeNullableLevel(p.lehrer, def.lehrer),
                schueler: normalizeNullableLevel(p.schueler, def.schueler)
            };
        } else {
            out[key] = {
                admin: def.admin,
                lehrer: def.lehrer,
                schueler: def.schueler
            };
        }
    });
    return out;
}

/**
 * @param {Record<string, { admin: PermLevel, lehrer: PermLevel|null, schueler: PermLevel|null }>} profiles
 * @param {string[]} keys
 * @param {'admin'|'lehrer'|'schueler'} role
 * @param {PermLevel|null} level
 */
export function patchListPermProfiles(profiles, keys, role, level) {
    const base = normalizeListPermProfiles(profiles);
    const roleKeys = keys && keys.length ? keys : [];
    roleKeys.forEach(function (k) {
        if (!base[k]) return;
        const next = { ...base[k] };
        if (role === 'admin') {
            next.admin = level === null ? base[k].admin : normalizeLevel(level, base[k].admin);
        } else if (role === 'lehrer') {
            next.lehrer = level === null ? null : normalizeLevel(level, base[k].lehrer || 'read');
        } else if (role === 'schueler') {
            next.schueler = level === null ? null : normalizeLevel(level, base[k].schueler || 'read');
        }
        base[k] = next;
    });
    return base;
}

/**
 * @param {PermLevel|null} level
 */
export function permLevelLabel(level) {
    if (level === null) return '—';
    const hit = PERM_LEVEL_OPTIONS.find(function (o) {
        return o.value === level;
    });
    return hit ? hit.label : String(level);
}

export { ROLES as STAMMDATEN_PERM_MATRIX_ROLES };

/**
 * @typedef {{ groupId: string, groupLabel: string, cells: Record<string, PermLevel|null> }} StammdatenGrantRow
 */

/**
 * @param {unknown} v
 * @param {PermLevel|null} fallback
 * @returns {PermLevel|null}
 */
function normalizeCellLevel(v, fallback) {
    if (v === null || v === undefined || v === '' || v === 'none' || v === '—') return null;
    return normalizeLevel(v, fallback || 'read');
}

/**
 * @param {unknown} raw
 * @returns {StammdatenGrantRow[]}
 */
export function normalizeGrantRows(raw) {
    const list = Array.isArray(raw) ? raw : [];
    /** @type {StammdatenGrantRow[]} */
    const out = [];
    list.forEach(function (item) {
        if (!item || typeof item !== 'object') return;
        const groupId = String(item.groupId || item.id || '').trim();
        const groupLabel = String(item.groupLabel || item.label || '').trim();
        const cellsRaw = item.cells && typeof item.cells === 'object' ? item.cells : {};
        /** @type {Record<string, PermLevel|null>} */
        const cells = {};
        STAMMDATEN_PERM_MATRIX_COLS.forEach(function (col) {
            const def = STAMMDATEN_LIST_PERM_PROFILES[col.key];
            const fallback = def ? def.admin : 'read';
            cells[col.key] = normalizeCellLevel(cellsRaw[col.key], fallback);
        });
        if (!groupId && !groupLabel && !Object.values(cells).some(Boolean)) return;
        out.push({ groupId, groupLabel, cells });
    });
    return out;
}

/**
 * @param {'admin'|'lehrer'|'schueler'} roleKey
 * @returns {Record<string, PermLevel|null>}
 */
export function defaultCellsForAudienceRole(roleKey) {
    /** @type {Record<string, PermLevel|null>} */
    const cells = {};
    STAMMDATEN_PERM_MATRIX_COLS.forEach(function (col) {
        const def = STAMMDATEN_LIST_PERM_PROFILES[col.key];
        if (!def) {
            cells[col.key] = null;
            return;
        }
        const level = def[roleKey];
        cells[col.key] = level === undefined ? null : level;
    });
    return cells;
}

/**
 * @param {string} listKey
 * @param {StammdatenGrantRow[]} grantRows
 * @returns {{ groupId: string, groupLabel: string, level: PermLevel }[]}
 */
export function grantsForListKey(listKey, grantRows) {
    const key = String(listKey || '').trim();
    const rows = normalizeGrantRows(grantRows);
    /** @type {{ groupId: string, groupLabel: string, level: PermLevel }[]} */
    const out = [];
    rows.forEach(function (row) {
        const level = row.cells[key];
        if (!level) return;
        out.push({
            groupId: row.groupId,
            groupLabel: row.groupLabel,
            level
        });
    });
    return out;
}

/**
 * Altes Format (3 Gruppen + listProfiles) → grantRows.
 * @param {Record<string, unknown>} cfg
 */
export function migrateLegacyPermConfig(cfg) {
    const c = cfg && typeof cfg === 'object' ? cfg : {};
    const profiles = normalizeListPermProfiles(c.listProfiles);
    const slots = [
        { role: 'admin', groupId: c.groupAdminId, groupLabel: c.groupAdmin },
        { role: 'lehrer', groupId: c.groupLehrerId, groupLabel: c.groupLehrer },
        { role: 'schueler', groupId: c.groupSchuelerId, groupLabel: c.groupSchueler }
    ];
    /** @type {StammdatenGrantRow[]} */
    const rows = [];
    slots.forEach(function (slot) {
        const groupId = String(slot.groupId || '').trim();
        const groupLabel = String(slot.groupLabel || '').trim();
        if (!groupId && !groupLabel) return;
        /** @type {Record<string, PermLevel|null>} */
        const cells = {};
        STAMMDATEN_PERM_MATRIX_COLS.forEach(function (col) {
            const p = profiles[col.key];
            if (!p) {
                cells[col.key] = null;
                return;
            }
            const level = p[slot.role];
            cells[col.key] = level === undefined ? null : level;
        });
        rows.push({ groupId, groupLabel, cells });
    });
    return normalizeGrantRows(rows);
}

/**
 * Standard-Zeilen ohne Gruppe (Picker offen), Zellen wie frühere Rollen-Defaults.
 * @returns {StammdatenGrantRow[]}
 */
export function seedDefaultGrantRows() {
    return normalizeGrantRows([
        { groupId: '', groupLabel: '', cells: defaultCellsForAudienceRole('admin') },
        { groupId: '', groupLabel: '', cells: defaultCellsForAudienceRole('lehrer') },
        { groupId: '', groupLabel: '', cells: defaultCellsForAudienceRole('schueler') }
    ]);
}

/**
 * @param {StammdatenGrantRow[]} grantRows
 * @param {Record<string, unknown>} cfg mit groupAdminId …
 */
export function mergeAudienceSlotsIntoGrantRows(grantRows, cfg) {
    const base = normalizeGrantRows(grantRows);
    const c = cfg && typeof cfg === 'object' ? cfg : {};
    const slots = [
        { role: 'admin', groupId: c.groupAdminId, groupLabel: c.groupAdmin },
        { role: 'lehrer', groupId: c.groupLehrerId, groupLabel: c.groupLehrer },
        { role: 'schueler', groupId: c.groupSchuelerId, groupLabel: c.groupSchueler }
    ];
    slots.forEach(function (slot) {
        const gid = String(slot.groupId || '').trim();
        const glabel = String(slot.groupLabel || '').trim();
        if (!gid && !glabel) return;
        const hit = base.find(function (r) {
            return gid && r.groupId === gid;
        });
        if (hit) {
            if (glabel && !hit.groupLabel) hit.groupLabel = glabel;
            return;
        }
        base.push({
            groupId: gid,
            groupLabel: glabel,
            cells: defaultCellsForAudienceRole(slot.role)
        });
    });
    return normalizeGrantRows(base);
}
