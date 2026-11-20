/**
 * Normalisierung von Entra-Gruppen-Einträgen für Planner-/Listen-Rechte.
 */

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * @param {unknown} raw
 * @returns {Array<{ groupId: string, groupLabel: string }>}
 */
export function normalizeEntraGroupEntries(raw) {
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
            groupLabel: String(o.groupLabel || o.label || o.group || o.name || o.displayName || '').trim()
        });
    });
    return out;
}

/**
 * Legacy-Einzelgruppe (id + label) in die Liste mergen (ohne Duplikate).
 * @param {unknown} groups
 * @param {string} [legacyId]
 * @param {string} [legacyLabel]
 */
export function mergeLegacyEntraGroup(groups, legacyId, legacyLabel) {
    const list = normalizeEntraGroupEntries(groups);
    const id = String(legacyId || '').trim();
    if (!id || !GUID_RE.test(id)) return list;
    const key = id.toLowerCase();
    if (list.some((g) => g.groupId.toLowerCase() === key)) return list;
    list.unshift({
        groupId: id,
        groupLabel: String(legacyLabel || '').trim()
    });
    return list;
}

/**
 * Kurzlabel für UI-Zusammenfassung (1–n Gruppen).
 * @param {Array<{ groupId?: string, groupLabel?: string }>} groups
 */
export function formatEntraGroupsSummary(groups) {
    const list = normalizeEntraGroupEntries(groups);
    if (!list.length) return '';
    if (list.length === 1) return list[0].groupLabel || list[0].groupId;
    const first = list[0].groupLabel || list[0].groupId;
    return first + ' +' + (list.length - 1) + ' Gruppe(n)';
}
