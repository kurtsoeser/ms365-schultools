/**
 * Datei-Migration: Ordnerlisten + Copy-Jobs (ohne Graph).
 */

function norm(s) {
    return String(s == null ? '' : s).trim();
}

/**
 * Drive-Item für UI-Zeile normalisieren.
 * @param {object} item Graph driveItem
 */
export function normalizeDriveItem(item) {
    const it = item || {};
    return {
        id: norm(it.id),
        name: norm(it.name) || '(ohne Namen)',
        size: Number(it.size) || 0,
        isFolder: !!it.folder,
        webUrl: norm(it.webUrl),
        lastModified: norm(it.lastModifiedDateTime),
        childCount: it.folder && it.folder.childCount != null ? Number(it.folder.childCount) : null
    };
}

/**
 * @param {array} items
 */
export function sortDriveItems(items) {
    return (Array.isArray(items) ? items.slice() : []).sort(function (a, b) {
        if (!!a.isFolder !== !!b.isFolder) return a.isFolder ? -1 : 1;
        return String(a.name || '').localeCompare(String(b.name || ''), 'de', { sensitivity: 'base' });
    });
}

/**
 * Copy-Body für Graph POST …/items/{id}/copy
 * @param {{ destDriveId: string, destFolderId?: string, newName?: string }} opts
 */
export function buildCopyBody(opts) {
    const o = opts || {};
    const destDriveId = norm(o.destDriveId);
    const destFolderId = norm(o.destFolderId) || 'root';
    if (!destDriveId) throw new Error('Ziel-Drive-ID fehlt.');
    const body = {
        parentReference: {
            driveId: destDriveId,
            id: destFolderId === 'root' ? 'root' : destFolderId
        }
    };
    if (norm(o.newName)) body.name = norm(o.newName);
    return body;
}

/**
 * Validiert Auswahl vor Migration.
 */
export function validateMigrationSelection(input) {
    const issues = [];
    const sourceGroupId = norm(input && input.sourceGroupId);
    const destGroupId = norm(input && input.destGroupId);
    const itemIds = Array.isArray(input && input.itemIds) ? input.itemIds.map(norm).filter(Boolean) : [];
    if (!sourceGroupId) issues.push('Quell-Team fehlt');
    if (!destGroupId) issues.push('Ziel-Team fehlt');
    if (sourceGroupId && destGroupId && sourceGroupId === destGroupId) {
        issues.push('Quelle und Ziel dürfen nicht gleich sein');
    }
    if (!itemIds.length) issues.push('Keine Dateien/Ordner ausgewählt');
    return {
        ok: issues.length === 0,
        issues,
        sourceGroupId,
        destGroupId,
        itemIds,
        note:
            'Chat und Aufgaben werden nicht kopiert – nur Dateien in der Team-Dokumentbibliothek. ' +
            'Große Ordner können länger dauern; Graph kopiert asynchron.'
    };
}

export default {
    normalizeDriveItem,
    sortDriveItems,
    buildCopyBody,
    validateMigrationSelection
};
