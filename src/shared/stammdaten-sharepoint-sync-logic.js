/**
 * Pfade & Konstanten für Stammdaten-Übergabe via SharePoint Document Library.
 * Kein Graph, kein DOM – testbar.
 *
 * Design (v1):
 * - Eine JSON-Datei (Browser-Backup-Format) in der Standard-Dokumentbibliothek
 *   der Intranet-Site – NICHT als SharePoint-Listenzeilen (Schüler/Eltern/komplettes Setup).
 * - Öffentliche Ausschnitte bleiben Listen: Lehrerliste, Schultermine.
 */

export const DEFAULT_FOLDER = 'MS365-Schulverwaltung';
export const CURRENT_FILE = 'ms365-stammdaten-aktuell.json';

/**
 * Relativer Pfad unter drive/root (ohne führenden Slash).
 * @param {string} [folder]
 * @param {string} [fileName]
 */
export function buildDriveRelativePath(folder, fileName) {
    const rawFolder = folder == null ? DEFAULT_FOLDER : String(folder);
    const f = rawFolder
        .trim()
        .replace(/^\/+|\/+$/g, '')
        .replace(/\\/g, '/');
    const name = String(fileName || '')
        .trim()
        .replace(/^\/+/, '');
    if (!f) return name || CURRENT_FILE;
    if (!name) return f;
    return f + '/' + name;
}

/**
 * Graph-Pfadsegment: root:/path:/…
 * @param {string} relativePath
 */
export function encodeDriveRootPath(relativePath) {
    const rel = String(relativePath || '')
        .replace(/^\/+/, '')
        .split('/')
        .filter(Boolean)
        .map(function (seg) {
            return encodeURIComponent(seg);
        })
        .join('/');
    return 'root:/' + rel + ':';
}

/**
 * Kurzer Hinweistext für UI/Hilfe.
 */
export function designHintDe() {
    return (
        'Die Stammdaten liegen als eine JSON-Datei in der Dokumentbibliothek der Intranet-Site ' +
        '(Ordner „' +
        DEFAULT_FOLDER +
        '“). Schüler- und Elternlisten werden absichtlich nicht als SharePoint-Listenzeilen geschrieben – ' +
        'dafür sind Browser-Backup und diese Datei gedacht. Öffentliche Listen (Lehrer, Schultermine) bleiben separat.'
    );
}

/**
 * @param {{ schoolName?: string, domain?: string, exportedAt?: string, keyCount?: number, inventorySummary?: string }} meta
 */
export function describeRemoteBackup(meta) {
    const m = meta || {};
    const bits = [];
    if (m.schoolName) bits.push(String(m.schoolName));
    else if (m.domain) bits.push(String(m.domain));
    if (m.exportedAt) bits.push(String(m.exportedAt).replace('T', ' ').replace(/\.\d+Z$/, ' UTC'));
    if (m.keyCount != null) bits.push(String(m.keyCount) + ' Schlüssel');
    if (m.inventorySummary) bits.push(String(m.inventorySummary));
    return bits.join(' · ') || 'Backup auf SharePoint';
}

export default {
    DEFAULT_FOLDER,
    CURRENT_FILE,
    buildDriveRelativePath,
    encodeDriveRootPath,
    designHintDe,
    describeRemoteBackup
};
