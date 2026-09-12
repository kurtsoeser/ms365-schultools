/**
 * Pfade & Konstanten für Stammdaten-Übergabe via SharePoint Document Library.
 * Kein Graph, kein DOM – testbar.
 *
 * Design (v2):
 * - Eigene IT-Dokumentbibliothek „MS365-IT-Stammdaten“ (nicht die öffentliche Standard-Bibliothek).
 * - Rechte: Vererbung brechen, Site-Besucher/Mitglieder entfernen, IT-Gruppe (Contribute).
 * - JSON-Backup-Datei darin – keine Schüler/Eltern als Listenzeilen.
 */

export const DEFAULT_FOLDER = 'Backups';
export const CURRENT_FILE = 'ms365-stammdaten-aktuell.json';
export const IT_LIBRARY_TITLE = 'MS365-IT-Stammdaten';
export const IT_LIBRARY_DESC =
    'Nur IT/Verwaltung: Browser-Backup der MS365-Schulverwaltung (Stammdaten). Nicht öffentlich.';

/** SharePoint Standard-Rollen (RoleDefinitionId). */
export const SPO_ROLE = {
    read: 1073741826,
    contribute: 1073741827,
    edit: 1073741830,
    fullControl: 1073741829
};

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
        'Stammdaten liegen als JSON in der eigenen IT-Bibliothek „' +
        IT_LIBRARY_TITLE +
        '“ (nicht in „Dokumente“ für alle). ' +
        'Rechte: Vererbung gebrochen, nur Site-Besitzer + gewählte IT-/Verwaltungsgruppe. ' +
        'Schüler-/Elternlisten bleiben bewusst keine SharePoint-Listenzeilen.'
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

/**
 * Ob ein RoleAssignment-Mitglied „Besucher“ oder „Mitglieder“ der Site ist
 * (soll aus der IT-Bibliothek entfernt werden).
 * @param {{ Title?: string, LoginName?: string, PrincipalType?: number }} member
 */
export function isBroadSiteAudience(member) {
    const title = String((member && member.Title) || '').toLowerCase();
    const login = String((member && member.LoginName) || '').toLowerCase();
    if (/visitor|besucher|everyone except external|jeder außer/.test(title)) return true;
    if (/member|mitglieder/.test(title) && !/owner|besitzer/.test(title)) return true;
    if (login.indexOf('visitors') !== -1) return true;
    if (login.indexOf('members') !== -1 && login.indexOf('owners') === -1) return true;
    return false;
}

/**
 * LoginName für Entra-Gruppe in SharePoint (EnsureUser).
 * @param {string} groupObjectId
 */
export function entraGroupLogonName(groupObjectId) {
    const id = String(groupObjectId || '').trim();
    if (!id) return '';
    return 'c:0o.c|federateddirectoryclaimprovider|' + id;
}

/**
 * @param {{ listTitle?: string, itGroupId?: string, itGroupMail?: string }} input
 */
export function buildItLibraryPlan(input) {
    const listTitle = String((input && input.listTitle) || IT_LIBRARY_TITLE).trim() || IT_LIBRARY_TITLE;
    const itGroupId = String((input && input.itGroupId) || '').trim();
    const itGroupMail = String((input && input.itGroupMail) || '').trim();
    const issues = [];
    if (!itGroupId && !itGroupMail) issues.push('IT-/Verwaltungsgruppe fehlt (Objekt-ID oder E-Mail)');
    return {
        ok: issues.length === 0,
        issues,
        listTitle,
        description: IT_LIBRARY_DESC,
        itGroupId,
        itGroupMail,
        roleDefId: SPO_ROLE.contribute,
        folder: DEFAULT_FOLDER,
        currentFile: CURRENT_FILE
    };
}

export default {
    DEFAULT_FOLDER,
    CURRENT_FILE,
    IT_LIBRARY_TITLE,
    IT_LIBRARY_DESC,
    SPO_ROLE,
    buildDriveRelativePath,
    encodeDriveRootPath,
    designHintDe,
    describeRemoteBackup,
    isBroadSiteAudience,
    entraGroupLogonName,
    buildItLibraryPlan
};
