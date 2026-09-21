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

/**
 * Ob die IT-Bibliothek lokal als eingerichtet gilt (Drive-ID vorhanden).
 * @param {{ driveId?: string }|null|undefined} meta
 */
export function isItLibraryConfigured(meta) {
    return !!(meta && String(meta.driveId || '').trim());
}

/**
 * Quellen von ms365-tenant-settings-changed, die keinen Auto-Push auslösen sollen.
 * @param {string|undefined|null} sourceOrReason
 */
export function isAutoSyncIgnoredChangeSource(sourceOrReason) {
    const s = String(sourceOrReason || '')
        .trim()
        .toLowerCase();
    if (!s) return false;
    return (
        s === 'browser-backup-import' ||
        s === 'spo-auto-pull' ||
        s === 'spo-auto-push' ||
        s === 'render' ||
        s.indexOf('spo-auto-') === 0
    );
}

/**
 * Ob das SharePoint-Backup den lokalen Stand ersetzen soll (Session-Pull).
 * Bei localDirty gewinnt lokal (zuerst pushen, nicht überschreiben).
 *
 * @param {{
 *   remoteExists?: boolean,
 *   remoteLastModified?: string,
 *   remoteExportedAt?: string,
 *   localDirty?: boolean,
 *   localRemoteLastModified?: string,
 *   localRemoteExportedAt?: string
 * }} input
 */
export function shouldApplyRemoteBackup(input) {
    const i = input || {};
    if (!i.remoteExists) return { apply: false, reason: 'missing' };
    if (i.localDirty) return { apply: false, reason: 'local-dirty' };
    const remoteLm = String(i.remoteLastModified || '').trim();
    const localLm = String(i.localRemoteLastModified || '').trim();
    if (remoteLm && localLm && remoteLm === localLm) {
        return { apply: false, reason: 'same-modified' };
    }
    const remoteEx = String(i.remoteExportedAt || '').trim();
    const localEx = String(i.localRemoteExportedAt || '').trim();
    if (remoteEx && localEx && remoteEx === localEx && remoteLm && localLm) {
        return { apply: false, reason: 'same-export' };
    }
    if (!localLm && !localEx) return { apply: true, reason: 'never-synced' };
    if (remoteLm && localLm && remoteLm > localLm) return { apply: true, reason: 'newer-modified' };
    if (remoteEx && localEx && remoteEx > localEx) return { apply: true, reason: 'newer-export' };
    if (remoteLm && !localLm) return { apply: true, reason: 'has-remote' };
    return { apply: false, reason: 'local-current' };
}

/**
 * Kurzer Sync-Status für die UI.
 * @param {{
 *   ready?: boolean,
 *   phase?: string,
 *   dirty?: boolean,
 *   lastAt?: string,
 *   lastDirection?: string,
 *   error?: string,
 *   libraryTitle?: string
 * }} state
 */
export function formatSyncStatusDe(state) {
    const s = state || {};
    if (!s.ready) {
        return (
            'SharePoint-IT-Bibliothek noch nicht eingerichtet – Sichern/Einlesen öffnet die Ersteinrichtung.'
        );
    }
    const bits = ['Bereit: „' + (s.libraryTitle || IT_LIBRARY_TITLE) + '“'];
    const phase = String(s.phase || 'idle');
    if (phase === 'pulling') bits.push('Lade von SharePoint …');
    else if (phase === 'pushing') bits.push('Sichere nach SharePoint …');
    else if (phase === 'error' && s.error) bits.push('Sync-Fehler: ' + s.error);
    else if (s.dirty) bits.push('Änderungen ausstehend (Auto-Sync)');
    else if (s.lastAt) {
        const when = String(s.lastAt).replace('T', ' ').replace(/\.\d+Z$/, '');
        const dir =
            s.lastDirection === 'pull'
                ? 'Zuletzt von SharePoint geladen'
                : s.lastDirection === 'push'
                  ? 'Zuletzt nach SharePoint gesichert'
                  : 'Zuletzt synchronisiert';
        bits.push(dir + ': ' + when);
    } else {
        bits.push('Noch kein Auto-Sync in dieser Sitzung');
    }
    return bits.join(' · ');
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
    buildItLibraryPlan,
    isItLibraryConfigured,
    isAutoSyncIgnoredChangeSource,
    shouldApplyRemoteBackup,
    formatSyncStatusDe
};
