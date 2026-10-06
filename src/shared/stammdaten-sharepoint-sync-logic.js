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
export { CONFIG_FOLDER, CONFIG_MANIFEST_FILE } from './stammdaten-sharepoint-config-bundle.js';
export const IT_LIBRARY_TITLE = 'MS365-IT-Stammdaten';
export const IT_LIBRARY_DESC =
    'Nur IT/Verwaltung: vollständiges Browser-Backup von MS365-Schul-Tools (Stammdaten, Automationen, Werkzeugstände). Nicht öffentlich.';

/** Entra-Objekt-ID (Gruppe/Benutzer). */
export const ENTRA_GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

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
        'Schul- und App-Daten liegen als vollständiges JSON-Backup in der eigenen IT-Bibliothek „' +
        IT_LIBRARY_TITLE +
        '“ (Ordner „' +
        DEFAULT_FOLDER +
        '“: ' +
        CURRENT_FILE +
        '). Zusätzlich werden Schul-Konfigurationen aufgeteilt unter „config/“ (manifest.json, Stammdaten, Dashboard, permissions-schularbeiten.json, permissions-freistellung.json inkl. optionaler Jahrgangs-Gruppen). ' +
        'Enthalten: Stammdaten, Dashboard-Zugriff, Power-Automate-/Freistellungs-Konfiguration und weitere Werkzeugstände im Monolithen. ' +
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
 * IT-Bibliothek-Metadaten normalisieren (localStorage / Setup).
 * @param {object|null|undefined} raw
 */
export function normalizeItLibraryMeta(raw) {
    if (!raw || typeof raw !== 'object') return {};
    return {
        listTitle: raw.listTitle ? String(raw.listTitle).trim() : '',
        listId: raw.listId ? String(raw.listId).trim() : '',
        driveId: raw.driveId ? String(raw.driveId).trim() : '',
        webUrl: raw.webUrl ? String(raw.webUrl).trim() : '',
        itGroupId: raw.itGroupId ? String(raw.itGroupId).trim() : '',
        itGroupMail: raw.itGroupMail ? String(raw.itGroupMail).trim() : '',
        securedAt: raw.securedAt ? String(raw.securedAt) : null,
        siteUrl: raw.siteUrl ? String(raw.siteUrl).trim() : '',
        linkedAt: raw.linkedAt ? String(raw.linkedAt) : null,
        autoLinked: raw.autoLinked === true
    };
}

/**
 * Hinweise für Auto-Verknüpfung (Site, Bibliothek, IT-Gruppe) aus lokalen Quellen.
 * @param {{ itMeta?: object, setup?: object, formDraft?: object }} input
 */
export function collectItLibraryLinkHints(input) {
    const src = input || {};
    const it = normalizeItLibraryMeta(src.itMeta);
    const setup = src.setup && typeof src.setup === 'object' ? src.setup : {};
    const draft = src.formDraft && typeof src.formDraft === 'object' ? src.formDraft : {};

    const siteUrl = String(it.siteUrl || draft.siteUrl || setup.intranetSiteUrl || '').trim();
    const listTitle =
        String(it.listTitle || draft.libraryTitle || IT_LIBRARY_TITLE).trim() || IT_LIBRARY_TITLE;

    let itGroupId = it.itGroupId;
    let itGroupMail = it.itGroupMail;
    const draftGroup = String(draft.itGroup || '').trim();
    if (!itGroupId && !itGroupMail && draftGroup) {
        if (ENTRA_GUID_RE.test(draftGroup)) itGroupId = draftGroup;
        else itGroupMail = draftGroup;
    }
    if (!itGroupId && !itGroupMail) {
        const matched = setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
        if (matched.verwaltungGroupId) itGroupId = String(matched.verwaltungGroupId).trim();
    }

    return {
        siteUrl: siteUrl,
        listTitle: listTitle,
        itGroupId: itGroupId || '',
        itGroupMail: itGroupMail || '',
        hasMinimum: !!(siteUrl && listTitle)
    };
}

/** Lesbare Namen für häufige Backup-Schlüssel (Vergleichsdialog). */
const STORAGE_KEY_LABELS = {
    'ms365-schooltool-data-v2': 'Zentrale Schuldaten',
    'ms365-tenant-settings-v1': 'Stammdaten-Einstellungen',
    'ms365-stammdaten-it-library-v1': 'IT-Sicherungsbibliothek (Verknüpfung)',
    'ms365-stammdaten-it-library-by-tenant-v2': 'IT-Bibliothek pro Mandant',
    'ms365-stammdaten-spo-sync-v1': 'SharePoint-Sync-Status',
    'ms365-stammdaten-spo-sync-by-tenant-v2': 'SharePoint-Sync pro Mandant',
    'ms365-stammdaten-listen-perms-v1': 'SharePoint-Listen-Berechtigungen',
    'ms365-demo-mode-v1': 'Legacy (Demo-Modus, entfernt)',
    'ms365-dashboard-favorites-v1': 'Dashboard-Favoriten',
    'ms365-dashboard-recent-tools-v1': 'Dashboard zuletzt verwendet',
    'ms365-dashboard-audience-groups-v1': 'Dashboard Entra-Gruppen (Personas)',
    'ms365-dashboard-tool-access-v1': 'Dashboard Werkzeug-Zugriff (Matrix)',
    'ms365-dashboard-order-catalog-v1': 'Dashboard Katalog-Reihenfolge',
    'ms365-dash-catalog-fold-v1': 'Dashboard Katalog aufgeklappt',
    'webuntis-teams-creator-state-v1': 'Kursteams / WebUntis',
    'ms365-freistellung-setup-v1': 'Freistellungen Setup',
    'ms365-freistellung-setup-step-v1': 'Freistellungen Setup (Schritt)',
    'ms365-freistellung-perms-v1': 'Freistellungs-Planer Berechtigungen',
    'ms365-freistellung-kategorien-extra-v1': 'Freistellung Zusatz-Kategorien',
    'ms365-power-automate-recipes-v1': 'Power Automate Rezepte',
    'ms365-pa-termine-sync-v1': 'PA Schultermine',
    'ms365-pa-antraege-v1': 'PA Anträge',
    'ms365-pa-schularbeiten-mail-v1': 'PA Schularbeiten-Mail',
    'ms365-pa-projektwochen-mail-v1': 'PA Projektwochen-Mail',
    'ms365-sa-settings-v1': 'Schularbeiten-Planer Einstellungen',
    'ms365-schularbeiten-perms-v1': 'Schularbeiten-Planer Berechtigungen',
    'ms365-akt-planer-site-v1': 'Schulaktivitäten-Planer Site',
    'ms365-playbook-eltern-v1': 'Playbook Eltern',
    'ms365-playbook-elternsprechtag-v1': 'Playbook Elternsprechtag',
    'ms365-playbook-freistellungen-v1': 'Playbook Freistellungen',
    'ms365-playbook-intranet-v1': 'Playbook Intranet',
    'ms365-playbook-kursteams-v1': 'Playbook Kursteams',
    'ms365-playbook-schularbeiten-v1': 'Playbook Schularbeiten',
    'ms365-playbook-schuljahresstart-v1': 'Playbook Schuljahresstart',
    'ms365-cleanup-playbook-v1': 'Cleanup-Playbook',
    'ms365-elternsprechtag-bookings-v1': 'Elternsprechtag Buchungen',
    'ms365-hygiene-scan-v2': 'Datenhygiene-Scan',
    'ms365-student-lifecycle-prev-v1': 'Schüler-Lifecycle Snapshot'
};

function labelForStorageKey(key) {
    const k = String(key || '').trim();
    return STORAGE_KEY_LABELS[k] || k;
}

/** Lesbarer Name für einen localStorage-Schlüssel im Abgleich-UI. */
export function storageKeyLabel(key) {
    return labelForStorageKey(key);
}

function stableStorageValue(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string') return v;
    try {
        return JSON.stringify(v);
    } catch {
        return String(v);
    }
}

function parseIsoMs(iso) {
    const s = String(iso || '').trim();
    if (!s) return 0;
    const t = Date.parse(s);
    return Number.isFinite(t) ? t : 0;
}

function formatWhenDe(iso) {
    const s = String(iso || '').trim();
    if (!s) return '–';
    return s.replace('T', ' ').replace(/\.\d{3}Z$/, ' UTC').slice(0, 22);
}

/**
 * @param {object|null|undefined} payload Browser-Backup (buildBackup / SharePoint-JSON)
 */
export function isLikelyFreshLocalBackup(payload) {
    const loc = payload && payload.localStorage;
    if (!loc || typeof loc !== 'object') return true;
    const keys = Object.keys(loc).filter(function (k) {
        return k.indexOf('ms365-') === 0 || k.indexOf('webuntis-') === 0;
    });
    const raw = loc['ms365-schooltool-data-v2'];
    if (raw != null && raw !== '') {
        try {
            const data = typeof raw === 'object' ? raw : JSON.parse(String(raw));
            const years = data.years && typeof data.years === 'object' ? Object.keys(data.years) : [];
            if (years.length) return false;
        } catch {
            /* weiter unten */
        }
    }
    if (keys.length <= 2) return true;
    if (raw == null || raw === '') return true;
    return keys.length < 5;
}

/**
 * Vergleicht lokales und SharePoint-Backup (Inhalt/Fingerabdruck/Schlüssel).
 * @param {object|null|undefined} local
 * @param {object|null|undefined} remote
 */
export function compareBackupPayloads(local, remote) {
    const loc = local && typeof local === 'object' ? local : {};
    const rem = remote && typeof remote === 'object' ? remote : {};
    const locFp = String(loc.contentFingerprint || '').trim();
    const remFp = String(rem.contentFingerprint || '').trim();
    if (locFp && remFp && locFp === remFp) {
        return {
            identical: true,
            newerSide: 'equal',
            localExportedAt: loc.exportedAt || '',
            remoteExportedAt: rem.exportedAt || '',
            localSchool: String(loc.schoolName || loc.domain || '').trim(),
            remoteSchool: String(rem.schoolName || rem.domain || '').trim(),
            changedCount: 0,
            onlyLocalCount: 0,
            onlyRemoteCount: 0,
            changedKeys: [],
            onlyLocalKeys: [],
            onlyRemoteKeys: []
        };
    }

    const locAt = parseIsoMs(loc.exportedAt);
    const remAt = parseIsoMs(rem.exportedAt);
    let newerSide = 'unknown';
    if (locAt && remAt) {
        if (remAt > locAt) newerSide = 'remote';
        else if (locAt > remAt) newerSide = 'local';
        else newerSide = 'equal';
    } else if (remAt) newerSide = 'remote';
    else if (locAt) newerSide = 'local';

    const locStore = loc.localStorage && typeof loc.localStorage === 'object' ? loc.localStorage : {};
    const remStore = rem.localStorage && typeof rem.localStorage === 'object' ? rem.localStorage : {};
    const locKeys = Object.keys(locStore);
    const remKeys = Object.keys(remStore);
    const remKeySet = {};
    remKeys.forEach(function (k) {
        remKeySet[k] = true;
    });
    const onlyLocal = locKeys.filter(function (k) {
        return !remKeySet[k];
    });
    const onlyRemote = remKeys.filter(function (k) {
        return locKeys.indexOf(k) === -1;
    });
    const changed = locKeys.filter(function (k) {
        if (!remKeySet[k]) return false;
        return stableStorageValue(locStore[k]) !== stableStorageValue(remStore[k]);
    });

    return {
        identical: false,
        newerSide: newerSide,
        localExportedAt: loc.exportedAt || '',
        remoteExportedAt: rem.exportedAt || '',
        localSchool: String(loc.schoolName || loc.domain || '').trim(),
        remoteSchool: String(rem.schoolName || rem.domain || '').trim(),
        localKeyCount: locKeys.length,
        remoteKeyCount: remKeys.length,
        changedCount: changed.length,
        onlyLocalCount: onlyLocal.length,
        onlyRemoteCount: onlyRemote.length,
        changedKeys: changed.sort(),
        onlyLocalKeys: onlyLocal.sort(),
        onlyRemoteKeys: onlyRemote.sort()
    };
}

/**
 * Text für Bestätigungsdialog (Deutsch).
 * @param {ReturnType<typeof compareBackupPayloads>} cmp
 * @param {{ remoteLastModified?: string }} [ctx]
 */
export function formatBackupCompareDe(cmp, ctx) {
    const c = cmp || {};
    const lines = [];
    lines.push(
        'Lokal: ' +
            formatWhenDe(c.localExportedAt) +
            (c.localSchool ? ' · ' + c.localSchool : '') +
            (c.localKeyCount != null ? ' · ' + c.localKeyCount + ' Schlüssel' : '')
    );
    lines.push(
        'SharePoint: ' +
            formatWhenDe(c.remoteExportedAt) +
            (c.remoteSchool ? ' · ' + c.remoteSchool : '') +
            (c.remoteKeyCount != null ? ' · ' + c.remoteKeyCount + ' Schlüssel' : '')
    );
    if (ctx && ctx.remoteLastModified) {
        lines.push('Datei auf SharePoint geändert: ' + formatWhenDe(ctx.remoteLastModified));
    }
    if (c.identical) {
        lines.push('');
        lines.push('Inhalt ist identisch (gleicher Fingerabdruck).');
        return lines.join('\n');
    }
    lines.push('');
    if (c.newerSide === 'remote') lines.push('Einschätzung: SharePoint-Stand wirkt aktueller.');
    else if (c.newerSide === 'local') lines.push('Einschätzung: Lokaler Stand wirkt aktueller.');
    else lines.push('Einschätzung: Zeitstempel unklar – bitte Inhalt prüfen.');

    function listKeys(title, keys, total) {
        if (!total) return;
        lines.push('');
        lines.push(title + ' (' + total + '):');
        (keys || []).forEach(function (k) {
            lines.push('  • ' + labelForStorageKey(k));
        });
        const shown = keys || [];
        if (total > shown.length) {
            lines.push('  … und ' + (total - shown.length) + ' weitere');
        }
    }
    listKeys('Nur lokal', (c.onlyLocalKeys || []).slice(0, 15), c.onlyLocalCount);
    listKeys('Nur auf SharePoint', (c.onlyRemoteKeys || []).slice(0, 15), c.onlyRemoteCount);
    listKeys('Inhalt unterschiedlich', (c.changedKeys || []).slice(0, 25), c.changedCount);
    return lines.join('\n');
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
        s === 'spo-backup-import' ||
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
    normalizeItLibraryMeta,
    collectItLibraryLinkHints,
    ENTRA_GUID_RE,
    isLikelyFreshLocalBackup,
    compareBackupPayloads,
    formatBackupCompareDe,
    storageKeyLabel,
    isAutoSyncIgnoredChangeSource,
    shouldApplyRemoteBackup,
    formatSyncStatusDe
};
