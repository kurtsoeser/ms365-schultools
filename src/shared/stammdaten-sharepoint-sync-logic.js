/**
 * Pfade & Konstanten für Stammdaten-Übergabe via SharePoint Document Library.
 * Kein Graph, kein DOM – testbar.
 *
 * Design (v2):
 * - Eigene IT-Dokumentbibliothek „MS365-IT-Stammdaten“ (nicht die öffentliche Standard-Bibliothek).
 * - Rechte: Vererbung brechen, Site-Besucher/Mitglieder entfernen, IT-Gruppe (Contribute).
 * - JSON-Backup-Datei darin – keine Schüler/Eltern als Listenzeilen.
 */

import { CONFIG_FOLDER, CONFIG_MANIFEST_FILE } from './stammdaten-sharepoint-config-bundle.js';

export const DEFAULT_FOLDER = 'Backups';
export const CURRENT_FILE = 'ms365-stammdaten-aktuell.json';
export { CONFIG_FOLDER, CONFIG_MANIFEST_FILE };
export const IT_LIBRARY_TITLE = 'MS365-IT-Stammdaten';
export const IT_LIBRARY_DESC =
    'Nur IT/Verwaltung: vollständiges Browser-Backup von MS365-Schul-Tools (Stammdaten, Automationen, Werkzeugstände). Nicht öffentlich.';

/** Entra-Objekt-ID (Gruppe/Benutzer). */
export const ENTRA_GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** SharePoint Standard-Rollen (RoleDefinitionId, SPO well-known). */
export const SPO_ROLE = {
    read: 1073741826,
    contribute: 1073741827,
    /** Gestaltung / Web Designer – u. a. „Listen verwalten“ (nötig bei Elementregel „nur eigene“). */
    design: 1073741828,
    fullControl: 1073741829,
    edit: 1073741830
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
 * Relativer Ordnerpfad für Drive-Listing (leer = Bibliotheks-Root).
 * @param {string} [folder]
 */
export function normalizeDriveListFolder(folder) {
    return String(folder == null ? '' : folder)
        .trim()
        .replace(/^\/+|\/+$/g, '')
        .replace(/\\/g, '/');
}

/**
 * @param {Array<{ name?: string, folder?: object, file?: object }>} items
 */
export function sortDriveBrowserItems(items) {
    return sortDriveBrowserItemsByColumn(items, 'name', 1);
}

/**
 * @param {Array<{ name?: string, size?: number, lastModifiedDateTime?: string, folder?: object, file?: object }>} items
 * @param {'name'|'modified'|'size'} key
 * @param {1|-1} dir
 */
/**
 * Ob „Backup übernehmen“ in der Bibliotheks-Ansicht angeboten werden soll.
 * @param {string} fileName
 * @param {{ folder?: string }} [opts] relativer Ordner in der IT-Bibliothek
 */
export function isLikelyImportableBackupFileName(fileName, opts) {
    const name = String(fileName || '').trim();
    if (!/\.json$/i.test(name)) return false;
    const folder = normalizeDriveListFolder(opts && opts.folder);
    if (folder === CONFIG_FOLDER || folder.indexOf(CONFIG_FOLDER + '/') === 0) {
        return false;
    }
    if (folder === DEFAULT_FOLDER || folder.indexOf(DEFAULT_FOLDER + '/') === 0) {
        return true;
    }
    const lower = name.toLowerCase();
    if (lower === CURRENT_FILE.toLowerCase()) return true;
    if (/^ms365-stammdaten-.*\.json$/i.test(name)) return true;
    if (/^ms365-browser-backup-.*\.json$/i.test(name)) return true;
    return false;
}

export function sortDriveBrowserItemsByColumn(items, key, dir) {
    const list = (items || []).filter(Boolean);
    const d = dir === -1 ? -1 : 1;
    const sortKey = key === 'modified' || key === 'size' ? key : 'name';
    list.sort(function (a, b) {
        const af = !!(a && a.folder);
        const bf = !!(b && b.folder);
        if (af !== bf) return af ? -1 : 1;
        let cmp = 0;
        if (sortKey === 'name') {
            cmp = String((a && a.name) || '').localeCompare(String((b && b.name) || ''), 'de', {
                sensitivity: 'base'
            });
        } else if (sortKey === 'modified') {
            const at = Date.parse(String((a && a.lastModifiedDateTime) || '')) || 0;
            const bt = Date.parse(String((b && b.lastModifiedDateTime) || '')) || 0;
            cmp = at - bt;
        } else {
            const as = af ? -1 : Number((a && a.size) || 0);
            const bs = bf ? -1 : Number((b && b.size) || 0);
            cmp = as - bs;
        }
        return cmp * d;
    });
    return list;
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
 * Stichpunkte für UI (ausklappbar) und als Grundlage für designHintDe().
 * @returns {string[]}
 */
export function designHintBulletsDe() {
    return [
        'Vollständiges Browser-Backup als JSON in „' +
            IT_LIBRARY_TITLE +
            '“ (Ordner „' +
            DEFAULT_FOLDER +
            '“, Datei ' +
            CURRENT_FILE +
            ').',
        'Zusätzlich aufgeteilte Schul-Konfiguration unter „config/“ (manifest.json, Stammdaten, Dashboard, Berechtigungsdateien inkl. optionaler Jahrgangs-Gruppen).',
        'Enthalten: Stammdaten, Dashboard-Zugriff, Power Automate/Freistellung und weitere Werkzeugstände im Monolithen.',
        'Rechte: Vererbung gebrochen – nur Site-Besitzer und die gewählte IT-/Verwaltungsgruppe.',
        'Schüler-/Elternlisten sind bewusst keine SharePoint-Listenzeilen.'
    ];
}

/** Einzeiler für die Hero-Karte (textarm). */
export function designHintSummaryDe() {
    return (
        'Gesamtes Schul- und App-Backup in eine private IT-Dokumentbibliothek – getrennt von der öffentlichen Intranet-Bibliothek. ' +
        'Speicherort: „' +
        IT_LIBRARY_TITLE +
        '“, Ordner „' +
        DEFAULT_FOLDER +
        '“.'
    );
}

/**
 * Fließtext für Protokoll/Hilfe (aus Stichpunkten zusammengesetzt).
 */
export function designHintDe() {
    return designHintBulletsDe().join(' ');
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
 * Link für „Bibliothek im Browser öffnen“ – Bibliotheks-URL vor Datei-URL
 * (Graph-webUrl einer .json-Datei führt oft zum Download statt zur Bibliothek).
 * @param {{ webUrl?: string }|null|undefined} itMeta
 * @param {{ webUrl?: string }|null|undefined} syncMeta
 */
export function resolveItLibraryBrowserHref(itMeta, syncMeta) {
    const lib = itMeta && itMeta.webUrl ? String(itMeta.webUrl).trim() : '';
    if (lib) return lib;
    const file = syncMeta && syncMeta.webUrl ? String(syncMeta.webUrl).trim() : '';
    return file;
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

/** Graph-Sitesuche: Begriffe, um die IT-Bibliothek im Tenant zu finden. */
export const IT_LIBRARY_SITE_SEARCH_TERMS = [
    IT_LIBRARY_TITLE,
    'MS365-IT-Stammdaten',
    'MS365-Schultools',
    'Schultools',
    'MS365 Schule'
];

/**
 * Niedriger = besser geeignet als Kandidat für die IT-Sicherungsbibliothek.
 * @param {string} displayName
 * @param {string} webUrl
 */
export function scoreItLibraryDiscoverySite(displayName, webUrl) {
    const n = String(displayName || '').toLowerCase();
    const u = String(webUrl || '').toLowerCase();
    if (n.indexOf('it-stammdaten') >= 0 || u.indexOf('it-stammdaten') >= 0) return 0;
    if (n.indexOf('schultools') >= 0 || u.indexOf('schultools') >= 0) return 1;
    if (n.indexOf('ms365') >= 0 || u.indexOf('ms365') >= 0) return 2;
    if (n.indexOf('intranet') >= 0 || u.indexOf('intranet') >= 0) return 3;
    if (n.indexOf('it') >= 0 || /\/sites\/it(\b|$)/i.test(u)) return 4;
    return 5;
}

/**
 * Eindeutige Site-URLs für Auto-Link (Reihenfolge = Priorität).
 * @param {...string} urls
 * @returns {string[]}
 */
export function uniqueItLibrarySiteUrls() {
    const seen = new Set();
    const out = [];
    for (let i = 0; i < arguments.length; i++) {
        const u = String(arguments[i] || '')
            .trim()
            .replace(/\/$/, '');
        if (!u) continue;
        const key = u.toLowerCase();
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(u);
    }
    return out;
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

    const siteUrls = uniqueItLibrarySiteUrls(
        it.siteUrl,
        draft.siteUrl,
        setup.intranetSiteUrl,
        setup.schoolIntranetSiteUrl
    );
    const siteUrl = siteUrls[0] || '';
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
        if (matched.schulleitungGroupId) itGroupId = String(matched.schulleitungGroupId).trim();
        else if (matched.verwaltungGroupId) itGroupId = String(matched.verwaltungGroupId).trim();
    }

    return {
        siteUrl: siteUrl,
        siteUrls: siteUrls,
        listTitle: listTitle,
        itGroupId: itGroupId || '',
        itGroupMail: itGroupMail || '',
        hasMinimum: !!(siteUrl && listTitle)
    };
}

/**
 * Kurzer deutscher Hinweis für Auto-Link-Skip-Codes (Toast/Status).
 * @param {string} skipped
 */
export function formatItLibraryLinkSkipDe(skipped) {
    const s = String(skipped || '').trim();
    if (s === 'no-hints') {
        return 'Keine Site-URL lokal bekannt – Intranet/Hub setzen oder Bibliothek suchen.';
    }
    if (s === 'library-not-found') {
        return 'Bibliothek „' + IT_LIBRARY_TITLE + '“ auf der Site nicht gefunden.';
    }
    if (s === 'site-unresolved' || s === 'site-missing') {
        return 'SharePoint-Site konnte nicht aufgelöst werden.';
    }
    if (s === 'no-token') {
        return 'SharePoint-Rechte fehlen (Sites.ReadWrite.All) oder Zustimmung ausstehend.';
    }
    if (s === 'drive-unreadable' || s === 'drive-missing') {
        return 'Dokumentbibliothek nicht lesbar (Rechte prüfen).';
    }
    if (s === 'list-search-failed' || s === 'list-id-missing') {
        return 'Liste/Bibliothek konnte nicht gelesen werden.';
    }
    if (!s || s === 'already-configured') return '';
    return 'Verknüpfung nicht möglich (' + s + ').';
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

/** Wichtige localStorage-Schlüssel für den Abgleich-Stichproben-Bericht. */
export const BACKUP_COVERAGE_SPOTLIGHT_KEYS = [
    'ms365-schooltool-data-v2',
    'ms365-dashboard-tool-access-v1',
    'ms365-dashboard-audience-groups-v1',
    'ms365-freistellung-setup-v1',
    'ms365-freistellung-perms-v1',
    'ms365-lfr-setup-v1',
    'ms365-lfr-perms-v1',
    'ms365-schularbeiten-perms-v1',
    'ms365-sa-settings-v1',
    'ms365-stammdaten-listen-perms-v1',
    'ms365-schueler-lehrer-gruppen-v2',
    'ms365-schulstruktur-sync-v1',
    'ms365-schulstruktur-match-v1',
    'webuntis-teams-creator-state-v1',
    'ms365-power-automate-recipes-v1'
];

function localStorageSliceFromBackupPayload(payload) {
    const p = payload && typeof payload === 'object' ? payload : {};
    const store = p.localStorage && typeof p.localStorage === 'object' ? p.localStorage : {};
    return store;
}

function parseSchooltoolV2FromStore(store) {
    const raw = store && store['ms365-schooltool-data-v2'];
    if (raw == null || raw === '') return null;
    if (typeof raw === 'object' && !Array.isArray(raw)) return raw;
    try {
        return JSON.parse(String(raw));
    } catch {
        return null;
    }
}

function currentYearBucket(container) {
    const c = container && typeof container === 'object' ? container : null;
    if (!c || !c.years) return null;
    const cur = String(c.years.current || '').trim();
    if (!cur || !c.years.byLabel || !c.years.byLabel[cur]) return null;
    return c.years.byLabel[cur];
}

/**
 * Kurzstatistik aus ms365-schooltool-data-v2 (für Abgleich-UI).
 * @param {object|null|undefined} payload Browser-Backup
 */
export function summarizeSchooltoolFromBackup(payload) {
    const store = localStorageSliceFromBackupPayload(payload);
    const v2 = parseSchooltoolV2FromStore(store);
    if (!v2) {
        return { present: false };
    }
    const core = v2.core && typeof v2.core === 'object' ? v2.core : {};
    const bucket = currentYearBucket(v2);
    const setup = v2.setup && typeof v2.setup === 'object' ? v2.setup : {};
    const matched = setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
    const sp = core.schoolProfile && typeof core.schoolProfile === 'object' ? core.schoolProfile : {};
    return {
        present: true,
        schoolName: String(core.schoolName || '').trim(),
        domain: String(core.domain || '').trim(),
        teacherCount: Array.isArray(core.teachers) ? core.teachers.length : 0,
        subjectCount: Array.isArray(core.subjects) ? core.subjects.length : 0,
        studentCount: bucket && Array.isArray(bucket.students) ? bucket.students.length : 0,
        classCount: bucket && Array.isArray(bucket.classes) ? bucket.classes.length : 0,
        verwaltungAudienceGroupCount: Array.isArray(core.verwaltungAudienceGroups)
            ? core.verwaltungAudienceGroups.length
            : 0,
        adminAudienceMembershipCount: Array.isArray(core.adminAudienceMemberships)
            ? core.adminAudienceMemberships.length
            : 0,
        hasSchoolProfileLogo: !!String(sp.logoDataUrl || '').trim(),
        schuelerGroupLinked: !!String(matched.schuelerGroupId || '').trim(),
        lehrerGroupLinked: !!String(matched.lehrerGroupId || '').trim(),
        schoolYear: String(v2.years && v2.years.current ? v2.years.current : '').trim()
    };
}

/**
 * @param {object|null|undefined} localPayload
 * @param {object|null|undefined} remotePayload
 */
export function compareSchooltoolSummaries(localPayload, remotePayload) {
    const local = summarizeSchooltoolFromBackup(localPayload);
    const remote = summarizeSchooltoolFromBackup(remotePayload);
    /** @type {Array<{ id: string, label: string, local: string, remote: string, match: boolean }>} */
    const rows = [];
    function row(id, label, lv, rv) {
        const ls = lv == null || lv === '' ? '–' : String(lv);
        const rs = rv == null || rv === '' ? '–' : String(rv);
        rows.push({
            id: id,
            label: label,
            local: ls,
            remote: rs,
            match: ls === rs
        });
    }
    if (!local.present && !remote.present) {
        return rows;
    }
    row('schoolName', 'Schulname', local.schoolName, remote.schoolName);
    row('domain', 'Domain', local.domain, remote.domain);
    row('schoolYear', 'Aktuelles Schuljahr', local.schoolYear, remote.schoolYear);
    row('teachers', 'Lehrer (Anzahl)', local.teacherCount, remote.teacherCount);
    row('subjects', 'Fächer (Anzahl)', local.subjectCount, remote.subjectCount);
    row('students', 'Schüler akt. SJ', local.studentCount, remote.studentCount);
    row('classes', 'Klassen akt. SJ', local.classCount, remote.classCount);
    row(
        'verwaltungAudience',
        'Verwaltungs-Zielgruppen',
        local.verwaltungAudienceGroupCount,
        remote.verwaltungAudienceGroupCount
    );
    row(
        'adminAudienceMemberships',
        'Zielgruppen-Mitgliedschaften',
        local.adminAudienceMembershipCount,
        remote.adminAudienceMembershipCount
    );
    row('logo', 'Schullogo gesetzt', local.hasSchoolProfileLogo ? 'ja' : 'nein', remote.hasSchoolProfileLogo ? 'ja' : 'nein');
    row(
        'slgSchueler',
        'SLG Schüler verknüpft',
        local.schuelerGroupLinked ? 'ja' : 'nein',
        remote.schuelerGroupLinked ? 'ja' : 'nein'
    );
    row(
        'slgLehrer',
        'SLG Lehrer verknüpft',
        local.lehrerGroupLinked ? 'ja' : 'nein',
        remote.lehrerGroupLinked ? 'ja' : 'nein'
    );
    return rows;
}

/**
 * Stichprobe: wichtige Schlüssel lokal vs. SharePoint-Backup.
 * @param {object|null|undefined} localPayload
 * @param {object|null|undefined} remotePayload
 * @param {ReturnType<typeof compareBackupPayloads>} [cmp]
 */
export function buildBackupCoverageReport(localPayload, remotePayload, cmp) {
    const comparison =
        cmp && typeof cmp === 'object'
            ? cmp
            : compareBackupPayloads(localPayload, remotePayload);
    const locStore = localStorageSliceFromBackupPayload(localPayload);
    const remStore = localStorageSliceFromBackupPayload(remotePayload);

    const spotlight = BACKUP_COVERAGE_SPOTLIGHT_KEYS.map(function (key) {
        const hasLocal = Object.prototype.hasOwnProperty.call(locStore, key);
        const hasRemote = Object.prototype.hasOwnProperty.call(remStore, key);
        let status = 'missing';
        if (hasLocal && hasRemote) {
            status =
                stableStorageValue(locStore[key]) === stableStorageValue(remStore[key]) ? 'same' : 'differs';
        } else if (hasLocal) status = 'local-only';
        else if (hasRemote) status = 'remote-only';
        return {
            key: key,
            label: labelForStorageKey(key),
            status: status,
            hasLocal: hasLocal,
            hasRemote: hasRemote
        };
    });

    const schooltoolRows = compareSchooltoolSummaries(localPayload, remotePayload);
    const mismatchSpotlight = spotlight.filter(function (s) {
        return s.status !== 'same' && s.status !== 'missing';
    }).length;
    const mismatchSchooltool = schooltoolRows.filter(function (r) {
        return !r.match;
    }).length;

    return {
        identical: !!comparison.identical,
        spotlight: spotlight,
        schooltoolRows: schooltoolRows,
        mismatchSpotlight: mismatchSpotlight,
        mismatchSchooltool: mismatchSchooltool,
        readyToSync: comparison.identical === true
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
    IT_LIBRARY_SITE_SEARCH_TERMS,
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
    uniqueItLibrarySiteUrls,
    formatItLibraryLinkSkipDe,
    scoreItLibraryDiscoverySite,
    ENTRA_GUID_RE,
    isLikelyFreshLocalBackup,
    compareBackupPayloads,
    BACKUP_COVERAGE_SPOTLIGHT_KEYS,
    summarizeSchooltoolFromBackup,
    compareSchooltoolSummaries,
    buildBackupCoverageReport,
    formatBackupCompareDe,
    storageKeyLabel,
    isAutoSyncIgnoredChangeSource,
    shouldApplyRemoteBackup,
    formatSyncStatusDe
};
