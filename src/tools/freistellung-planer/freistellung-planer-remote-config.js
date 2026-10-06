/**
 * Planer-Gruppen mandantenweit: JSON auf der SharePoint-Site (Site Assets oder Standard-Bibliothek).
 */
import {
    normalizePermissionsConfig,
    savePermissionsConfig,
    loadPermissionsConfig,
    loadEffectivePermissionsConfig,
    entraGroupsConfigured
} from './freistellung-planer-permissions.js';
import { overlaySchoolAudienceOnPermissions } from '../../shared/school-audience-groups.js';
import {
    loadExtraKategorien,
    normalizeExtraKategorien,
    saveExtraKategorien
} from './freistellung-planer-kategorien.js';
import { tryResolveFrListId, patchFreistellungKlasseColumn } from './freistellung-planer-graph.js';
import { LIST_TITLE_DEFAULT } from './freistellung-planer-schema.js';
import { persistSiteUrl, persistSetupFields } from './freistellung-planer-state.js';
import {
    classTeamLinksFromAppData,
    collectAllClassRows,
    classRowsFromTeamLinks
} from './freistellung-planer-class-context.js';
import {
    normalizeClassTeamLinks,
    normalizeClassCatalog
} from './freistellung-planer-permissions.js';
import { loadStammdaten } from './freistellung-planer-state.js';

/** Relativ zum Drive-Root (Site Assets oder „Dokumente“) – ohne Site-Namen im Pfad. */
export const REMOTE_CONFIG_REL_PATH = 'ms365/freistellung-planer-groups.json';

/** @deprecated Altpfad – wird beim Lesen noch versucht */
export const REMOTE_CONFIG_LEGACY_PATHS = ['MS365-Schultools/freistellung-planer-groups.json'];

export const REMOTE_CONFIG_VERSION = 1;

/** In der Listen-Beschreibung (lesbar für alle mit Listen-Zugriff). */
export const LIST_DESCRIPTION_MARKER = 'MS365_FR_GROUPS:';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

const SCOPES_READ = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
];

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

export function encodeDriveRootPathForUpload(relativePath) {
    return encodeDriveRootPath(relativePath);
}

function encodeDriveRootPath(relativePath) {
    const parts = String(relativePath || '')
        .replace(/^\/+/, '')
        .split('/')
        .filter(Boolean);
    if (!parts.length) return 'root:';
    return 'root:/' + parts.map((p) => encodeURIComponent(p)).join('/') + ':';
}

/**
 * @param {ReturnType<typeof normalizePermissionsConfig>} config
 */
export function permissionsToRemotePayload(config, opts) {
    const c = normalizePermissionsConfig(config);
    const extraKat = normalizeExtraKategorien(
        (opts && opts.kategorienExtra) || loadExtraKategorien()
    );
    const payload = {
        version: REMOTE_CONFIG_VERSION,
        updatedAt: new Date().toISOString(),
        groupDirektion: c.groupDirektion,
        groupDirektionId: c.groupDirektionId,
        groupKv: c.groupKv,
        groupKvId: c.groupKvId,
        groupSchueler: c.groupSchueler,
        groupSchuelerId: c.groupSchuelerId,
        direktionUsers: c.direktionUsers,
        kvUsers: c.kvUsers,
        schuelerUsers: c.schuelerUsers,
        allowedJahrgang: c.allowedJahrgang,
        jahrgangGroups: c.jahrgangGroups,
        classTeamLinks: c.classTeamLinks,
        classCatalog: c.classCatalog
    };
    if (extraKat.length) payload.kategorienExtra = extraKat;
    return payload;
}

/** Kompakt für Listen-Beschreibung (Zeichenlimit). */
export function permissionsToListDescriptionPayload(config, opts) {
    const c = normalizePermissionsConfig(config);
    const o = opts || {};
    const payload = {
        v: REMOTE_CONFIG_VERSION,
        gd: c.groupDirektionId,
        gk: c.groupKvId,
        gs: c.groupSchuelerId,
        du: compactUserList(c.direktionUsers),
        ku: compactUserList(c.kvUsers),
        su: compactUserList(c.schuelerUsers)
    };
    const w = String(o.siteWebUrl || '').trim().replace(/\/$/, '');
    const lid = String(o.listId || '').trim();
    if (w) payload.w = w;
    if (lid) payload.lid = lid;
    const links = normalizeClassTeamLinks(
        (opts && opts.classTeamLinks) || c.classTeamLinks || []
    );
    if (links.length) {
        payload.ct = links.slice(0, 120).map((row) => ({
            c: row.code,
            g: row.groupId,
            n: row.name && row.name !== row.code ? row.name : undefined
        }));
    }
    const catalog = normalizeClassCatalog((opts && opts.classCatalog) || c.classCatalog || []);
    if (catalog.length) {
        payload.cl = catalog.slice(0, 200).map((row) => ({
            c: row.code,
            n: row.name && row.name !== row.code ? row.name : undefined
        }));
    }
    return payload;
}

/** Stammdaten-Klassen (IT-Browser) für Veröffentlichung an Schüler. */
export function classCatalogFromSchoolStammdaten() {
    const st = loadStammdaten();
    const rows = collectAllClassRows({ stammdaten: { classes: st.classes || [] } });
    return normalizeClassCatalog(
        rows.map((r) => ({
            code: r.code || r.name,
            name: r.name || r.code
        }))
    );
}

/** Klassen-Teams aus App-Daten in Permissions mergen (für Listen-Beschreibung). */
export function mergeClassTeamLinksForPublish(config) {
    const c = normalizePermissionsConfig(config || loadEffectivePermissionsConfig());
    const merged = normalizeClassTeamLinks(c.classTeamLinks);
    const seen = new Set(merged.map((r) => r.code.toLowerCase() + '|' + r.groupId.toLowerCase()));
    classTeamLinksFromAppData().forEach((row) => {
        const key = row.code.toLowerCase() + '|' + row.groupId.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        merged.push({
            code: row.code,
            groupId: row.groupId,
            name: row.name || row.code
        });
    });
    return normalizePermissionsConfig({ ...c, classTeamLinks: merged });
}

/** Gruppen + Klassenliste + Team-IDs für SharePoint (Schüler ohne IT-Stammdaten lokal). */
export function mergePlannerPublishConfig(config) {
    const withTeams = mergeClassTeamLinksForPublish(config);
    const fromStamm = classCatalogFromSchoolStammdaten();
    // Stammdaten (1A, 1B, …) haben Vorrang vor alten Listenwerten (1AHW …).
    let catalog = fromStamm.length
        ? fromStamm.slice()
        : normalizeClassCatalog(withTeams.classCatalog);
    if (!fromStamm.length) {
        const seen = new Set(catalog.map((r) => r.code.toLowerCase()));
        classRowsFromTeamLinks(withTeams.classTeamLinks || []).forEach((row) => {
            const key = row.code.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            catalog.push({ code: row.code, name: row.name || row.code });
        });
    }
    catalog = catalog.filter((r) => !/^[1-5]AHW$/i.test(String(r.code || '')));
    return normalizePermissionsConfig({
        ...withTeams,
        classCatalog: catalog
    });
}

/**
 * @param {object} raw JSON aus Listen-Beschreibung (nach Marker)
 */
export function listDescriptionBootstrapMeta(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const siteUrl = String(raw.w || raw.siteUrl || '').trim().replace(/\/$/, '');
    const listId = String(raw.lid || raw.listId || '').trim();
    if (!siteUrl && !listId) return null;
    return { siteUrl, listId };
}

function compactUserList(users) {
    return (users || []).map((u) => ({
        id: u.id || '',
        m: u.mail || ''
    }));
}

function expandUserList(raw) {
    const arr = Array.isArray(raw) ? raw : [];
    return arr.map((u) => ({
        id: u.id || '',
        mail: u.m || u.mail || '',
        displayName: u.m || u.mail || ''
    }));
}

function expandListDescriptionPayload(raw) {
    if (!raw || typeof raw !== 'object') return null;
    let base = null;
    if (raw.groupSchuelerId || raw.groupKvId || raw.groupDirektionId) {
        base = remotePayloadToPermissions(raw);
    } else if (raw.gd || raw.gk || raw.gs) {
        base = normalizePermissionsConfig({
            groupDirektionId: raw.gd || '',
            groupKvId: raw.gk || '',
            groupSchuelerId: raw.gs || '',
            direktionUsers: expandUserList(raw.du),
            kvUsers: expandUserList(raw.ku),
            schuelerUsers: expandUserList(raw.su)
        });
    } else {
        return null;
    }
    const ct = expandClassTeamLinks(raw.ct);
    const cl = expandClassCatalog(raw.cl);
    return normalizePermissionsConfig({
        ...base,
        classTeamLinks: ct.length ? ct : base.classTeamLinks,
        classCatalog: cl.length ? cl : base.classCatalog
    });
}

function expandClassCatalog(raw) {
    const arr = Array.isArray(raw) ? raw : [];
    return arr.map((row) => ({
        code: row.c || row.code || '',
        name: row.n || row.name || row.c || row.code || ''
    }));
}

function expandClassTeamLinks(raw) {
    const arr = Array.isArray(raw) ? raw : [];
    return arr.map((row) => ({
        code: row.c || row.code || '',
        groupId: row.g || row.groupId || '',
        name: row.n || row.name || row.c || row.code || ''
    }));
}

/**
 * @param {object} raw
 */
export function remotePayloadToPermissions(raw) {
    if (!raw || typeof raw !== 'object') return normalizePermissionsConfig({});
    return normalizePermissionsConfig(raw);
}

async function resolveSite(tok, siteWebUrl) {
    const site = await G().resolveSiteFromWebUrl(tok, siteWebUrl);
    if (!site || !site.id) throw new Error('SharePoint-Site nicht gefunden.');
    return site;
}

/**
 * Site Assets (typisch für alle angemeldeten Site-Mitglieder lesbar), sonst Standard-drive.
 */
export async function resolveConfigDrive(tok, siteWebUrl) {
    const site = await resolveSite(tok, siteWebUrl);
    const siteId = site.id;
    try {
        const filter = encodeURIComponent("displayName eq 'Site Assets'");
        const lists = await G().graphJson(
            'GET',
            G().graphPathSite(siteId) + '/lists?$filter=' + filter + '&$select=id,displayName',
            tok,
            undefined,
            'v1.0'
        );
        const sa = (lists && lists.value && lists.value[0]) || null;
        if (sa && sa.id) {
            const drive = await G().graphJson(
                'GET',
                G().graphPathSite(siteId) +
                    '/lists/' +
                    encodeURIComponent(sa.id) +
                    '/drive?$select=id',
                tok,
                undefined,
                'v1.0'
            );
            if (drive && drive.id) {
                return { siteId, driveId: drive.id, driveLabel: 'Site Assets' };
            }
        }
    } catch {
        /* fallback */
    }
    const drive = await G().graphJson(
        'GET',
        G().graphPathSite(siteId) + '/drive?$select=id',
        tok,
        undefined,
        'v1.0'
    );
    if (!drive || !drive.id) throw new Error('Dokumentbibliothek der Site nicht gefunden.');
    return { siteId, driveId: drive.id, driveLabel: 'Standard-Bibliothek' };
}

async function fetchJsonFromDrive(driveId, relPath, tok) {
    const enc = encodeDriveRootPath(relPath);
    const getUrl = G().graphBase('v1.0') + '/drives/' + encodeURIComponent(driveId) + '/' + enc + '/content';
    const res = await fetch(getUrl, {
        method: 'GET',
        headers: { Authorization: 'Bearer ' + tok }
    });
    if (res.status === 404) return null;
    if (!res.ok) {
        const text = await res.text();
        throw new Error('Planer-Konfiguration von SharePoint: ' + (text || res.status));
    }
    return JSON.parse(await res.text());
}

/**
 * IT: Gruppen-JSON auf die Site legen (Schüler/KV lesen beim Planer-Start).
 * @param {string} siteWebUrl
 * @param {ReturnType<typeof normalizePermissionsConfig>} [config]
 */
/**
 * @param {string} siteWebUrl
 * @param {string} listId
 * @param {ReturnType<typeof normalizePermissionsConfig>} [config]
 */
export async function publishPlannerGroupsToList(siteWebUrl, listId, config) {
    const url = String(siteWebUrl || '').trim().replace(/\/$/, '');
    const id = String(listId || '').trim();
    if (!url || !id) return { ok: false, reason: 'no-list' };
    const cfg = mergePlannerPublishConfig(config || loadEffectivePermissionsConfig());
    if (!entraGroupsConfigured(cfg)) return { ok: false, reason: 'no-groups' };
    const tok = await G().getGraphToken(SCOPES);
    const site = await resolveSite(tok, url);
    const payload =
        LIST_DESCRIPTION_MARKER +
        JSON.stringify(
            permissionsToListDescriptionPayload(cfg, {
                siteWebUrl: url,
                listId: id,
                classTeamLinks: cfg.classTeamLinks,
                classCatalog: cfg.classCatalog
            })
        );
    await G().graphJson(
        'PATCH',
        G().graphPathSite(site.id) + '/lists/' + encodeURIComponent(id),
        tok,
        { description: payload },
        'v1.0'
    );
    return { ok: true, via: 'list-description' };
}

/**
 * @param {string} siteWebUrl
 * @param {string} listId
 */
function parseListDescriptionMarker(desc) {
    const full = parseListDescriptionMarkerFull(desc);
    return full && full.permissions ? full.permissions : null;
}

/**
 * @param {string} desc
 */
export function parseListDescriptionMarkerFull(desc) {
    const text = String(desc || '');
    if (!text.startsWith(LIST_DESCRIPTION_MARKER)) return null;
    try {
        const raw = JSON.parse(text.slice(LIST_DESCRIPTION_MARKER.length));
        const expanded = expandListDescriptionPayload(raw);
        const permissions = expanded || remotePayloadToPermissions(raw);
        return {
            permissions,
            meta: listDescriptionBootstrapMeta(raw),
            raw
        };
    } catch {
        return null;
    }
}

async function readListDescriptionText(siteWebUrl, listId) {
    const url = String(siteWebUrl || '').trim().replace(/\/$/, '');
    const id = String(listId || '').trim();
    if (!url || !id) return '';
    const readScopes = [
        ...SCOPES_READ,
        'https://graph.microsoft.com/Sites.ReadWrite.All'
    ];
    let desc = '';
    try {
        const tok = await G().getGraphToken(readScopes);
        const site = await resolveSite(tok, url);
        const list = await G().graphJson(
            'GET',
            G().graphPathSite(site.id) + '/lists/' + encodeURIComponent(id) + '?$select=description',
            tok,
            undefined,
            'v1.0'
        );
        desc = String((list && list.description) || '');
    } catch {
        desc = '';
    }
    if (!desc) {
        try {
            desc = await fetchListDescriptionViaSpoRest(url, id);
        } catch {
            desc = '';
        }
    }
    return desc;
}

async function fetchListDescriptionViaSpoRest(siteWebUrl, listId) {
    const url = String(siteWebUrl || '').trim().replace(/\/$/, '');
    const id = String(listId || '').trim();
    if (!url || !id) return '';
    let host = '';
    try {
        host = new URL(url).hostname;
    } catch {
        return '';
    }
    if (!host) return '';
    const spoScope = 'https://' + host + '/AllSites.Read';
    const spoTok = await G().getGraphToken([spoScope, 'https://graph.microsoft.com/User.Read']);
    const digest = await G().getSpoRequestDigest(url, spoTok);
    const api = "/_api/web/lists(guid'" + id.replace(/'/g, "''") + "')?$select=Description";
    const res = await G().spoRestFetch(url, spoTok, digest, 'GET', api);
    if (!res || !res.ok) return '';
    const d = res.data && (res.data.Description || (res.data.d && res.data.d.Description));
    return String(d || '');
}

export async function fetchPlannerPermissionsFromList(siteWebUrl, listId) {
    const bundle = await fetchPlannerListDescriptionBundle(siteWebUrl, listId);
    return bundle && bundle.permissions ? bundle.permissions : null;
}

/**
 * @param {string} siteWebUrl
 * @param {string} listId
 */
export async function fetchPlannerListDescriptionBundle(siteWebUrl, listId) {
    const desc = await readListDescriptionText(siteWebUrl, listId);
    if (!desc) return null;
    const full = parseListDescriptionMarkerFull(desc);
    if (!full || !full.permissions) return null;
    return full;
}

export async function publishPlannerPermissionsToSite(siteWebUrl, config, listId) {
    const url = String(siteWebUrl || '').trim().replace(/\/$/, '');
    if (!url) return { ok: false, reason: 'no-site' };
    const cfg = mergePlannerPublishConfig(config || loadEffectivePermissionsConfig());
    if (!entraGroupsConfigured(cfg)) return { ok: false, reason: 'no-groups' };
    const tok = await G().getGraphToken(SCOPES);
    let listOk = false;
    if (listId) {
        try {
            const r = await publishPlannerGroupsToList(url, listId, cfg);
            listOk = !!(r && r.ok);
            if (listOk && cfg.classCatalog && cfg.classCatalog.length) {
                try {
                    const site = await resolveSite(tok, url);
                    await patchFreistellungKlasseColumn(
                        site.id,
                        listId,
                        cfg.classCatalog.map((row) => row.code),
                        { webUrl: url }
                    );
                } catch {
                    /* Spalte Klasse optional */
                }
            }
        } catch {
            /* Datei-Fallback */
        }
    }
    const { driveId, driveLabel } = await resolveConfigDrive(tok, url);
    const body = JSON.stringify(permissionsToRemotePayload(cfg), null, 2);
    const enc = encodeDriveRootPath(REMOTE_CONFIG_REL_PATH);
    const putUrl = G().graphBase('v1.0') + '/drives/' + encodeURIComponent(driveId) + '/' + enc + '/content';
    const res = await fetch(putUrl, {
        method: 'PUT',
        headers: {
            Authorization: 'Bearer ' + tok,
            'Content-Type': 'application/json; charset=utf-8'
        },
        body
    });
    if (!res.ok) {
        const text = await res.text();
        throw new Error('Planer-Konfiguration konnte nicht auf SharePoint gespeichert werden: ' + text);
    }
    return { ok: true, path: REMOTE_CONFIG_REL_PATH, driveLabel, listDescription: listOk };
}

/**
 * @param {string} siteWebUrl
 */
export async function fetchPlannerPermissionsFromSite(siteWebUrl) {
    const url = String(siteWebUrl || '').trim().replace(/\/$/, '');
    if (!url) return null;
    const tok = await G().getGraphToken(SCOPES_READ);
    const { driveId } = await resolveConfigDrive(tok, url);
    let raw = await fetchJsonFromDrive(driveId, REMOTE_CONFIG_REL_PATH, tok);
    if (!raw) {
        for (const legacy of REMOTE_CONFIG_LEGACY_PATHS) {
            raw = await fetchJsonFromDrive(driveId, legacy, tok);
            if (raw) break;
        }
    }
    if (!raw) return null;
    return { permissions: remotePayloadToPermissions(raw), raw };
}

/** @deprecated Nutze fetchPlannerPermissionsFromSite – liefert nur Permissions */
export async function fetchPlannerPermissionsConfigFromSite(siteWebUrl) {
    const packed = await fetchPlannerPermissionsFromSite(siteWebUrl);
    return packed && packed.permissions ? packed.permissions : null;
}

/**
 * @param {string} siteWebUrl
 */
export async function fetchPlannerRemoteRawFromSite(siteWebUrl) {
    const packed = await fetchPlannerPermissionsFromSite(siteWebUrl);
    if (!packed) return null;
    return packed.raw || packed;
}

/**
 * Lädt Remote-Konfiguration und merged in localStorage (nur wenn mindestens eine Gruppen-ID).
 * @param {string} siteWebUrl
 */
/**
 * @param {string} siteWebUrl
 * @param {string} [listId]
 */
export function plannerRemoteHasSchoolData(remote) {
    const r = normalizePermissionsConfig(remote || {});
    return (
        entraGroupsConfigured(r) ||
        normalizeClassCatalog(r.classCatalog).length > 0 ||
        normalizeClassTeamLinks(r.classTeamLinks).length > 0
    );
}

export async function syncPlannerPermissionsFromSite(siteWebUrl, listId, opts) {
    const url = String(siteWebUrl || '').trim().replace(/\/$/, '');
    let id = String(listId || '').trim();
    const listName =
        String((opts && opts.listName) || LIST_TITLE_DEFAULT).trim() || LIST_TITLE_DEFAULT;
    if (!id && url) {
        try {
            id = await tryResolveFrListId(url, {
                listId: (opts && opts.listId) || '',
                listName
            });
        } catch {
            id = '';
        }
    }
    let remote = null;
    if (id) {
        try {
            const bundle = await fetchPlannerListDescriptionBundle(url, id);
            if (bundle) {
                remote = bundle.permissions;
                if (bundle.meta) {
                    if (bundle.meta.siteUrl) persistSiteUrl(bundle.meta.siteUrl);
                    persistSetupFields({
                        siteUrl: bundle.meta.siteUrl || url,
                        listId: bundle.meta.listId || id,
                        listName
                    });
                }
            }
        } catch {
            /* ignore */
        }
    }
    let raw = null;
    if (!remote || !entraGroupsConfigured(remote)) {
        try {
            const packed = await fetchPlannerPermissionsFromSite(url);
            if (packed) {
                remote = packed.permissions;
                raw = packed.raw;
            }
        } catch {
            /* ignore */
        }
    }
    if (!remote || !plannerRemoteHasSchoolData(remote)) return { source: 'none', changed: false };
    savePermissionsConfig(overlaySchoolAudienceOnPermissions(remote));
    if (raw && raw.kategorienExtra) {
        saveExtraKategorien(raw.kategorienExtra);
    }
    return { source: id ? 'list-or-file' : 'file', changed: true };
}

/** @deprecated Nutze syncPlannerPermissionsFromSite */
export const hydratePlannerPermissionsFromSite = syncPlannerPermissionsFromSite;

/**
 * Anzeige-Hinweis für IT (ohne garantierte Web-URL – Drive kann Site Assets sein).
 */
export function remoteConfigPathHint() {
    return 'SiteAssets/' + REMOTE_CONFIG_REL_PATH + ' (oder Standard-Bibliothek/ms365/…)';
}
