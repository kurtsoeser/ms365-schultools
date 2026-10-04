/**
 * Planer-Gruppen mandantenweit: JSON auf der SharePoint-Site (Site Assets).
 */
import {
    normalizePermissionsConfig,
    savePermissionsConfig,
    loadPermissionsConfig
} from './schularbeiten-planer-permissions.js';
import { entraGroupsConfigured } from './schularbeiten-planer-entra-role.js';

export const REMOTE_CONFIG_REL_PATH = 'ms365/schularbeiten-planer-groups.json';
export const REMOTE_CONFIG_LEGACY_PATHS = ['MS365-Schultools/schularbeiten-planer-groups.json'];
export const REMOTE_CONFIG_VERSION = 1;
export const LIST_DESCRIPTION_MARKER = 'MS365_SA_GROUPS:';

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

function encodeDriveRootPath(relativePath) {
    const parts = String(relativePath || '')
        .replace(/^\/+/, '')
        .split('/')
        .filter(Boolean);
    if (!parts.length) return 'root:';
    return 'root:/' + parts.map((p) => encodeURIComponent(p)).join('/') + ':';
}

export function permissionsToRemotePayload(config) {
    const c = normalizePermissionsConfig(config);
    return {
        version: REMOTE_CONFIG_VERSION,
        updatedAt: new Date().toISOString(),
        groupAdmin: c.groupAdmin,
        groupAdminId: c.groupAdminId,
        groupLehrer: c.groupLehrer,
        groupLehrerId: c.groupLehrerId,
        groupSchueler: c.groupSchueler,
        groupSchuelerId: c.groupSchuelerId,
        adminUsers: c.adminUsers,
        lehrerUsers: c.lehrerUsers,
        schuelerUsers: c.schuelerUsers
    };
}

export function permissionsToListDescriptionPayload(config) {
    const c = normalizePermissionsConfig(config);
    return {
        v: REMOTE_CONFIG_VERSION,
        ga: c.groupAdminId,
        gl: c.groupLehrerId,
        gs: c.groupSchuelerId,
        au: compactUserList(c.adminUsers),
        lu: compactUserList(c.lehrerUsers),
        su: compactUserList(c.schuelerUsers)
    };
}

function compactUserList(users) {
    return (users || []).map((u) => ({ id: u.id || '', m: u.mail || '' }));
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
    if (raw.groupLehrerId || raw.groupAdminId || raw.groupSchuelerId) {
        return remotePayloadToPermissions(raw);
    }
    if (!raw.ga && !raw.gl && !raw.gs) return null;
    return normalizePermissionsConfig({
        groupAdminId: raw.ga || '',
        groupLehrerId: raw.gl || '',
        groupSchuelerId: raw.gs || '',
        adminUsers: expandUserList(raw.au),
        lehrerUsers: expandUserList(raw.lu),
        schuelerUsers: expandUserList(raw.su)
    });
}

export function remotePayloadToPermissions(raw) {
    if (!raw || typeof raw !== 'object') return normalizePermissionsConfig({});
    return normalizePermissionsConfig(raw);
}

async function resolveSite(tok, siteWebUrl) {
    const site = await G().resolveSiteFromWebUrl(tok, siteWebUrl);
    if (!site || !site.id) throw new Error('SharePoint-Site nicht gefunden.');
    return site;
}

async function resolveConfigDrive(tok, siteWebUrl) {
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
            if (drive && drive.id) return { siteId, driveId: drive.id };
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
    return { siteId, driveId: drive.id };
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

function parseListDescriptionMarker(desc) {
    const text = String(desc || '');
    if (!text.startsWith(LIST_DESCRIPTION_MARKER)) return null;
    try {
        const raw = JSON.parse(text.slice(LIST_DESCRIPTION_MARKER.length));
        const expanded = expandListDescriptionPayload(raw);
        if (expanded) return expanded;
        return remotePayloadToPermissions(raw);
    } catch {
        return null;
    }
}

export async function fetchPlannerPermissionsFromList(siteWebUrl, listId) {
    const url = String(siteWebUrl || '').trim().replace(/\/$/, '');
    const id = String(listId || '').trim();
    if (!url || !id) return null;
    const readScopes = [...SCOPES_READ, 'https://graph.microsoft.com/Sites.ReadWrite.All'];
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
    return parseListDescriptionMarker(desc);
}

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

export async function syncPlannerPermissionsFromSite(siteWebUrl, listId) {
    let remote = null;
    if (listId) {
        try {
            remote = await fetchPlannerPermissionsFromList(siteWebUrl, listId);
        } catch {
            /* ignore */
        }
    }
    if (!remote || !entraGroupsConfigured(remote)) {
        try {
            const packed = await fetchPlannerPermissionsFromSite(siteWebUrl);
            if (packed) remote = packed.permissions;
        } catch {
            /* ignore */
        }
    }
    if (!remote || !entraGroupsConfigured(remote)) return { source: 'none', changed: false };
    savePermissionsConfig(remote);
    return { source: listId ? 'list-or-file' : 'file', changed: true };
}

export async function publishPlannerPermissionsToSite(siteWebUrl, listId, config) {
    const url = String(siteWebUrl || '').trim().replace(/\/$/, '');
    if (!url) return { ok: false, reason: 'no-site' };
    const cfg = normalizePermissionsConfig(config || loadPermissionsConfig());
    if (!entraGroupsConfigured(cfg)) return { ok: false, reason: 'no-groups' };
    const tok = await G().getGraphToken(SCOPES);
    const { driveId } = await resolveConfigDrive(tok, url);
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
        throw new Error('Schularbeiten-Planer-Konfiguration konnte nicht gespeichert werden: ' + text);
    }
    if (listId) {
        try {
            const site = await resolveSite(tok, url);
            const payload = LIST_DESCRIPTION_MARKER + JSON.stringify(permissionsToListDescriptionPayload(cfg));
            await G().graphJson(
                'PATCH',
                G().graphPathSite(site.id) + '/lists/' + encodeURIComponent(listId),
                tok,
                { description: payload },
                'v1.0'
            );
        } catch {
            /* Datei reicht */
        }
    }
    return { ok: true };
}
