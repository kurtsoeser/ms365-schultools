/**
 * Graph-/MSAL-Helfer für Schulstruktur-Sync (Analyse 02 Phase B).
 * Move-first aus schulstruktur-sync.js – Verhalten unverändert.
 * (Noch eigener Client; später optional shared/graph-client.js.)
 */

import { compareDe } from '../../shared/utils/strings.js';
import {
    loadTenantCache,
    saveTenantCache,
    loadAdGroupFlags
} from './schulstruktur-sync-state.js';
import {
    isInteractionRequired,
    sleep,
    parseTeamsOperationPathFromLocation,
    groupIsTeam,
    graphErrorLooksLikeNotFound,
    isGraphDuplicateRefError,
    odataEscape,
    directoryObjectRef
} from './schulstruktur-sync-graph-helpers.js';

export const GRAPH_SCOPES_TENANT_READ = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.Read.All'
];
/** Gruppen + Benutzerliste (Schritt 4 Grundkonfiguration); User.Read.All für GET /users. */
export const GRAPH_SCOPES_TENANT_INVENTORY = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/Group.Read.All'
];
export const GRAPH_SCOPES_TENANT_WRITE = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.ReadWrite.All'
];
export const GRAPH_SCOPES_TENANT_OWNER_MANAGE = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/Group.ReadWrite.All'
];
/** POST …/teams/{id}/archive|unarchive (Tenant-Details „Team-Archiv“) */
export const GRAPH_SCOPES_TENANT_TEAM_ARCHIVE = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.ReadWrite.All',
    'https://graph.microsoft.com/TeamSettings.ReadWrite.All'
];
/** Gruppe per Graph + Benutzer (Person) per Graph POST /users (Administratorzustimmung User.ReadWrite.All). */
export const GRAPH_SCOPES_GRAPH_OBJECT_CREATE = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/Group.ReadWrite.All',
    'https://graph.microsoft.com/User.ReadWrite.All'
];

// --- Graph (Tenant read) ---
let msalMod = null;
let pca = null;

export async function loadMsal() {
    if (msalMod) return msalMod;
    try {
        msalMod = await import('https://esm.sh/@azure/msal-browser@3.26.1');
    } catch {
        msalMod = await import('https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.26.1/+esm');
    }
    return msalMod;
}

export function resolveMsalConfig() {
    let cfg = window.MS365_MSAL_CONFIG;
    if (!cfg) cfg = {};
    let id = String(cfg.clientId || '').trim();
    if (!id) {
        const meta = document.querySelector('meta[name="ms365-graph-client-id"]');
        const fromMeta = meta && meta.getAttribute('content') ? meta.getAttribute('content').trim() : '';
        if (fromMeta) id = fromMeta;
    }
    if (!id) throw new Error('Keine clientId: ms365-config.js fehlt/leer oder blockiert.');
    return {
        clientId: id,
        authority: cfg.authority || 'https://login.microsoftonline.com/organizations',
        redirectUri: (cfg.redirectUri || window.location.href.split('#')[0]).trim()
    };
}

export async function getPca() {
    const m = await loadMsal();
    const PublicClientApplication = m.PublicClientApplication || (m.default && m.default.PublicClientApplication);
    if (!PublicClientApplication) throw new Error('MSAL: PublicClientApplication nicht gefunden.');
    const cfg = resolveMsalConfig();
    if (!pca) {
        pca = new PublicClientApplication({
            auth: { clientId: cfg.clientId, authority: cfg.authority, redirectUri: cfg.redirectUri },
            cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: true }
        });
        await pca.initialize();
        await pca.handleRedirectPromise();
    }
    return pca;
}

export async function getGraphToken(scopes) {
    // Globaler Login (Header-Widget) – wenn vorhanden, nutzen wir ihn.
    if (typeof window.ms365AuthAcquireToken === 'function') {
        return await window.ms365AuthAcquireToken(scopes);
    }
    const instance = await getPca();
    let accounts = instance.getAllAccounts();
    if (!accounts.length) {
        // In eingebetteten Browsern (z.B. Cursor) bleibt ein Popup gelegentlich schwarz.
        // Redirect-Login ist deutlich robuster.
        try {
            // Nach der Anmeldung wieder zur aktuellen Tool-Seite zurückspringen.
            sessionStorage.setItem('ms365-post-login-url', window.location.href);
        } catch {
            // ignore
        }
        await instance.loginRedirect({ scopes, prompt: 'select_account', redirectStartPage: window.location.href });
        // loginRedirect navigiert weg; Code hier wird normalerweise nicht weiterlaufen.
        throw new Error('Weiterleitung zur Anmeldung …');
    }
    if (!accounts.length) throw new Error('Anmeldung abgebrochen.');
    const req = { scopes, account: accounts[0] };
    try {
        return (await instance.acquireTokenSilent(req)).accessToken;
    } catch (e) {
        if (isInteractionRequired(e)) {
            try {
                sessionStorage.setItem('ms365-post-login-url', window.location.href);
            } catch {
                // ignore
            }
            await instance.acquireTokenRedirect({ ...req, redirectStartPage: window.location.href });
            throw new Error('Weiterleitung zur Anmeldung …');
        }
        throw e;
    }
}

export async function graphRequest(method, pathOrUrl, token, body, extraHeaders) {
    const url = pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
    let attempt = 0;
    while (true) {
        const headers = { Authorization: 'Bearer ' + token };
        if (extraHeaders && typeof extraHeaders === 'object') {
            Object.assign(headers, extraHeaders);
        }
        let payload = undefined;
        if (body !== undefined) {
            headers['Content-Type'] = 'application/json';
            payload = JSON.stringify(body);
        }
        const res = await fetch(url, { method, headers, body: payload });
        if (res.status === 429 && attempt < 8) {
            const ra = parseInt(res.headers.get('Retry-After') || '5', 10);
            await sleep((isNaN(ra) ? 5 : ra) * 1000);
            attempt++;
            continue;
        }
        return res;
    }
}

export async function graphJson(method, pathOrUrl, token, body, extraHeaders) {
    const res = await graphRequest(method, pathOrUrl, token, body, extraHeaders);
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = text;
        }
    }
    if (!res.ok) {
        const msg = typeof data === 'object' && data && data.error ? JSON.stringify(data.error) : text || String(res.status);
        throw new Error(method + ' ' + pathOrUrl + ': ' + msg);
    }
    return data || {};
}

export async function pollTeamsAsyncOperationForTenant(token, operationPath) {
    const maxAttempts = 90;
    for (let i = 0; i < maxAttempts; i++) {
        await sleep(2000);
        const data = await graphJson('GET', operationPath, token, undefined, undefined);
        const st = String(data.status || data.Status || '').toLowerCase();
        if (st === 'succeeded') return;
        if (st === 'failed') {
            const errMsg =
                (data.error && (data.error.message || JSON.stringify(data.error))) || JSON.stringify(data);
            throw new Error('Teams-Operation fehlgeschlagen: ' + errMsg);
        }
    }
    throw new Error('Timeout: Teams-Archivierung nicht abgeschlossen.');
}

export async function setTenantTeamArchiveState(teamId, archive, spoReadonlyForMembers) {
    const token = await getGraphToken(GRAPH_SCOPES_TENANT_TEAM_ARCHIVE);
    const path = '/teams/' + encodeURIComponent(teamId) + (archive ? '/archive' : '/unarchive');
    let body = undefined;
    if (archive && spoReadonlyForMembers) {
        body = { shouldSetSpoSiteReadOnlyForMembers: true };
    }
    const res = await graphRequest('POST', path, token, body, undefined);
    if (res.status !== 202 && res.status !== 200) {
        const t = await res.text();
        throw new Error('HTTP ' + res.status + ' ' + t);
    }
    const loc = res.headers.get('Location') || res.headers.get('Content-Location');
    const opPath = parseTeamsOperationPathFromLocation(loc);
    if (opPath) await pollTeamsAsyncOperationForTenant(token, opPath);
}

export async function fetchAllPages(token, initialPath, onProgress, extraHeaders) {
    const out = [];
    let next = initialPath;
    let page = 0;
    while (next) {
        page++;
        const data = await graphJson('GET', next, token, undefined, extraHeaders);
        const vals = data.value;
        if (Array.isArray(vals)) for (let i = 0; i < vals.length; i++) out.push(vals[i]);
        next = data['@odata.nextLink'] || null;
        if (typeof onProgress === 'function') {
            onProgress({ page, loaded: out.length, hasMore: !!next });
        }
    }
    return out;
}

/*
 * `groupIsTeam` und `graphErrorLooksLikeNotFound` leben in
 * `schulstruktur-sync-graph-helpers.js`.
 */

/**
 * Teams-Archiv (Graph) gilt nur für Objekte mit Teams-Ressource (klassisches Team, Kursteam, …).
 * GET /teams/{id} liefert 404, wenn die Unified-Gruppe kein Team hat.
 * @returns {{ hasTeamsForArchive: boolean, teamIsArchived: boolean|null }}
 */
export async function resolveTeamsArchiveStateForUnifiedGroupId(groupId, token) {
    const gid = String(groupId || '').trim();
    if (!gid) return { hasTeamsForArchive: false, teamIsArchived: null };
    try {
        const team = await graphJson(
            'GET',
            '/teams/' + encodeURIComponent(gid) + '?$select=' + encodeURIComponent('id,isArchived'),
            token,
            undefined
        );
        return {
            hasTeamsForArchive: true,
            teamIsArchived: team.isArchived === true
        };
    } catch (e) {
        if (graphErrorLooksLikeNotFound(e)) {
            return { hasTeamsForArchive: false, teamIsArchived: null };
        }
        return { hasTeamsForArchive: true, teamIsArchived: null };
    }
}

export async function loadTenantGroupsLive(kind, onProgress) {
    const token = await getGraphToken(GRAPH_SCOPES_TENANT_READ);
    const selectBase =
        'id,displayName,description,expirationDateTime,mail,mailNickname,createdDateTime,groupTypes,resourceProvisioningOptions,securityEnabled,mailEnabled,visibility,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesDomainName,onPremisesSecurityIdentifier';

    /** @type {any[]} */
    const all = [];
    const k = String(kind || 'm365');

    // 1) M365 (Unified) Gruppen/Teams
    if (k === 'm365' || k === 'both') {
        const filter = encodeURIComponent("groupTypes/any(c:c eq 'Unified')");
        const initial = '/groups?$filter=' + filter + '&$select=' + encodeURIComponent(selectBase) + '&$top=999';
        const groups = await fetchAllPages(token, initial, onProgress, undefined);
        for (const g of groups) all.push(g);
    }

    // 2) Sicherheitsgruppen (ohne Unified)
    if (k === 'security' || k === 'both') {
        // Advanced query: requires ConsistencyLevel:eventual when using "not(...)"
        const filter = encodeURIComponent("securityEnabled eq true and not(groupTypes/any(c:c eq 'Unified'))");
        const initial =
            '/groups?$count=true&$filter=' + filter + '&$select=' + encodeURIComponent(selectBase) + '&$top=999';
        const groups = await fetchAllPages(token, initial, onProgress, { ConsistencyLevel: 'eventual' });
        for (const g of groups) all.push(g);
    }

    const mapped = all
        .map((g) => {
            const isUnified = Array.isArray(g.groupTypes) && g.groupTypes.indexOf('Unified') !== -1;
            // HiddenMembership ist i.d.R. eine Visibility-Variante (visibility === 'HiddenMembership').
            // Manche Tenants liefern zusätzlich groupTypes-Einträge; wir unterstützen beides.
            const vis = String(g.visibility || '').trim();
            const hiddenMembership =
                vis === 'HiddenMembership' ||
                (Array.isArray(g.groupTypes) && g.groupTypes.indexOf('HiddenMembership') !== -1);
            const isTeam = isUnified && groupIsTeam(g);
            const isSecurity = !!g.securityEnabled && !isUnified;
            const typeLabel = isTeam
                ? 'Team'
                : isUnified
                  ? 'Gruppe'
                  : isSecurity
                    ? g.mailEnabled
                        ? 'E‑Mail‑Sicherheitsgruppe'
                        : 'Sicherheitsgruppe'
                    : 'Gruppe';
            return {
                id: String(g.id || ''),
                bezeichnung: String(g.displayName || ''),
                typ: typeLabel,
                mail: String(g.mail || ''),
                alias: String(g.mailNickname || ''),
                description: String(g.description || ''),
                expirationDateTime: String(g.expirationDateTime || ''),
                visibility: vis,
                hiddenMembership: !!hiddenMembership,
                createdDateTime: String(g.createdDateTime || ''),
                onPremisesSyncEnabled: g.onPremisesSyncEnabled === true,
                onPremisesLastSyncDateTime: String(g.onPremisesLastSyncDateTime || ''),
                onPremisesSamAccountName: String(g.onPremisesSamAccountName || ''),
                onPremisesDomainName: String(g.onPremisesDomainName || ''),
                onPremisesSecurityIdentifier: String(g.onPremisesSecurityIdentifier || '')
            };
        })
        .filter((x) => x.id);

    const seen = new Set();
    const unique = [];
    for (const r of mapped) {
        const id = String(r.id);
        if (seen.has(id)) continue;
        seen.add(id);
        unique.push(r);
    }

    if (typeof onProgress === 'function') {
        onProgress({ phase: 'counts', page: 1, loaded: 0, hasMore: true, total: unique.length });
    }
    const counted = await enrichTenantRowsOwnerMemberCounts(unique, token);
    if (typeof onProgress === 'function') {
        onProgress({ phase: 'counts', page: 1, loaded: unique.length, hasMore: false, total: unique.length });
    }

    counted.sort((a, b) => compareDe(a.bezeichnung, b.bezeichnung));
    const prevUsers = loadTenantCache().users || [];
    const withFlags = applyAdFlagsToTenantRows(counted);
    saveTenantCache(withFlags, prevUsers);
    return withFlags;
}

export async function loadTenantUsersLive(onProgress) {
    const token = await getGraphToken(GRAPH_SCOPES_TENANT_INVENTORY);
    const select = encodeURIComponent(
        'id,displayName,givenName,surname,userPrincipalName,mail,mailNickname,otherMails,accountEnabled,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesDomainName,onPremisesSecurityIdentifier'
    );
    const initial = '/users?$select=' + select + '&$top=999';
    const raw = [];
    let next = initial;
    let page = 0;
    while (next && page < 50 && raw.length < 8000) {
        page++;
        const data = await graphJson('GET', next, token, undefined, undefined);
        const vals = data.value;
        if (Array.isArray(vals)) for (let i = 0; i < vals.length; i++) raw.push(vals[i]);
        next = data['@odata.nextLink'] || null;
        if (typeof onProgress === 'function') {
            onProgress({ phase: 'users', page, loaded: raw.length, hasMore: !!next });
        }
    }
    const mapped = raw
        .map((u) => ({
            id: String(u.id || ''),
            displayName: String(u.displayName || '').trim(),
            givenName: String(u.givenName || '').trim(),
            surname: String(u.surname || '').trim(),
            userPrincipalName: String(u.userPrincipalName || '').trim().toLowerCase(),
            mail: String(u.mail || '').trim().toLowerCase(),
            mailNickname: String(u.mailNickname || '').trim().toLowerCase(),
            otherMails: Array.isArray(u.otherMails)
                ? u.otherMails.map((m) => String(m || '').trim().toLowerCase()).filter(Boolean)
                : [],
            accountEnabled: u.accountEnabled !== false,
            onPremisesSyncEnabled: u.onPremisesSyncEnabled === true,
            onPremisesLastSyncDateTime: String(u.onPremisesLastSyncDateTime || ''),
            onPremisesSamAccountName: String(u.onPremisesSamAccountName || ''),
            onPremisesDomainName: String(u.onPremisesDomainName || ''),
            onPremisesSecurityIdentifier: String(u.onPremisesSecurityIdentifier || '')
        }))
        .filter((x) => x.id);
    mapped.sort((a, b) => compareDe(a.displayName || a.userPrincipalName, b.displayName || b.userPrincipalName));
    return mapped;
}

export async function loadTenantInventoryFull(onProgress) {
    const kind = 'both';
    await loadTenantGroupsLive(kind, (p) => {
        if (typeof onProgress !== 'function') return;
        if (p && p.phase === 'counts') onProgress(p);
        else onProgress(Object.assign({ phase: 'groups' }, p));
    });
    const afterG = loadTenantCache();
    const users = await loadTenantUsersLive((p) => {
        if (typeof onProgress === 'function') onProgress(p);
    });
    saveTenantCache(afterG.rows, users);
    return {
        groups: afterG.rows,
        users,
        loadedAt: new Date().toISOString()
    };
}

export async function fetchTenantGroupDetail(groupId) {
    const token = await getGraphToken(GRAPH_SCOPES_TENANT_READ);
    const sel =
        'id,displayName,description,expirationDateTime,mail,mailNickname,groupTypes,resourceProvisioningOptions,securityEnabled,mailEnabled,visibility,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesDomainName,onPremisesSecurityIdentifier';
    const g = await graphJson('GET', '/groups/' + encodeURIComponent(groupId) + '?$select=' + encodeURIComponent(sel), token, undefined);
    const isUnified = Array.isArray(g.groupTypes) && g.groupTypes.indexOf('Unified') !== -1;
    const vis = String(g.visibility || '').trim();
    const hiddenMembership =
        vis === 'HiddenMembership' ||
        (Array.isArray(g.groupTypes) && g.groupTypes.indexOf('HiddenMembership') !== -1);
    const isTeam = isUnified && groupIsTeam(g);
    const isSecurity = !!g.securityEnabled && !isUnified;
    const typ = isTeam ? 'Team' : isUnified ? 'Gruppe' : isSecurity ? (g.mailEnabled ? 'E‑Mail‑Sicherheitsgruppe' : 'Sicherheitsgruppe') : 'Gruppe';
    /** @type {boolean|null} */
    let teamIsArchived = null;
    /** @type {boolean} */
    let hasTeamsForArchive = false;
    if (isUnified) {
        const ar = await resolveTeamsArchiveStateForUnifiedGroupId(String(g.id || ''), token);
        hasTeamsForArchive = ar.hasTeamsForArchive;
        teamIsArchived = ar.teamIsArchived;
    }
    const row = {
        id: String(g.id || ''),
        bezeichnung: String(g.displayName || ''),
        typ,
        mail: String(g.mail || ''),
        alias: String(g.mailNickname || ''),
        description: String(g.description || ''),
        expirationDateTime: String(g.expirationDateTime || ''),
        visibility: vis,
        hiddenMembership: !!hiddenMembership,
        teamIsArchived,
        hasTeamsForArchive,
        onPremisesSyncEnabled: g.onPremisesSyncEnabled === true,
        onPremisesLastSyncDateTime: String(g.onPremisesLastSyncDateTime || ''),
        onPremisesSamAccountName: String(g.onPremisesSamAccountName || ''),
        onPremisesDomainName: String(g.onPremisesDomainName || ''),
        onPremisesSecurityIdentifier: String(g.onPremisesSecurityIdentifier || '')
    };
    return applyAdFlagsToTenantRows([row])[0];
}

/**
 * Markierungen aus localStorage auf Tenant-Zeilen legen.
 * @param {any[]} rows
 * @returns {any[]}
 */
export function applyAdFlagsToTenantRows(rows) {
    const flags = loadAdGroupFlags();
    return (Array.isArray(rows) ? rows : []).map((r) => {
        if (!r || !r.id) return r;
        const f = flags[String(r.id)] || null;
        return Object.assign({}, r, {
            adFlagged: !!(f && f.flagged),
            adFlagNote: f && f.note ? String(f.note) : '',
            adFlaggedAt: f && f.flaggedAt ? String(f.flaggedAt) : ''
        });
    });
}

/*
 * `personLabel` und `odataEscape` leben in
 * `schulstruktur-sync-graph-helpers.js`.
 */

export async function graphSearchUsersForOwner(token, query) {
    const q = String(query || '').trim();
    if (!q) return [];
    const esc = odataEscape(q);
    let filter;
    if (q.indexOf('@') !== -1) {
        filter = "(mail eq '" + esc + "' or userPrincipalName eq '" + esc + "')";
    } else {
        filter =
            "(startswith(displayName,'" +
            esc +
            "') or startswith(userPrincipalName,'" +
            esc +
            "') or startswith(mail,'" +
            esc +
            "'))";
    }
    const select = 'id,displayName,mail,userPrincipalName';
    const path =
        '/users?$filter=' +
        encodeURIComponent(filter) +
        '&$select=' +
        encodeURIComponent(select) +
        '&$top=25';
    const data = await graphJson('GET', path, token, undefined);
    return data.value || [];
}

export async function graphGetCollectionCount(token, groupId, segment) {
    const gid = String(groupId || '').trim();
    if (!gid) return -1;
    const seg = segment === 'owners' ? 'owners' : 'members';
    const path = '/groups/' + encodeURIComponent(gid) + '/' + seg + '/$count';
    const res = await graphRequest('GET', path, token, undefined, { ConsistencyLevel: 'eventual' });
    const text = await res.text();
    if (!res.ok) return -1;
    const n = parseInt(String(text).trim(), 10);
    return isNaN(n) ? -1 : n;
}

export async function mapWithConcurrencyLimited(items, limit, fn) {
    const results = new Array(items.length);
    let i = 0;
    async function worker() {
        while (true) {
            const idx = i++;
            if (idx >= items.length) return;
            results[idx] = await fn(items[idx], idx);
        }
    }
    const n = Math.max(1, Math.min(limit, items.length || 1));
    const workers = [];
    for (let w = 0; w < n; w++) workers.push(worker());
    await Promise.all(workers);
    return results;
}

/** Nach Gruppenliste: Owner-/Mitglieder-$count für Filter „ohne …“. */
export async function enrichTenantRowsOwnerMemberCounts(rows, token) {
    return await mapWithConcurrencyLimited(rows, 10, async (row) => {
        const id = String(row && row.id ? row.id : '').trim();
        if (!id) return Object.assign({}, row, { ownerCount: -1, memberCount: -1 });
        const oc = await graphGetCollectionCount(token, id, 'owners');
        const mc = await graphGetCollectionCount(token, id, 'members');
        return Object.assign({}, row, { ownerCount: oc, memberCount: mc });
    });
}

/* `directoryObjectRef` lebt in `schulstruktur-sync-graph-helpers.js`. */

export async function addGroupOwner(groupId, userId) {
    const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
    const body = { '@odata.id': directoryObjectRef(userId) };
    await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/owners/$ref', token, body);
}

export async function addGroupMember(groupId, userId) {
    const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
    const body = { '@odata.id': directoryObjectRef(userId) };
    await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/members/$ref', token, body, undefined);
}

/* `isGraphDuplicateRefError` lebt in `schulstruktur-sync-graph-helpers.js`. */

export async function addOwnerWithMemberFallback(groupId, userId) {
    try {
        await addGroupOwner(groupId, userId);
    } catch (e1) {
        // Manche Tenants/Policies verlangen, dass Owner auch Member ist.
        try {
            await addGroupMember(groupId, userId);
        } catch (e2) {
            if (!isGraphDuplicateRefError(e2)) throw e2;
        }
        await addGroupOwner(groupId, userId);
    }
}

export async function deleteTenantGroup(groupId) {
    const token = await getGraphToken(GRAPH_SCOPES_TENANT_WRITE);
    await graphJson('DELETE', '/groups/' + encodeURIComponent(groupId), token, undefined, undefined);
}

export async function createUnifiedGroup(displayName, description, mailNickname, visibility) {
    const token = await getGraphToken(GRAPH_SCOPES_TENANT_WRITE);
    const body = {
        displayName: String(displayName || '').trim(),
        description: String(description || '').trim(),
        mailEnabled: true,
        mailNickname: String(mailNickname || '').trim(),
        securityEnabled: false,
        groupTypes: ['Unified']
    };
    const vis = String(visibility || '').trim();
    if (vis === 'Private' || vis === 'Public' || vis === 'HiddenMembership') body.visibility = vis;
    if (!body.displayName) throw new Error('Bitte einen Anzeigenamen eingeben.');
    if (!body.mailNickname) throw new Error('Mail‑Nickname fehlt (Vorschlag ist leer).');
    return await graphJson('POST', '/groups', token, body, undefined);
}

export async function createMailEnabledSecurityGroup(displayName, description, mailNickname) {
    const token = await getGraphToken(GRAPH_SCOPES_TENANT_WRITE);
    const body = {
        displayName: String(displayName || '').trim(),
        description: String(description || '').trim(),
        mailEnabled: true,
        mailNickname: String(mailNickname || '').trim(),
        securityEnabled: true,
        groupTypes: []
    };
    if (!body.displayName) throw new Error('Bitte einen Anzeigenamen eingeben.');
    if (!body.mailNickname) throw new Error('Mail‑Nickname fehlt (Vorschlag ist leer).');
    return await graphJson('POST', '/groups', token, body, undefined);
}

export async function createSecurityGroup(displayName, description, mailNickname) {
    const token = await getGraphToken(GRAPH_SCOPES_TENANT_WRITE);
    const body = {
        displayName: String(displayName || '').trim(),
        description: String(description || '').trim(),
        mailEnabled: false,
        mailNickname: String(mailNickname || '').trim(),
        securityEnabled: true,
        groupTypes: []
    };
    if (!body.displayName) throw new Error('Bitte einen Anzeigenamen eingeben.');
    if (!body.mailNickname) throw new Error('Alias (mailNickname) fehlt.');
    return await graphJson('POST', '/groups', token, body, undefined);
}

/** Alias-Vorschlag aus Anzeigename (ASCII, Graph-tauglich). */
export function suggestGroupMailNickname(displayName) {
    let s = String(displayName || '')
        .trim()
        .toLowerCase()
        .replace(/ä/g, 'ae')
        .replace(/ö/g, 'oe')
        .replace(/ü/g, 'ue')
        .replace(/ß/g, 'ss');
    try {
        s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    } catch {
        /* ignore */
    }
    s = s.replace(/[^a-z0-9._-]+/g, '-').replace(/-+/g, '-').replace(/^[-._]+|[-._]+$/g, '');
    if (!s) s = 'gruppe';
    return s.slice(0, 60);
}

export async function createTeamForGroup(groupId) {
    const token = await getGraphToken(GRAPH_SCOPES_TENANT_WRITE);
    await graphJson('PUT', '/groups/' + encodeURIComponent(groupId) + '/team', token, {}, undefined);
}
