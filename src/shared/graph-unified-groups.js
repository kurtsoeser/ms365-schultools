(function () {
    'use strict';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/User.ReadWrite.All',
        'https://graph.microsoft.com/Group.ReadWrite.All',
        'https://graph.microsoft.com/Team.Create',
        'https://graph.microsoft.com/TeamSettings.ReadWrite.All'
    ];

    const PERSON_SELECT = 'id,displayName,mail,userPrincipalName';
    const USER_LICENSE_SELECT =
        'id,displayName,givenName,surname,mail,userPrincipalName,accountEnabled,userType,jobTitle,department,assignedLicenses';

    let msalMod = null;
    let pca = null;

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normEmail(v) {
        return normStr(v).toLowerCase();
    }

    function noteAction(entry) {
        try {
            if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                window.ms365ActionLog.append(entry);
                return;
            }
            const api = window.ms365AppDataV2;
            if (!api || typeof api.getSetup !== 'function' || typeof api.patchSetup !== 'function') return;
            const cur = api.getSetup() || {};
            const list = Array.isArray(cur.actionLog) ? cur.actionLog.slice() : [];
            list.push({
                at: new Date().toISOString(),
                tool: String((entry && entry.tool) || 'graph'),
                action: String((entry && entry.action) || 'write'),
                target: String((entry && entry.target) || ''),
                summary: String((entry && entry.summary) || ''),
                result: entry && entry.result === 'error' ? 'error' : 'ok'
            });
            while (list.length > 200) list.shift();
            api.patchSetup({ actionLog: list });
        } catch {
            /* Protokoll darf Writes nicht blockieren */
        }
    }

    async function loadMsal() {
        if (msalMod) return msalMod;
        const loader = await import('./msal-loader.js');
        if (typeof loader.loadMsalBrowser !== 'function') {
            throw new Error('MSAL-Loader: loadMsalBrowser fehlt.');
        }
        msalMod = await loader.loadMsalBrowser();
        return msalMod;
    }

    function isInteractionRequired(e) {
        return (
            e &&
            (e.name === 'InteractionRequiredAuthError' ||
                e.errorCode === 'interaction_required' ||
                (typeof e.message === 'string' && e.message.indexOf('interaction_required') !== -1))
        );
    }

    function resolveMsalConfig() {
        let cfg = window.MS365_MSAL_CONFIG;
        if (!cfg) cfg = {};
        let id = String(cfg.clientId || '').trim();
        if (!id) {
            const meta = document.querySelector('meta[name="ms365-graph-client-id"]');
            const fromMeta = meta && meta.getAttribute('content') ? meta.getAttribute('content').trim() : '';
            if (fromMeta) id = fromMeta;
        }
        if (!id) {
            throw new Error(
                'Keine clientId: ms365-config.js fehlt/leer oder blockiert. Seite mit Strg+F5 neu laden.'
            );
        }
        return {
            clientId: id,
            authority: cfg.authority || 'https://login.microsoftonline.com/organizations',
            redirectUri: (cfg.redirectUri || window.location.href.split('#')[0]).trim()
        };
    }

    async function getPca() {
        const m = await loadMsal();
        const PublicClientApplication = m.PublicClientApplication || (m.default && m.default.PublicClientApplication);
        if (!PublicClientApplication) {
            throw new Error('MSAL: PublicClientApplication nicht gefunden (Import).');
        }
        const cfg = resolveMsalConfig();
        if (!pca) {
            pca = new PublicClientApplication({
                auth: {
                    clientId: cfg.clientId,
                    authority: cfg.authority,
                    redirectUri: cfg.redirectUri
                },
                cache: {
                    cacheLocation: 'sessionStorage',
                    storeAuthStateInCookie: true
                }
            });
            await pca.initialize();
            await pca.handleRedirectPromise();
        }
        return pca;
    }

    async function getGraphToken() {
        const instance = await getPca();
        let accounts = instance.getAllAccounts();
        if (!accounts.length) {
            await instance.loginPopup({ scopes: GRAPH_SCOPES, prompt: 'select_account' });
            accounts = instance.getAllAccounts();
        }
        if (!accounts.length) {
            throw new Error('Anmeldung abgebrochen.');
        }
        const req = { scopes: GRAPH_SCOPES, account: accounts[0] };
        try {
            return (await instance.acquireTokenSilent(req)).accessToken;
        } catch (e) {
            if (isInteractionRequired(e)) {
                return (await instance.acquireTokenPopup(req)).accessToken;
            }
            throw e;
        }
    }

    function sleep(ms) {
        return new Promise(function (r) {
            setTimeout(r, ms);
        });
    }

    async function graphRequest(method, path, token, body, extraHeaders) {
        const url = path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path;
        let attempt = 0;
        while (true) {
            const headers = { Authorization: 'Bearer ' + token };
            if (extraHeaders && typeof extraHeaders === 'object') {
                Object.assign(headers, extraHeaders);
            }
            if (body !== undefined) headers['Content-Type'] = 'application/json';
            const res = await fetch(url, {
                method: method,
                headers: headers,
                body: body !== undefined ? JSON.stringify(body) : undefined
            });
            if (res.status === 429 && attempt < 8) {
                const ra = parseInt(res.headers.get('Retry-After') || '5', 10);
                await sleep((isNaN(ra) ? 5 : ra) * 1000);
                attempt++;
                continue;
            }
            return res;
        }
    }

    async function graphJson(method, path, token, body, extraHeaders) {
        const res = await graphRequest(method, path, token, body, extraHeaders);
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
            const msg =
                typeof data === 'object' && data && data.error ? JSON.stringify(data.error) : text || String(res.status);
            throw new Error(method + ' ' + path + ': ' + msg);
        }
        return data || {};
    }

    async function graphRawRequest(method, path, token, body, extraHeaders) {
        const url = path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path;
        let attempt = 0;
        while (true) {
            const headers = { Authorization: 'Bearer ' + token };
            if (extraHeaders && typeof extraHeaders === 'object') {
                Object.assign(headers, extraHeaders);
            }
            const res = await fetch(url, {
                method: method,
                headers: headers,
                body: body !== undefined ? body : undefined
            });
            if (res.status === 429 && attempt < 8) {
                const ra = parseInt(res.headers.get('Retry-After') || '5', 10);
                await sleep((isNaN(ra) ? 5 : ra) * 1000);
                attempt++;
                continue;
            }
            return res;
        }
    }

    const GROUP_PHOTO_MAX_BYTES = 4 * 1024 * 1024;

    function groupPhotoInitials(displayName) {
        const s = String(displayName || '').trim();
        if (!s) return '?';
        const parts = s.split(/\s+/).filter(Boolean);
        if (parts.length >= 2) {
            return (parts[0].charAt(0) + parts[1].charAt(0)).toUpperCase();
        }
        return s.slice(0, 2).toUpperCase();
    }

    async function fetchGroupPhotoBlob(token, groupId) {
        const gid = encodeURIComponent(normStr(groupId));
        if (!gid) return null;
        const path = '/groups/' + gid + '/photo/$value';
        const res = await graphRawRequest('GET', path, token);
        if (res.status === 404) return null;
        if (!res.ok) {
            const text = await res.text();
            throw new Error('GET ' + path + ': ' + (text || String(res.status)));
        }
        return res.blob();
    }

    async function setGroupPhoto(token, groupId, imageBlob, contentType) {
        const gid = encodeURIComponent(normStr(groupId));
        if (!gid) throw new Error('Gruppen-ID fehlt.');
        const blob = imageBlob instanceof Blob ? imageBlob : null;
        if (!blob || !blob.size) throw new Error('Kein Bild.');
        if (blob.size > GROUP_PHOTO_MAX_BYTES) {
            throw new Error('Bild zu groß (max. 4 MB).');
        }
        const ct = normStr(contentType || blob.type || 'image/jpeg') || 'image/jpeg';
        const path = '/groups/' + gid + '/photo/$value';
        const res = await graphRawRequest('PUT', path, token, blob, { 'Content-Type': ct });
        if (!res.ok) {
            const text = await res.text();
            throw new Error('PUT ' + path + ': ' + (text || String(res.status)));
        }
        noteAction({
            tool: 'graph',
            action: 'groupPhotoSet',
            target: decodeURIComponent(gid),
            summary: 'Gruppenbild gesetzt'
        });
    }

    async function deleteGroupPhoto(token, groupId) {
        const gid = encodeURIComponent(normStr(groupId));
        if (!gid) throw new Error('Gruppen-ID fehlt.');
        const path = '/groups/' + gid + '/photo/$value';
        const res = await graphRawRequest('DELETE', path, token);
        if (res.status === 404) return;
        if (!res.ok) {
            const text = await res.text();
            throw new Error('DELETE ' + path + ': ' + (text || String(res.status)));
        }
        noteAction({
            tool: 'graph',
            action: 'groupPhotoDelete',
            target: decodeURIComponent(gid),
            summary: 'Gruppenbild entfernt'
        });
    }

    async function setTeamPhoto(token, teamId, imageBlob, contentType) {
        const tid = encodeURIComponent(normStr(teamId));
        if (!tid) throw new Error('Team-ID fehlt.');
        const blob = imageBlob instanceof Blob ? imageBlob : null;
        if (!blob || !blob.size) throw new Error('Kein Bild.');
        if (blob.size > GROUP_PHOTO_MAX_BYTES) {
            throw new Error('Bild zu groß (max. 4 MB).');
        }
        const ct = normStr(contentType || blob.type || 'image/jpeg') || 'image/jpeg';
        const path = '/teams/' + tid + '/photo/$value';
        const res = await graphRawRequest('PUT', path, token, blob, { 'Content-Type': ct });
        if (!res.ok) {
            const text = await res.text();
            throw new Error('PUT ' + path + ': ' + (text || String(res.status)));
        }
        noteAction({
            tool: 'graph',
            action: 'teamPhotoSet',
            target: decodeURIComponent(tid),
            summary: 'Teams-Bild gesetzt'
        });
    }

    /**
     * Teams-Foto entfernen. Graph unterstützt DELETE für Teams laut Doku oft nicht –
     * Fehler werden abgefangen und als { ok: false } zurückgegeben.
     */
    async function deleteTeamPhoto(token, teamId) {
        const tid = encodeURIComponent(normStr(teamId));
        if (!tid) throw new Error('Team-ID fehlt.');
        const path = '/teams/' + tid + '/photo/$value';
        const res = await graphRawRequest('DELETE', path, token);
        if (res.status === 404) {
            return { ok: true, skipped: true, reason: 'Kein Teams-Bild vorhanden.' };
        }
        if (!res.ok) {
            const text = await res.text();
            return {
                ok: false,
                reason:
                    res.status === 405 || res.status === 501
                        ? 'Teams-Bild kann per Graph nicht gelöscht werden (Microsoft-Limitierung).'
                        : 'DELETE ' + path + ': ' + (text || String(res.status))
            };
        }
        noteAction({
            tool: 'graph',
            action: 'teamPhotoDelete',
            target: decodeURIComponent(tid),
            summary: 'Teams-Bild entfernt'
        });
        return { ok: true };
    }

    async function syncTeamPhotoForGroup(token, groupId, mode, imageBlob, contentType, opts) {
        const gid = normStr(groupId);
        if (!gid) return { ok: false, skipped: true, reason: 'Gruppen-ID fehlt.' };
        const o = opts && typeof opts === 'object' ? opts : {};
        let hasTeam = o.hasTeam === true || o.hasTeam === false ? o.hasTeam : null;
        if (hasTeam === null) {
            try {
                const g = await fetchGroup(token, gid);
                hasTeam = groupHasTeam(g);
            } catch {
                return { ok: false, skipped: true, reason: 'Team-Status unbekannt.' };
            }
        }
        if (!hasTeam) {
            return { ok: true, skipped: true, reason: 'Kein Microsoft Team.' };
        }
        try {
            if (mode === 'set') {
                await setTeamPhoto(token, gid, imageBlob, contentType);
                return { ok: true };
            }
            if (mode === 'delete') {
                return await deleteTeamPhoto(token, gid);
            }
        } catch (e) {
            return { ok: false, reason: String(e && e.message ? e.message : e) };
        }
        return { ok: false, reason: 'Unbekannter Modus.' };
    }

    function teamPhotoSyncHint(result) {
        if (!result || result.skipped) return '';
        if (result.ok && !result.skipped) return ' Teams-Bild mitgesetzt.';
        const r = String(result.reason || '').trim();
        if (!r) return ' Teams-Bild konnte nicht mitgesetzt werden.';
        return ' Teams: ' + r;
    }

    function odataEscape(s) {
        return String(s).replace(/'/g, "''");
    }

    function sanitizeMailNickname(name) {
        let n = String(name || '')
            .replace(/[^0-9a-zA-Z]/g, '')
            .slice(0, 60);
        if (!n) n = 'group';
        return n.toLowerCase();
    }

    /** group.mailNickname: unzulässig u. a. laut Microsoft Learn (validateProperties): @ ( ) \ [ ] " ; : < > , Leerzeichen */
    const GRAPH_MAILNICKNAME_INVALID = /[@()[\]\\";:<>,\s]/;

    function sanitizeUnifiedGroupMailNickname(raw) {
        const s = String(raw ?? '')
            .trim()
            .toLowerCase();
        let out = '';
        for (let i = 0; i < s.length; i++) {
            const c = s.charCodeAt(i);
            if (c < 32 || c === 127 || c > 127) continue;
            const ch = s.charAt(i);
            if (GRAPH_MAILNICKNAME_INVALID.test(ch)) continue;
            out += ch;
        }
        if (!out) out = 'group';
        return out.slice(0, 60);
    }

    function isUnifiedGroup(g) {
        const gt = g && g.groupTypes;
        return Array.isArray(gt) && gt.indexOf('Unified') !== -1;
    }

    function escapeSearchPhrase(raw) {
        return String(raw || '')
            .replace(/"/g, '\\"')
            .replace(/\r?\n/g, ' ')
            .trim();
    }

    async function searchUnifiedGroups(token, queryRaw) {
        const q = normStr(queryRaw);
        if (!q) return [];
        try {
            const phrase = escapeSearchPhrase(q);
            const aqs =
                '(displayName:' +
                phrase +
                ' OR mail:' +
                phrase +
                ' OR mailNickname:' +
                phrase +
                ' OR description:' +
                phrase +
                ')';
            const path =
                '/groups?$search=' +
                encodeURIComponent('"' + aqs + '"') +
                '&$select=' +
                encodeURIComponent('id,displayName,mail,mailNickname,groupTypes,description') +
                '&$top=25';
            const data = await graphJson('GET', path, token, undefined, { ConsistencyLevel: 'eventual' });
            const list = (data && data.value) || [];
            return list.filter(isUnifiedGroup);
        } catch {
            // fallback
        }
        const esc = odataEscape(q);
        const filter =
            "groupTypes/any(c:c eq 'Unified') and (" +
            "startswith(displayName,'" +
            esc +
            "') or startswith(mailNickname,'" +
            esc +
            "') or startswith(mail,'" +
            esc +
            "') )";
        const path =
            '/groups?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent('id,displayName,mail,mailNickname,groupTypes,description') +
            '&$top=25';
        const data = await graphJson('GET', path, token, undefined);
        return data.value || [];
    }

    async function fetchGroup(token, id) {
        const path =
            '/groups/' +
            encodeURIComponent(id) +
            '?$select=' +
            encodeURIComponent(
                'id,displayName,mail,mailNickname,groupTypes,description,visibility,createdDateTime,expirationDateTime,renewedDateTime,resourceProvisioningOptions'
            );
        const g = await graphJson('GET', path, token, undefined);
        if (g && !g.expirationDateTime) {
            try {
                const extra = await graphJson(
                    'GET',
                    'https://graph.microsoft.com/beta/groups/' +
                        encodeURIComponent(id) +
                        '?$select=' +
                        encodeURIComponent('expirationDateTime,renewedDateTime'),
                    token
                );
                if (extra && extra.expirationDateTime) g.expirationDateTime = extra.expirationDateTime;
                if (extra && extra.renewedDateTime && !g.renewedDateTime) g.renewedDateTime = extra.renewedDateTime;
            } catch {
                /* v1.0 bleibt */
            }
        }
        return g;
    }

    function groupHasTeam(g) {
        const opts = g && g.resourceProvisioningOptions;
        return Array.isArray(opts) && opts.indexOf('Team') !== -1;
    }

    function groupArtLabel(g) {
        if (isUnifiedGroup(g)) return 'Microsoft 365‑Gruppe';
        if (g && g.securityEnabled && g.mailEnabled) return 'E-Mail-Sicherheitsgruppe';
        if (g && g.securityEnabled && !g.mailEnabled) return 'Sicherheitsgruppe';
        if (g && g.mailEnabled) return 'Mail-aktivierte Gruppe';
        return '–';
    }

    function teamDeepLink(groupId, tenantId) {
        const gid = String(groupId || '').trim();
        if (!gid) return '';
        const tid = String(tenantId || '').trim();
        let url =
            'https://teams.microsoft.com/l/team/' +
            encodeURIComponent(gid) +
            '/conversations?groupId=' +
            encodeURIComponent(gid);
        if (tid) url += '&tenantId=' + encodeURIComponent(tid);
        return url;
    }

    async function getTenantId() {
        try {
            const instance = await getPca();
            const accounts = instance.getAllAccounts();
            const a = accounts[0];
            if (!a) return '';
            return String(a.tenantId || (a.idTokenClaims && a.idTokenClaims.tid) || '').trim();
        } catch {
            return '';
        }
    }

    async function fetchTeamWebUrl(token, groupId) {
        const gid = String(groupId || '').trim();
        if (!gid) return '';
        try {
            const t = await graphJson(
                'GET',
                '/teams/' + encodeURIComponent(gid) + '?$select=' + encodeURIComponent('id,webUrl'),
                token
            );
            const w = t && t.webUrl ? String(t.webUrl).trim() : '';
            if (w) return w;
        } catch {
            /* Fallback unten */
        }
        return teamDeepLink(gid, await getTenantId());
    }

    async function patchGroup(token, groupId, opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const body = {};
        if (o.displayName !== undefined && o.displayName !== null) {
            body.displayName = String(o.displayName).trim();
        }
        if (o.description !== undefined && o.description !== null) {
            body.description = String(o.description).trim();
        }
        const vis = o.visibility !== undefined && o.visibility !== null ? String(o.visibility).trim() : '';
        if (vis === 'Private' || vis === 'Public') {
            body.visibility = vis;
        }
        if (o.mailNickname !== undefined && o.mailNickname !== null) {
            body.mailNickname = sanitizeUnifiedGroupMailNickname(o.mailNickname);
        }
        if (!Object.keys(body).length) return {};
        try {
            const res = await graphJson('PATCH', '/groups/' + encodeURIComponent(groupId), token, body);
            noteAction({
                tool: 'graph',
                action: 'patch-group',
                target: String(groupId),
                summary: 'Gruppe aktualisiert' + (body.displayName ? ': ' + body.displayName : '')
            });
            return res;
        } catch (e) {
            noteAction({
                tool: 'graph',
                action: 'patch-group',
                target: String(groupId),
                summary: String(e && e.message ? e.message : e),
                result: 'error'
            });
            throw e;
        }
    }

    async function patchGroupDisplayName(token, groupId, displayName, description) {
        return patchGroup(token, groupId, { displayName: displayName, description: description });
    }

    async function renewGroup(token, groupId) {
        return graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/renew', token, undefined);
    }

    function userRef(userId) {
        return 'https://graph.microsoft.com/v1.0/users/' + userId;
    }

    function directoryObjectRef(id) {
        return 'https://graph.microsoft.com/v1.0/directoryObjects/' + id;
    }

    function personLabel(p) {
        if (!p || typeof p !== 'object') return '';
        const dn = p.displayName ? String(p.displayName).trim() : '';
        const upn = p.userPrincipalName || p.mail ? String(p.userPrincipalName || p.mail).trim() : '';
        if (dn && upn && dn !== upn) return dn + ' (' + upn + ')';
        return dn || upn || (p.id ? String(p.id) : '');
    }

    function isDuplicateMemberError(e) {
        const m = String((e && e.message) || e || '');
        return (
            m.indexOf('added object references already exist') !== -1 ||
            m.indexOf('One or more added object references already exist') !== -1 ||
            m.indexOf('already exist') !== -1
        );
    }

    function compareDe(a, b) {
        try {
            return String(a || '').localeCompare(String(b || ''), 'de', { sensitivity: 'base' });
        } catch {
            return String(a || '').localeCompare(String(b || ''));
        }
    }

    async function fetchAllPagesSimple(token, initialPath, maxItems, onProgress) {
        const limit = typeof maxItems === 'number' && maxItems > 0 ? maxItems : 4000;
        const out = [];
        let next = initialPath;
        let pages = 0;
        while (next && pages < 40 && out.length < limit) {
            pages++;
            const data = await graphJson('GET', next, token, undefined);
            const vals = data.value;
            if (Array.isArray(vals)) {
                for (let i = 0; i < vals.length; i++) out.push(vals[i]);
            }
            next = data['@odata.nextLink'] || null;
            if (typeof onProgress === 'function') onProgress(out.length, pages, !!next);
        }
        return out;
    }

    async function fetchSubscribedSkus(token) {
        try {
            const data = await graphJson(
                'GET',
                '/subscribedSkus?$select=' +
                    encodeURIComponent('skuId,skuPartNumber,prepaidUnits,consumedUnits,capabilityStatus'),
                token,
                undefined
            );
            const list = Array.isArray(data.value) ? data.value : [];
            return { ok: true, skus: list };
        } catch (e) {
            return { ok: false, skus: [], error: e };
        }
    }

    function skuLookupFromSubscribed(skus) {
        const map = new Map();
        (Array.isArray(skus) ? skus : []).forEach(function (s) {
            const id = String((s && s.skuId) || '').toLowerCase();
            if (!id) return;
            map.set(id, {
                skuId: id,
                skuPartNumber: String((s && s.skuPartNumber) || '')
            });
        });
        return map;
    }

    async function fetchUsersWithAssignedLicenses(token, onProgress) {
        const path = '/users?$select=' + encodeURIComponent(USER_LICENSE_SELECT) + '&$top=999';
        return fetchAllPagesSimple(token, path, 8000, onProgress);
    }

    async function fetchUsersByAssignedSkuIds(token, skuIds, onProgress) {
        const ids = (Array.isArray(skuIds) ? skuIds : [])
            .map(function (id) {
                return String(id || '').toLowerCase();
            })
            .filter(Boolean);
        const byId = new Map();
        let filterFailed = 0;
        for (let i = 0; i < ids.length; i++) {
            const skuId = ids[i];
            const filter = "assignedLicenses/any(s:s/skuId eq " + skuId + ")";
            const path =
                '/users?$filter=' +
                encodeURIComponent(filter) +
                '&$select=' +
                encodeURIComponent(USER_LICENSE_SELECT) +
                '&$top=999';
            try {
                const batch = await fetchAllPagesSimple(token, path, 4000, function (count) {
                    if (typeof onProgress === 'function') {
                        onProgress(byId.size + count, i + 1, ids.length);
                    }
                });
                for (let j = 0; j < batch.length; j++) {
                    const u = batch[j];
                    if (u && u.id) byId.set(u.id, u);
                }
            } catch {
                filterFailed++;
            }
        }
        if (!byId.size && filterFailed === ids.length && ids.length) {
            return fetchUsersWithAssignedLicenses(token, onProgress);
        }
        return Array.from(byId.values());
    }

    async function searchUsers(token, query) {
        const q = normStr(query);
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
        const path =
            '/users?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent(PERSON_SELECT) +
            '&$top=25';
        const data = await graphJson('GET', path, token, undefined);
        return data.value || [];
    }

    async function fetchGroupOwners(token, groupId) {
        const path =
            '/groups/' +
            encodeURIComponent(groupId) +
            '/owners?$select=' +
            encodeURIComponent(PERSON_SELECT) +
            '&$top=200';
        const owners = await fetchAllPagesSimple(token, path, 2000);
        owners.sort(function (a, b) {
            return compareDe(personLabel(a), personLabel(b));
        });
        return owners;
    }

    async function fetchGroupMemberCount(token, groupId) {
        const path = '/groups/' + encodeURIComponent(groupId) + '/members/$count';
        const res = await graphRequest('GET', path, token, undefined, { ConsistencyLevel: 'eventual' });
        const text = await res.text();
        if (!res.ok) return -1;
        const n = parseInt(String(text).trim(), 10);
        return isNaN(n) ? -1 : n;
    }

    async function fetchGroupMembers(token, groupId) {
        let next =
            '/groups/' +
            encodeURIComponent(groupId) +
            '/members?$select=' +
            encodeURIComponent(PERSON_SELECT) +
            '&$top=200';
        const out = [];
        let pages = 0;
        while (next && pages < 40 && out.length < 2000) {
            pages++;
            const data = await graphJson('GET', next, token, undefined);
            const vals = data.value || [];
            for (let i = 0; i < vals.length; i++) out.push(vals[i]);
            if (out.length >= 2000) break;
            next = data['@odata.nextLink'] || null;
        }
        out.sort(function (a, b) {
            return compareDe(personLabel(a), personLabel(b));
        });
        return { items: out, truncated: !!next || out.length >= 2000 };
    }

    async function addGroupOwner(token, groupId, userId) {
        await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/owners/$ref', token, {
            '@odata.id': directoryObjectRef(userId)
        });
    }

    async function deleteUnifiedGroup(token, groupId) {
        const id = String(groupId || '').trim();
        if (!id) throw new Error('Gruppen-ID fehlt.');
        try {
            await graphJson('DELETE', '/groups/' + encodeURIComponent(id), token, undefined);
            noteAction({
                tool: 'graph',
                action: 'delete-group',
                target: id,
                summary: 'Gruppe gelöscht'
            });
        } catch (e) {
            noteAction({
                tool: 'graph',
                action: 'delete-group',
                target: id,
                summary: String(e && e.message ? e.message : e),
                result: 'error'
            });
            throw e;
        }
    }

    async function removeGroupOwner(token, groupId, ownerId) {
        await graphJson(
            'DELETE',
            '/groups/' + encodeURIComponent(groupId) + '/owners/' + encodeURIComponent(ownerId) + '/$ref',
            token,
            undefined
        );
    }

    async function removeGroupMember(token, groupId, memberId) {
        await graphJson(
            'DELETE',
            '/groups/' + encodeURIComponent(groupId) + '/members/' + encodeURIComponent(memberId) + '/$ref',
            token,
            undefined
        );
    }

    async function addOwnerWithMemberFallback(token, groupId, userId) {
        try {
            await addGroupOwner(token, groupId, userId);
        } catch (e1) {
            try {
                await graphAddMember(token, groupId, userId);
            } catch (e2) {
                if (!isDuplicateMemberError(e2)) throw e2;
            }
            await addGroupOwner(token, groupId, userId);
        }
    }

    async function resolveUserByEmail(token, email) {
        const em = normEmail(email);
        if (!em || em.indexOf('@') === -1) return null;
        const esc = odataEscape(em);
        const filter = "(mail eq '" + esc + "' or userPrincipalName eq '" + esc + "')";
        const path =
            '/users?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent(PERSON_SELECT) +
            '&$top=5';
        const data = await graphJson('GET', path, token, undefined);
        const list = data.value || [];
        return list[0] || null;
    }

    async function resolveUserByEmailForImport(token, email) {
        const em = normEmail(email);
        if (!em || em.indexOf('@') === -1) return null;
        const esc = odataEscape(em);
        const filter = "(mail eq '" + esc + "' or userPrincipalName eq '" + esc + "')";
        const path =
            '/users?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent(USER_LICENSE_SELECT) +
            '&$top=5';
        const data = await graphJson('GET', path, token, undefined);
        const list = data.value || [];
        return list[0] || null;
    }

    async function resolveUsersByEmailsForImport(token, emails) {
        const list = Array.isArray(emails) ? emails : [];
        const seen = new Set();
        const unique = [];
        list.forEach(function (raw) {
            const em = normEmail(raw);
            if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
            seen.add(em);
            unique.push(em);
        });
        const out = [];
        for (let i = 0; i < unique.length; i++) {
            try {
                const u = await resolveUserByEmailForImport(token, unique[i]);
                if (u) out.push(u);
            } catch {
                /* Einzelner Treffer fehlgeschlagen – überspringen */
            }
        }
        return out;
    }

    async function createUnifiedGroup(token, displayName, mailNickname, description) {
        const nick = sanitizeUnifiedGroupMailNickname(mailNickname);
        const body = {
            displayName: String(displayName).trim(),
            description: description || 'MS365-Schulverwaltung – Microsoft 365-Gruppe',
            mailNickname: nick,
            mailEnabled: true,
            securityEnabled: false,
            groupTypes: ['Unified'],
            visibility: 'Private'
        };
        try {
            const group = await graphJson('POST', '/groups', token, body);
            await sleep(1500);
            noteAction({
                tool: 'graph',
                action: 'create-group',
                target: nick,
                summary: 'Gruppe angelegt: ' + String(displayName).trim()
            });
            return group;
        } catch (e) {
            noteAction({
                tool: 'graph',
                action: 'create-group',
                target: nick,
                summary: String(e && e.message ? e.message : e),
                result: 'error'
            });
            throw e;
        }
    }

    function buildPutTeamBody() {
        return {
            memberSettings: { allowCreatePrivateChannels: true, allowCreateUpdateChannels: true },
            messagingSettings: { allowUserEditMessages: true, allowUserDeleteMessages: true },
            funSettings: { allowGiphy: true, giphyContentRating: 'moderate' },
            guestSettings: { allowCreateUpdateChannels: false }
        };
    }

    async function provisionTeamForGroup(token, gid) {
        const teamUri = '/groups/' + encodeURIComponent(gid) + '/team';
        for (let i = 0; i < 8; i++) {
            try {
                await graphJson('PUT', teamUri, token, buildPutTeamBody());
                return;
            } catch (e) {
                const msg = String(e && e.message ? e.message : e);
                const looksLikeReplication = msg.indexOf('404') !== -1 || msg.indexOf('Request_ResourceNotFound') !== -1;
                if (i < 7 && looksLikeReplication) {
                    await sleep(10000);
                    token = await getGraphToken();
                    continue;
                }
                throw e;
            }
        }
    }

    async function graphAddMember(token, groupId, userId) {
        await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/members/$ref', token, {
            '@odata.id': directoryObjectRef(userId)
        });
    }

    async function ensureOwners(token, groupId, ownerEmails) {
        const emails = Array.isArray(ownerEmails) ? ownerEmails : [];
        let added = 0;
        for (let i = 0; i < emails.length; i++) {
            const em = emails[i];
            try {
                const u = await resolveUserByEmail(token, em);
                if (!u || !u.id) continue;
                try {
                    await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/owners/$ref', token, {
                        '@odata.id': userRef(u.id)
                    });
                    added++;
                } catch (e) {
                    if (!isDuplicateMemberError(e)) {
                        // ignore einzelne Fehler
                    }
                }
            } catch {
                // ignore
            }
            if ((i + 1) % 6 === 0) await sleep(120);
        }
        if (added === 0) {
            try {
                const me = await graphJson('GET', '/me', token, undefined);
                const meId = me && me.id;
                if (meId) {
                    try {
                        await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/owners/$ref', token, {
                            '@odata.id': userRef(meId)
                        });
                    } catch (e) {
                        if (!isDuplicateMemberError(e)) throw e;
                    }
                    try {
                        await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/members/$ref', token, {
                            '@odata.id': userRef(meId)
                        });
                    } catch (e) {
                        if (!isDuplicateMemberError(e)) throw e;
                    }
                }
            } catch {
                // optional
            }
        }
    }

    async function syncEmailsToGroup(token, groupId, emails, label, onLog) {
        const log =
            onLog ||
            function () {
                /* noop */
            };
        let ok = 0;
        let skip = 0;
        let fail = 0;
        for (let i = 0; i < emails.length; i++) {
            const em = emails[i];
            try {
                const u = await resolveUserByEmail(token, em);
                if (!u || !u.id) {
                    log(label + ': Kein Benutzer für ' + em, 'warn');
                    fail++;
                    continue;
                }
                try {
                    await graphAddMember(token, groupId, u.id);
                    ok++;
                    log(label + ': ' + em + ' → Mitglied', 'ok');
                } catch (e) {
                    if (isDuplicateMemberError(e)) {
                        skip++;
                        log(label + ': ' + em + ' (war schon Mitglied)', 'warn');
                    } else {
                        fail++;
                        log(label + ': ' + em + ' — ' + (e.message || e), 'err');
                    }
                }
            } catch (e) {
                fail++;
                log(label + ': ' + em + ' — ' + (e.message || e), 'err');
            }
            if ((i + 1) % 8 === 0) await sleep(120);
        }
        return { ok: ok, skip: skip, fail: fail };
    }

    async function removeEmailsFromGroup(token, groupId, emails, label, onLog) {
        const log =
            onLog ||
            function () {
                /* noop */
            };
        let ok = 0;
        let skip = 0;
        let fail = 0;
        for (let i = 0; i < emails.length; i++) {
            const em = emails[i];
            try {
                const u = await resolveUserByEmail(token, em);
                if (!u || !u.id) {
                    log(label + ': Kein Benutzer für ' + em, 'warn');
                    fail++;
                    continue;
                }
                try {
                    await removeGroupMember(token, groupId, u.id);
                    ok++;
                    log(label + ': ' + em + ' → entfernt', 'ok');
                } catch (e) {
                    const msg = String(e && e.message ? e.message : e);
                    if (msg.indexOf('404') !== -1 || msg.indexOf('Request_ResourceNotFound') !== -1) {
                        skip++;
                        log(label + ': ' + em + ' (war kein Mitglied)', 'warn');
                    } else {
                        fail++;
                        log(label + ': ' + em + ' — ' + msg, 'err');
                    }
                }
            } catch (e) {
                fail++;
                log(label + ': ' + em + ' — ' + (e.message || e), 'err');
            }
            if ((i + 1) % 8 === 0) await sleep(120);
        }
        return { ok: ok, skip: skip, fail: fail };
    }

    window.ms365GraphUnifiedGroups = {
        GRAPH_SCOPES,
        PERSON_SELECT,
        USER_LICENSE_SELECT,
        getGraphToken,
        sleep,
        graphRequest,
        graphJson,
        odataEscape,
        sanitizeMailNickname,
        sanitizeUnifiedGroupMailNickname,
        isUnifiedGroup,
        groupHasTeam,
        groupArtLabel,
        fetchTeamWebUrl,
        teamDeepLink,
        searchUnifiedGroups,
        fetchGroup,
        createUnifiedGroup,
        provisionTeamForGroup,
        graphAddMember,
        addGroupOwner,
        removeGroupOwner,
        removeGroupMember,
        addOwnerWithMemberFallback,
        deleteUnifiedGroup,
        fetchGroupOwners,
        fetchGroupMembers,
        fetchGroupMemberCount,
        searchUsers,
        personLabel,
        userRef,
        directoryObjectRef,
        isDuplicateMemberError,
        resolveUserByEmail,
        resolveUserByEmailForImport,
        resolveUsersByEmailsForImport,
        ensureOwners,
        syncEmailsToGroup,
        removeEmailsFromGroup,
        patchGroup,
        patchGroupDisplayName,
        renewGroup,
        GROUP_PHOTO_MAX_BYTES,
        groupPhotoInitials,
        fetchGroupPhotoBlob,
        setGroupPhoto,
        deleteGroupPhoto,
        setTeamPhoto,
        deleteTeamPhoto,
        syncTeamPhotoForGroup,
        teamPhotoSyncHint,
        fetchSubscribedSkus,
        skuLookupFromSubscribed,
        fetchUsersWithAssignedLicenses,
        fetchUsersByAssignedSkuIds
    };
})();
