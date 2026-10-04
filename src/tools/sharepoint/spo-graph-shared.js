(function (global) {
    'use strict';

    async function getGraphToken(scopes) {
        // Popup zuerst: Formulare bleiben erhalten (kein Seiten-Redirect).
        if (typeof global.ms365AuthAcquireTokenPopup === 'function') {
            try {
                return await global.ms365AuthAcquireTokenPopup(scopes);
            } catch (e) {
                const msg = String((e && e.message) || e || '');
                const code = String((e && (e.errorCode || e.code)) || '');
                // Nutzer hat Popup geschlossen → nicht still auf Redirect umschalten.
                if (
                    /abgebrochen|cancelled|canceled|user_cancelled/i.test(msg) ||
                    /user_cancelled/i.test(code)
                ) {
                    throw e;
                }
                // Popup blockiert o. Ä. → Redirect-Fallback (Formular wird von der Aufrufer-Seite gesichert).
            }
        }
        if (typeof global.ms365AuthAcquireToken === 'function') {
            return await global.ms365AuthAcquireToken(scopes);
        }
        throw new Error('Bitte oben rechts anmelden (MSAL-Widget nicht verfügbar).');
    }

    function sleep(ms) {
        return new Promise(function (r) {
            setTimeout(r, ms);
        });
    }

    function graphBase(version) {
        const v = version === 'beta' ? 'beta' : 'v1.0';
        return 'https://graph.microsoft.com/' + v;
    }

    async function graphRequest(method, pathOrUrl, token, body, version) {
        let url = pathOrUrl;
        if (url.indexOf('http') !== 0) {
            url = graphBase(version) + (pathOrUrl.indexOf('/') === 0 ? pathOrUrl : '/' + pathOrUrl);
        }
        let attempt = 0;
        while (true) {
            const headers = { Authorization: 'Bearer ' + token };
            if (body !== undefined) {
                headers['Content-Type'] = 'application/json';
            }
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

    async function graphJson(method, pathOrUrl, token, body, version) {
        const res = await graphRequest(method, pathOrUrl, token, body, version);
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
                typeof data === 'object' && data && data.error
                    ? (data.error.message || JSON.stringify(data.error))
                    : text || String(res.status);
            const err = new Error(method + ' ' + pathOrUrl + ': ' + msg);
            err.status = res.status;
            err.payload = data;
            throw err;
        }
        return data || {};
    }

    async function getSharePointHostname(token) {
        const sitesRead = [
            'https://graph.microsoft.com/User.Read',
            'https://graph.microsoft.com/Sites.Read.All'
        ];
        let t = token;
        if (!t) t = await getGraphToken(sitesRead);
        const root = await graphJson('GET', '/sites/root', t, undefined, 'v1.0');
        const w = root && root.webUrl ? String(root.webUrl) : '';
        if (!w) return '';
        try {
            return new URL(w).hostname;
        } catch {
            return '';
        }
    }

    /**
     * Location nach Site-Create zeigt manchmal auf *.sharepoint.com (_api/v2…),
     * nicht auf graph.microsoft.com → Graph-Token → invalidAudienceUri.
     * Operation-ID extrahieren und immer über Graph beta pollen.
     * @param {string} operationUrl
     * @returns {string}
     */
    function normalizeSiteOperationPollUrl(operationUrl) {
        const raw = String(operationUrl || '').trim();
        if (!raw) return '';

        let abs = raw;
        if (abs.indexOf('http') !== 0) {
            abs = 'https://graph.microsoft.com' + (abs.indexOf('/') === 0 ? '' : '/') + abs;
        }

        let opId = '';
        const mFn = abs.match(/getOperationStatus\s*\(\s*operationId\s*=\s*'([^']+)'\s*\)/i);
        if (mFn) opId = mFn[1];
        if (!opId) {
            const mFn2 = abs.match(/getOperationStatus\s*\(\s*operationId\s*=\s*"([^"]+)"\s*\)/i);
            if (mFn2) opId = mFn2[1];
        }
        if (!opId) {
            try {
                const u = new URL(abs);
                opId = u.searchParams.get('operationId') || u.searchParams.get('opId') || '';
            } catch {
                /* ignore */
            }
        }
        if (!opId) {
            const mPath = abs.match(/operations?\/([A-Za-z0-9_\-=+%]+)/i);
            if (mPath) opId = decodeURIComponent(mPath[1]);
        }

        if (opId) {
            return (
                'https://graph.microsoft.com/beta/sites/getOperationStatus(operationId=\'' +
                opId.replace(/'/g, '') +
                '\')'
            );
        }

        try {
            const u = new URL(abs);
            if (/\.sharepoint\.com$/i.test(u.hostname) || /sharepoint\.com$/i.test(u.hostname)) {
                /* Keine ID → Caller muss Fallback (Site-Existenz) nutzen */
                return '';
            }
        } catch {
            /* ignore */
        }
        return abs;
    }

    function isInvalidAudienceError(status, text) {
        if (status !== 401 && status !== 403) return false;
        return /invalidAudienceUri|Invalid audience|audience Uri/i.test(String(text || ''));
    }

    /**
     * Wartet, bis die Site per Graph erreichbar ist (Fallback wenn Op-URL SPO/Audience-Problem).
     * @param {string} siteWebUrl
     * @param {string} token
     */
    async function waitUntilSiteExists(siteWebUrl, token) {
        const max = 45;
        for (let i = 0; i < max; i++) {
            try {
                const site = await resolveSiteFromWebUrl(token, siteWebUrl);
                if (site && site.id) {
                    return {
                        status: 'succeeded',
                        resourceId: site.id,
                        resourceLocation: site.webUrl || siteWebUrl,
                        fallback: 'site-exists'
                    };
                }
            } catch (e) {
                const st = e && e.status;
                if (st && st !== 404 && !/itemNotFound|not found/i.test(String(e && e.message ? e.message : e))) {
                    /* andere Fehler kurz tolerieren, Site kann noch provisionieren */
                }
            }
            await sleep(2000);
        }
        throw new Error(
            'Timeout: Site noch nicht sichtbar unter ' +
                siteWebUrl +
                '. Bitte im SharePoint Admin Center prüfen – die Erstellung kann trotzdem laufen.'
        );
    }

    /**
     * @param {string} operationUrl Vollständige URL aus dem Location-Header (202)
     * @param {string} token Graph-Zugriffstoken
     * @param {{ siteWebUrl?: string }} [opts] Fallback: auf Site-Existenz pollen
     */
    async function pollRichLongRunningOperation(operationUrl, token, opts) {
        const siteWebUrl = opts && opts.siteWebUrl ? String(opts.siteWebUrl).trim() : '';
        let pollUrl = normalizeSiteOperationPollUrl(operationUrl);

        if (!pollUrl) {
            if (siteWebUrl) return await waitUntilSiteExists(siteWebUrl, token);
            throw new Error(
                'Vorgangs-URL zeigt auf SharePoint (nicht Graph) und enthält keine operationId. ' +
                    'Site ggf. im Admin Center prüfen.'
            );
        }

        const max = 45;
        for (let i = 0; i < max; i++) {
            let res;
            try {
                res = await fetch(pollUrl, {
                    method: 'GET',
                    headers: { Authorization: 'Bearer ' + token },
                    redirect: 'manual'
                });
            } catch (e) {
                if (siteWebUrl && i > 2) return await waitUntilSiteExists(siteWebUrl, token);
                throw e;
            }

            /* Cross-Host-Redirect auf SPO / opaque → Graph-Token ungültig */
            if (res.type === 'opaqueredirect' || (res.status >= 300 && res.status < 400)) {
                const loc = res.headers.get('Location') || res.headers.get('location') || '';
                const rewritten = loc ? normalizeSiteOperationPollUrl(loc) : '';
                if (rewritten && rewritten !== pollUrl) {
                    pollUrl = rewritten;
                    continue;
                }
                if (siteWebUrl) return await waitUntilSiteExists(siteWebUrl, token);
                throw new Error('Vorgang: Redirect auf nicht-Graph-URL – ' + (loc || res.status));
            }

            const text = await res.text();
            let data = null;
            if (text) {
                try {
                    data = JSON.parse(text);
                } catch {
                    data = { raw: text };
                }
            }
            if (!res.ok) {
                if (isInvalidAudienceError(res.status, text)) {
                    const again = normalizeSiteOperationPollUrl(pollUrl);
                    if (again && again !== pollUrl) {
                        pollUrl = again;
                        continue;
                    }
                    if (siteWebUrl) return await waitUntilSiteExists(siteWebUrl, token);
                }
                throw new Error('Vorgang: HTTP ' + res.status + ' – ' + (text || ''));
            }
            const status = data && (data.status || data.Status);
            const s = String(status || '').toLowerCase();
            if (s === 'succeeded' || s === 'completed' || s === 'complete') {
                return data;
            }
            if (s === 'failed' || s === 'cancelled' || s === 'canceled') {
                throw new Error(
                    'Vorgang fehlgeschlagen: ' +
                        (data && (data.error || data.resourceId)
                            ? JSON.stringify(data.error || data)
                            : JSON.stringify(data))
                );
            }
            /* notStarted, running, waiting … → weiter pollen */
            await sleep(2000);
        }
        if (siteWebUrl) {
            try {
                return await waitUntilSiteExists(siteWebUrl, token);
            } catch {
                /* unten Timeout */
            }
        }
        throw new Error('Timeout: Die Site-Erstellung dauert ungewöhnlich lange. Bitte im SharePoint Admin Center prüfen.');
    }

    /**
     * SharePoint REST: RequestDigest, dann RegisterHubSite (kann im Browser an CORS scheitern).
     * @param {string} siteWebUrl z. B. https://tenant.sharepoint.com/sites/Intranet
     * @param {string} spoToken Zugriffstoken mit Audience https://tenant.sharepoint.com/
     */
    async function registerHubSiteViaSpoRest(siteWebUrl, spoToken) {
        const origin = String(siteWebUrl || '').replace(/\/+$/, '');
        if (!origin || !spoToken) throw new Error('Site-URL oder SharePoint-Token fehlt.');
        const digest = await getSpoRequestDigest(origin, spoToken);
        const hubRes = await spoRestFetch(origin, spoToken, digest, 'POST', '/_api/site/RegisterHubSite', '');
        if (!hubRes.ok) {
            throw new Error('RegisterHubSite: ' + hubRes.status + ' ' + (hubRes.text || ''));
        }
        return hubRes.data;
    }

    /**
     * @param {string} siteWebUrl
     * @param {string} spoToken
     * @returns {Promise<string>} FormDigestValue
     */
    async function getSpoRequestDigest(siteWebUrl, spoToken) {
        const origin = String(siteWebUrl || '').replace(/\/+$/, '');
        if (!origin || !spoToken) throw new Error('Site-URL oder SharePoint-Token fehlt.');
        const ctxRes = await fetch(origin + '/_api/contextinfo', {
            method: 'POST',
            headers: {
                Accept: 'application/json;odata=nometadata',
                'Content-Type': 'application/json;odata=nometadata;charset=utf-8',
                Authorization: 'Bearer ' + spoToken
            }
        });
        const ctxText = await ctxRes.text();
        if (!ctxRes.ok) {
            throw new Error('contextinfo: ' + ctxRes.status + ' ' + (ctxText || ''));
        }
        let ctxJson = null;
        try {
            ctxJson = JSON.parse(ctxText);
        } catch {
            throw new Error('contextinfo: keine JSON-Antwort');
        }
        const digest =
            (ctxJson && ctxJson.FormDigestValue) ||
            (ctxJson &&
                ctxJson.d &&
                ctxJson.d.GetContextWebInformation &&
                ctxJson.d.GetContextWebInformation.FormDigestValue) ||
            '';
        if (!digest) throw new Error('Kein FormDigestValue erhalten.');
        return digest;
    }

    /**
     * @returns {Promise<{ ok: boolean, status: number, text: string, data: any }>}
     */
    async function spoRestFetch(siteWebUrl, spoToken, digest, method, apiPath, body) {
        const origin = String(siteWebUrl || '').replace(/\/+$/, '');
        const path = String(apiPath || '');
        const url = path.indexOf('http') === 0 ? path : origin + (path.indexOf('/') === 0 ? path : '/' + path);
        const headers = {
            Accept: 'application/json;odata=nometadata',
            Authorization: 'Bearer ' + spoToken
        };
        if (digest) headers['X-RequestDigest'] = digest;
        let payload = body;
        if (payload !== undefined && payload !== '' && typeof payload !== 'string') {
            headers['Content-Type'] = 'application/json;odata=nometadata;charset=utf-8';
            payload = JSON.stringify(payload);
        } else if (payload !== undefined && payload !== '') {
            headers['Content-Type'] = 'application/json;odata=nometadata;charset=utf-8';
        }
        const res = await fetch(url, {
            method: method || 'GET',
            headers: headers,
            body: payload === undefined || payload === '' ? undefined : payload
        });
        const text = await res.text();
        let data = null;
        if (text) {
            try {
                data = JSON.parse(text);
            } catch {
                data = { raw: text };
            }
        }
        return { ok: res.ok, status: res.status, text: text, data: data };
    }

    /**
     * Vererbung an einer Liste/Bibliothek brechen.
     * @param {boolean} [copyRoleAssignments=true]
     */
    async function spoBreakListInheritance(siteWebUrl, spoToken, digest, listTitle, copyRoleAssignments) {
        const title = String(listTitle || '').trim();
        if (!title) throw new Error('Listen-Titel fehlt.');
        const copy = copyRoleAssignments !== false;
        const api =
            "/_api/web/lists/getbytitle('" +
            title.replace(/'/g, "''") +
            "')/breakroleinheritance(copyRoleAssignments=" +
            (copy ? 'true' : 'false') +
            ',clearSubscopes=true)';
        const res = await spoRestFetch(siteWebUrl, spoToken, digest, 'POST', api, '');
        if (!res.ok && res.status !== 204) {
            throw new Error('breakroleinheritance: ' + res.status + ' ' + (res.text || ''));
        }
        return res.data;
    }

    async function spoListRoleAssignments(siteWebUrl, spoToken, digest, listTitle) {
        const title = String(listTitle || '').trim();
        const api =
            "/_api/web/lists/getbytitle('" +
            title.replace(/'/g, "''") +
            "')/roleassignments?$expand=Member,RoleDefinitionBindings";
        const res = await spoRestFetch(siteWebUrl, spoToken, digest, 'GET', api);
        if (!res.ok) throw new Error('roleassignments: ' + res.status + ' ' + (res.text || ''));
        const value = (res.data && (res.data.value || (res.data.d && res.data.d.results))) || [];
        return Array.isArray(value) ? value : [];
    }

    async function spoRemoveRoleAssignment(siteWebUrl, spoToken, digest, listTitle, principalId) {
        const title = String(listTitle || '').trim();
        const pid = Number(principalId);
        if (!pid) throw new Error('principalId fehlt.');
        const api =
            "/_api/web/lists/getbytitle('" +
            title.replace(/'/g, "''") +
            "')/roleassignments/getbyprincipalid(" +
            pid +
            ')';
        const origin = String(siteWebUrl || '').replace(/\/+$/, '');
        const del = await fetch(origin + api, {
            method: 'POST',
            headers: {
                Accept: 'application/json;odata=nometadata',
                Authorization: 'Bearer ' + spoToken,
                'X-RequestDigest': digest,
                'X-HTTP-Method': 'DELETE',
                'IF-MATCH': '*'
            }
        });
        const text = await del.text();
        if (!del.ok && del.status !== 204) {
            throw new Error('remove roleassignment: ' + del.status + ' ' + text);
        }
        return true;
    }

    /**
     * EnsureUser: Entra-Gruppe oder User → PrincipalId.
     * @param {string} logonName z. B. c:0o.c|federateddirectoryclaimprovider|{guid}
     */
    async function spoEnsureUser(siteWebUrl, spoToken, digest, logonName) {
        const login = String(logonName || '').trim();
        if (!login) throw new Error('logonName fehlt.');
        const res = await spoRestFetch(siteWebUrl, spoToken, digest, 'POST', '/_api/web/ensureuser', {
            logonName: login
        });
        if (!res.ok) throw new Error('ensureuser: ' + res.status + ' ' + (res.text || ''));
        const d = res.data || {};
        const id = d.Id != null ? d.Id : d.d && d.d.Id;
        if (!id) throw new Error('ensureuser: keine Id in der Antwort.');
        return { id: Number(id), title: d.Title || (d.d && d.d.Title) || '', loginName: d.LoginName || login };
    }

    async function spoAddRoleAssignment(siteWebUrl, spoToken, digest, listTitle, principalId, roleDefId) {
        const title = String(listTitle || '').trim();
        const pid = Number(principalId);
        const rid = Number(roleDefId);
        if (!pid || !rid) throw new Error('principalId/roleDefId fehlen.');
        const api =
            "/_api/web/lists/getbytitle('" +
            title.replace(/'/g, "''") +
            "')/roleassignments/addroleassignment(principalid=" +
            pid +
            ',roledefid=' +
            rid +
            ')';
        const res = await spoRestFetch(siteWebUrl, spoToken, digest, 'POST', api, '');
        if (!res.ok && res.status !== 204) {
            throw new Error('addroleassignment: ' + res.status + ' ' + (res.text || ''));
        }
        return true;
    }

    /**
     * Dokumentbibliothek anlegen (BaseTemplate 101) – oft zuverlässiger als Graph POST /lists.
     * @returns {Promise<{ Id?: string, Title?: string, id?: string }>}
     */
    async function spoCreateDocumentLibrary(siteWebUrl, spoToken, digest, title, description) {
        const name = String(title || '').trim();
        if (!name) throw new Error('Bibliotheksname fehlt.');
        const res = await spoRestFetch(siteWebUrl, spoToken, digest, 'POST', '/_api/web/lists', {
            Title: name,
            Description: String(description || ''),
            BaseTemplate: 101,
            AllowContentTypes: true,
            ContentTypesEnabled: false
        });
        if (!res.ok) {
            const msg =
                (res.data && (res.data.error_description || res.data['odata.error'] || res.data.error)) ||
                res.text ||
                String(res.status);
            throw new Error('Bibliothek anlegen (SPO REST): ' + (typeof msg === 'string' ? msg : JSON.stringify(msg)));
        }
        const d = res.data || {};
        return {
            Id: d.Id || (d.d && d.d.Id) || '',
            Title: d.Title || (d.d && d.d.Title) || name,
            id: d.Id || (d.d && d.d.Id) || ''
        };
    }

    async function spoGetListByTitle(siteWebUrl, spoToken, digest, title) {
        const name = String(title || '').trim();
        const api =
            "/_api/web/lists/getbytitle('" +
            name.replace(/'/g, "''") +
            "')?$select=Id,Title,BaseTemplate,RootFolder/ServerRelativeUrl&$expand=RootFolder";
        const res = await spoRestFetch(siteWebUrl, spoToken, digest, 'GET', api);
        if (!res.ok) {
            if (res.status === 404) return null;
            const text = String(res.text || '');
            if (/not exist|nicht vorhanden|ListDoesNotExist/i.test(text)) return null;
            throw new Error('getbytitle: ' + res.status + ' ' + text);
        }
        const d = res.data || {};
        return {
            Id: d.Id || (d.d && d.d.Id) || '',
            Title: d.Title || (d.d && d.d.Title) || name,
            id: d.Id || (d.d && d.d.Id) || ''
        };
    }

    /**
     * Listenansicht anlegen, falls Titel noch nicht existiert (Gruppierung + Sortierung).
     * @param {{ title: string, groupBy?: string, orderBy?: string, fields?: string[], rowLimit?: number }} def
     */
    async function spoEnsureListView(siteWebUrl, spoToken, digest, listTitle, def) {
        const listName = String(listTitle || '').trim();
        const viewTitle = String((def && def.title) || '').trim();
        if (!listName || !viewTitle) throw new Error('Listen- oder Ansichtstitel fehlt.');
        const esc = listName.replace(/'/g, "''");
        const existing = await spoRestFetch(
            siteWebUrl,
            spoToken,
            digest,
            'GET',
            "/_api/web/lists/getbytitle('" + esc + "')/views?$select=Title,Id"
        );
        if (!existing.ok) {
            throw new Error('views lesen: ' + existing.status + ' ' + (existing.text || ''));
        }
        const views =
            (existing.data && (existing.data.value || (existing.data.d && existing.data.d.results))) || [];
        const hit = (Array.isArray(views) ? views : []).find(function (v) {
            return String((v && v.Title) || '').toLowerCase() === viewTitle.toLowerCase();
        });
        if (hit) return { created: false, id: hit.Id || hit.id || '', title: viewTitle };

        const groupBy = String((def && def.groupBy) || '').trim();
        const orderBy = String((def && def.orderBy) || 'Nachname').trim();
        let viewQuery = '';
        if (groupBy) {
            viewQuery +=
                '<GroupBy Collapse="TRUE" GroupLimit="100"><FieldRef Name="' +
                groupBy +
                '" /></GroupBy>';
        }
        if (orderBy) {
            viewQuery +=
                '<OrderBy><FieldRef Name="' + orderBy + '" Ascending="TRUE" /></OrderBy>';
        }
        const body = {
            Title: viewTitle,
            PersonalView: false,
            ViewType: 'HTML',
            RowLimit: Number((def && def.rowLimit) || 100) || 100,
            ViewQuery: viewQuery
        };

        const res = await spoRestFetch(
            siteWebUrl,
            spoToken,
            digest,
            'POST',
            "/_api/web/lists/getbytitle('" + esc + "')/views",
            body
        );
        if (!res.ok) {
            throw new Error(
                'view anlegen „' + viewTitle + '": ' + res.status + ' ' + (res.text || '')
            );
        }
        const d = res.data || {};
        return {
            created: true,
            id: d.Id || (d.d && d.d.Id) || '',
            title: viewTitle
        };
    }

    /**
     * Beliebiges Listenfeld per MERGE patchen (ConditionalShowFormula, CustomFormatter, Required, …).
     * @param {Record<string, unknown>} body
     */
    async function spoPatchListField(siteWebUrl, spoToken, digest, listTitle, fieldInternalName, body) {
        const listName = String(listTitle || '').trim();
        const field = String(fieldInternalName || '').trim();
        if (!listName || !field) throw new Error('Liste oder Feldname fehlt.');
        const esc = listName.replace(/'/g, "''");
        const origin = String(siteWebUrl || '').replace(/\/+$/, '');
        const api =
            "/_api/web/lists/getbytitle('" +
            esc +
            "')/fields/getbyinternalnameortitle('" +
            field.replace(/'/g, "''") +
            "')";
        const payload = body && typeof body === 'object' ? body : {};
        const res = await fetch(origin + api, {
            method: 'POST',
            headers: {
                Accept: 'application/json;odata=nometadata',
                Authorization: 'Bearer ' + spoToken,
                'X-RequestDigest': digest,
                'Content-Type': 'application/json;odata=nometadata;charset=utf-8',
                'X-HTTP-Method': 'MERGE',
                'IF-MATCH': '*'
            },
            body: JSON.stringify(payload)
        });
        const text = await res.text();
        if (!res.ok && res.status !== 204) {
            throw new Error('Feld „' + field + '": ' + res.status + ' ' + text);
        }
        return true;
    }

    /**
     * ClientFormCustomFormatter am List-ContentType „Item“/„Element“.
     * @param {string} formatterJsonString
     */
    async function spoSetListClientFormCustomFormatter(
        siteWebUrl,
        spoToken,
        digest,
        listTitle,
        formatterJsonString
    ) {
        const listName = String(listTitle || '').trim();
        if (!listName) throw new Error('Listen-Titel fehlt.');
        const esc = listName.replace(/'/g, "''");
        const ctRes = await spoRestFetch(
            siteWebUrl,
            spoToken,
            digest,
            'GET',
            "/_api/web/lists/getbytitle('" + esc + "')/contenttypes?$select=StringId,Name"
        );
        if (!ctRes.ok) {
            throw new Error('contenttypes: ' + ctRes.status + ' ' + (ctRes.text || ''));
        }
        const cts =
            (ctRes.data && (ctRes.data.value || (ctRes.data.d && ctRes.data.d.results))) || [];
        const list = Array.isArray(cts) ? cts : [];
        let ct =
            list.find(function (c) {
                const n = String((c && c.Name) || '').toLowerCase();
                return n === 'item' || n === 'element';
            }) ||
            list.find(function (c) {
                const id = String((c && c.StringId) || '');
                return /^0x01/i.test(id) && !/^0x0120/i.test(id);
            });
        if (!ct || !ct.StringId) throw new Error('Content Type Item/Element nicht gefunden.');
        const origin = String(siteWebUrl || '').replace(/\/+$/, '');
        const api =
            "/_api/web/lists/getbytitle('" +
            esc +
            "')/contenttypes('" +
            String(ct.StringId).replace(/'/g, "''") +
            "')";
        const res = await fetch(origin + api, {
            method: 'POST',
            headers: {
                Accept: 'application/json;odata=verbose',
                Authorization: 'Bearer ' + spoToken,
                'X-RequestDigest': digest,
                'Content-Type': 'application/json;odata=verbose;charset=utf-8',
                'X-HTTP-Method': 'MERGE',
                'IF-MATCH': '*'
            },
            body: JSON.stringify({
                __metadata: { type: 'SP.ContentType' },
                ClientFormCustomFormatter: String(formatterJsonString || '')
            })
        });
        const text = await res.text();
        if (!res.ok && res.status !== 204) {
            throw new Error('ClientFormCustomFormatter: ' + res.status + ' ' + text);
        }
        return { stringId: ct.StringId, name: ct.Name || '' };
    }

    /**
     * Title-Feld: Anzeigename + optional nicht mehr pflichtig (Listenformular nutzt Nachname/Vorname).
     */
    async function spoPatchTitleField(siteWebUrl, spoToken, digest, listTitle, opts) {
        const o = opts || {};
        return spoPatchListField(siteWebUrl, spoToken, digest, listTitle, 'Title', {
            Title: String(o.displayName || 'Name (Nachname, Vorname)'),
            Description: String(
                o.description ||
                    'Anzeige aus Nachname und Vorname (Formular blendet dieses Feld aus).'
            ),
            Required: o.required === true
        });
    }

    /**
     * @returns {{ host: string, serverRelativeUrl: string } | null}
     */
    function parseSharePointWebUrl(input) {
        let raw = String(input || '').trim();
        if (!raw) return null;
        if (!/^https?:\/\//i.test(raw)) raw = 'https://' + raw;
        let u;
        try {
            u = new URL(raw);
        } catch {
            return null;
        }
        const host = String(u.hostname || '')
            .trim()
            .toLowerCase();
        if (!host) return null;
        let pathname = u.pathname || '';
        if (pathname.length > 1 && pathname.endsWith('/')) pathname = pathname.slice(0, -1);
        if (!pathname || pathname === '') pathname = '/';
        return { host: host, serverRelativeUrl: pathname };
    }

    /**
     * GET /sites/{hostname}:/{path} – liefert Site-Ressource inkl. id.
     */
    async function resolveSiteFromWebUrl(token, webUrlInput) {
        const parts = parseSharePointWebUrl(webUrlInput);
        if (!parts) throw new Error('Ungültige SharePoint-Website-URL.');
        const hostPath = parts.host + ':' + parts.serverRelativeUrl;
        const seg = encodeURIComponent(hostPath);
        return await graphJson('GET', '/sites/' + seg, token, undefined, 'v1.0');
    }

    function graphPathSite(siteId) {
        return '/sites/' + encodeURIComponent(siteId);
    }

    function normEmailKey(email) {
        return String(email || '').trim().toLowerCase();
    }

    function membershipLogonName(email) {
        const em = normEmailKey(email);
        if (!em || em.indexOf('@') === -1) return '';
        return 'i:0#.f|membership|' + em;
    }

    /**
     * Systemliste „User Information List“ (interner Listenname users, lokalisiertes displayName).
     * Graph liefert sie nur mit $select=…,system; Erkennung über name, nicht nur englisches displayName.
     * @returns {Promise<object|null>}
     */
    async function findSiteUserInformationList(token, siteId) {
        let path = graphPathSite(siteId) + '/lists?$select=id,displayName,name,system&$top=999';
        while (path) {
            const data = await graphJson('GET', path, token, undefined, 'v1.0');
            const lists = (data && data.value) || [];
            const userList =
                lists.find(function (l) {
                    return String(l.name || '').trim().toLowerCase() === 'users';
                }) ||
                lists.find(function (l) {
                    return String(l.displayName || '') === 'User Information List';
                }) ||
                lists.find(function (l) {
                    return /user information/i.test(String(l.displayName || ''));
                }) ||
                lists.find(function (l) {
                    return /benutzerinformationsliste/i.test(String(l.displayName || ''));
                }) ||
                null;
            if (userList && userList.id) return userList;
            path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
        }
        return null;
    }

    /**
     * Lädt die User Information List und baut E-Mail → Listenelement-ID (LookupId).
     * @returns {Promise<Map<string, string>>}
     */
    async function loadSiteUserInfoEmailIndex(token, siteId) {
        const userList = await findSiteUserInformationList(token, siteId);
        if (!userList || !userList.id) {
            throw new Error('User Information List auf dieser Site nicht gefunden.');
        }
        const index = new Map();
        let path =
            graphPathSite(siteId) +
            '/lists/' +
            encodeURIComponent(userList.id) +
            '/items?$expand=fields&$top=200';
        while (path) {
            const data = await graphJson('GET', path, token, undefined, 'v1.0');
            const rows = (data && data.value) || [];
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const f = (row && row.fields) || {};
                const id = row && row.id != null ? String(row.id) : '';
                if (!id) continue;
                const candidates = [f.EMail, f.Email, f.UserName, f.SipAddress, f.Name]
                    .map(function (x) {
                        return String(x || '').trim().toLowerCase();
                    })
                    .filter(Boolean);
                for (let c = 0; c < candidates.length; c++) {
                    const raw = candidates[c];
                    if (raw.indexOf('@') === -1) continue;
                    if (!index.has(raw)) index.set(raw, id);
                    const pipe = raw.indexOf('|');
                    if (pipe !== -1) {
                        const tail = raw.slice(pipe + 1);
                        if (tail.indexOf('@') !== -1 && !index.has(tail)) index.set(tail, id);
                    }
                }
            }
            path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
        }
        return index;
    }

    /**
     * @param {object} fields
     * @param {string} columnName interner Spaltenname (ohne LookupId)
     * @param {number[]} lookupIds SharePoint User Information LookupIds
     */
    function applyMultiPersonLookupFields(fields, columnName, lookupIds) {
        const name = String(columnName || '').trim();
        const nums = (Array.isArray(lookupIds) ? lookupIds : [])
            .map(function (x) {
                return Number(x);
            })
            .filter(function (n) {
                return n > 0;
            });
        if (!name || !nums.length) return fields;
        fields[name + 'LookupId@odata.type'] = 'Collection(Edm.Int32)';
        fields[name + 'LookupId'] = nums;
        return fields;
    }

    function applySinglePersonLookupField(fields, columnName, lookupId) {
        const name = String(columnName || '').trim();
        const id = Number(lookupId);
        if (!name || !id) return fields;
        fields[name + 'LookupId'] = id;
        return fields;
    }

    /**
     * Personen-LookupIds für eine Site (Index + optional ensureUser).
     * @param {string} siteWebUrl
     * @param {string} graphToken
     * @param {string} siteId
     * @param {{ ensureMissing?: boolean, write?: function }} [opts]
     */
    async function createSitePersonResolver(siteWebUrl, graphToken, siteId, opts) {
        const write = opts && typeof opts.write === 'function' ? opts.write : function () {};
        const ensureMissing = !(opts && opts.ensureMissing === false);
        const index = await loadSiteUserInfoEmailIndex(graphToken, siteId);
        let spoToken = '';
        let digest = '';
        if (ensureMissing) {
            let host = '';
            try {
                host = new URL(siteWebUrl).hostname;
            } catch {
                host = '';
            }
            if (host) {
                const spoScope = 'https://' + host + '/AllSites.Write';
                try {
                    spoToken = await getGraphToken([spoScope]);
                    digest = await getSpoRequestDigest(siteWebUrl, spoToken);
                } catch (e) {
                    write(
                        'Hinweis: ensureUser nicht verfügbar (' +
                            (e && e.message ? e.message : e) +
                            ') – nur bereits bekannte Site-Benutzer werden verknüpft.'
                    );
                    spoToken = '';
                }
            }
        }

        async function lookupIdForEmail(email) {
            const em = normEmailKey(email);
            if (!em) return 0;
            if (index.has(em)) return Number(index.get(em));
            if (!ensureMissing || !spoToken || !digest) return 0;
            const logon = membershipLogonName(em);
            if (!logon) return 0;
            try {
                const principal = await spoEnsureUser(siteWebUrl, spoToken, digest, logon);
                const id = principal && principal.id ? Number(principal.id) : 0;
                if (id) index.set(em, String(id));
                return id;
            } catch (e) {
                write('ensureUser fehlgeschlagen für ' + em + ': ' + (e && e.message ? e.message : e));
                return 0;
            }
        }

        return { lookupIdForEmail: lookupIdForEmail, index: index };
    }

    global.ms365SpoGraph = {
        getGraphToken: getGraphToken,
        sleep: sleep,
        graphRequest: graphRequest,
        graphJson: graphJson,
        getSharePointHostname: getSharePointHostname,
        pollRichLongRunningOperation: pollRichLongRunningOperation,
        registerHubSiteViaSpoRest: registerHubSiteViaSpoRest,
        getSpoRequestDigest: getSpoRequestDigest,
        spoRestFetch: spoRestFetch,
        spoBreakListInheritance: spoBreakListInheritance,
        spoListRoleAssignments: spoListRoleAssignments,
        spoRemoveRoleAssignment: spoRemoveRoleAssignment,
        spoEnsureUser: spoEnsureUser,
        spoAddRoleAssignment: spoAddRoleAssignment,
        spoCreateDocumentLibrary: spoCreateDocumentLibrary,
        spoGetListByTitle: spoGetListByTitle,
        spoEnsureListView: spoEnsureListView,
        spoPatchTitleField: spoPatchTitleField,
        spoPatchListField: spoPatchListField,
        spoSetListClientFormCustomFormatter: spoSetListClientFormCustomFormatter,
        graphBase: graphBase,
        parseSharePointWebUrl: parseSharePointWebUrl,
        resolveSiteFromWebUrl: resolveSiteFromWebUrl,
        graphPathSite: graphPathSite,
        membershipLogonName: membershipLogonName,
        findSiteUserInformationList: findSiteUserInformationList,
        loadSiteUserInfoEmailIndex: loadSiteUserInfoEmailIndex,
        applyMultiPersonLookupFields: applyMultiPersonLookupFields,
        applySinglePersonLookupField: applySinglePersonLookupField,
        createSitePersonResolver: createSitePersonResolver
    };
})(typeof window !== 'undefined' ? window : globalThis);
