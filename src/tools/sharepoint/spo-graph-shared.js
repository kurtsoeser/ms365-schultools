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
     * @param {string} operationUrl Vollständige URL aus dem Location-Header (202)
     * @param {string} token Graph-Zugriffstoken
     */
    async function pollRichLongRunningOperation(operationUrl, token) {
        const max = 45;
        for (let i = 0; i < max; i++) {
            const res = await fetch(operationUrl, {
                method: 'GET',
                headers: { Authorization: 'Bearer ' + token }
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
            if (!res.ok) {
                throw new Error('Vorgang: HTTP ' + res.status + ' – ' + (text || ''));
            }
            const status = data && (data.status || data.Status);
            const s = String(status || '').toLowerCase();
            if (s === 'succeeded' || s === 'completed' || s === 'complete') {
                return data;
            }
            if (s === 'failed' || s === 'cancelled' || s === 'canceled') {
                throw new Error('Vorgang fehlgeschlagen: ' + (data && (data.error || data.resourceId) ? JSON.stringify(data.error || data) : JSON.stringify(data)));
            }
            /* notStarted, running, waiting … → weiter pollen */
            await sleep(2000);
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
        graphPathSite: graphPathSite
    };
})(typeof window !== 'undefined' ? window : globalThis);
