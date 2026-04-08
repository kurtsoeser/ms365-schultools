(function () {
    'use strict';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/Group.ReadWrite.All'
    ];

    const GROUP_LIST_SELECT =
        'id,displayName,mail,mailNickname,createdDateTime,groupTypes,resourceProvisioningOptions';
    const PERSON_SELECT = 'id,displayName,mail,userPrincipalName';
    const MEMBERS_FETCH_TOP = 200;
    const MAX_MEMBER_NAMES_SHOWN = 40;

    let msalMod = null;
    let pca = null;
    /** @type {{ id: string, displayName: string, mail: string, created: string, createdTs: number, owners: string[], ownerKey: string, members: string[], memberCount: number, memberCountIsEstimate: boolean, memberTruncated: boolean }[]} */
    let loadedRows = [];
    /** @type {string | null} */
    let selectedGroupId = null;
    /** @type {{ key: string, dir: 'asc' | 'desc' }} */
    let sortState = { key: 'displayName', dir: 'asc' };

    const SORT_LABELS = {
        displayName: 'Anzeigename',
        mail: 'E-Mail',
        id: 'Gruppen-ID',
        createdTs: 'Erstellt',
        ownerKey: 'Besitzer',
        memberCount: 'Mitglieder'
    };

    function toast(msg) {
        const el = document.getElementById('toast');
        if (el) {
            el.textContent = msg;
            el.classList.add('show');
            clearTimeout(toast._t);
            toast._t = setTimeout(() => el.classList.remove('show'), 3800);
        } else {
            window.alert(msg);
        }
    }

    async function loadMsal() {
        if (msalMod) return msalMod;
        try {
            msalMod = await import('https://esm.sh/@azure/msal-browser@3.26.1');
        } catch {
            msalMod = await import('https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.26.1/+esm');
        }
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

    async function graphRequest(method, pathOrUrl, token, body, extraHeaders) {
        const url =
            pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
        let attempt = 0;
        while (true) {
            const headers = { Authorization: 'Bearer ' + token };
            if (extraHeaders && typeof extraHeaders === 'object') {
                Object.assign(headers, extraHeaders);
            }
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

    async function graphJson(method, pathOrUrl, token, body) {
        const res = await graphRequest(method, pathOrUrl, token, body);
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
                    ? JSON.stringify(data.error)
                    : text || String(res.status);
            throw new Error(method + ' ' + pathOrUrl + ': ' + msg);
        }
        return data || {};
    }

    function appendLog(msg, kind) {
        const el = document.getElementById('guLog');
        if (!el) return;
        const line = document.createElement('div');
        line.textContent = new Date().toLocaleTimeString() + '  ' + msg;
        if (kind === 'err') line.style.color = '#b00020';
        else if (kind === 'ok') line.style.color = '#0d8050';
        else if (kind === 'warn') line.style.color = '#856404';
        else line.style.color = '#212529';
        el.appendChild(line);
        el.scrollTop = el.scrollHeight;
    }

    function clearLog() {
        const el = document.getElementById('guLog');
        if (el) el.replaceChildren();
    }

    function formatPerson(o) {
        if (!o || typeof o !== 'object') return '';
        const dn = o.displayName ? String(o.displayName).trim() : '';
        if (dn) return dn;
        const mail = o.mail || o.userPrincipalName;
        if (mail) return String(mail).trim();
        return o.id ? String(o.id) : '';
    }

    function groupHasTeamProvisioning(g) {
        const opts = g && g.resourceProvisioningOptions;
        return Array.isArray(opts) && opts.indexOf('Team') !== -1;
    }

    async function fetchAllPages(token, initialPath) {
        const out = [];
        let next = initialPath;
        while (next) {
            const data = await graphJson('GET', next, token, undefined);
            const vals = data.value;
            if (Array.isArray(vals)) {
                for (let i = 0; i < vals.length; i++) out.push(vals[i]);
            }
            next = data['@odata.nextLink'] || null;
        }
        return out;
    }

    async function fetchOwnersForGroup(token, groupId) {
        const path =
            '/groups/' +
            encodeURIComponent(groupId) +
            '/owners?$select=' +
            encodeURIComponent(PERSON_SELECT);
        return fetchAllPages(token, path);
    }

    /**
     * Gesamtanzahl Mitglieder (schnell, unabhängig von der Namensliste).
     */
    async function fetchMemberCount(token, groupId) {
        const path = '/groups/' + encodeURIComponent(groupId) + '/members/$count';
        const res = await graphRequest('GET', path, token, undefined, { ConsistencyLevel: 'eventual' });
        const text = await res.text();
        if (!res.ok) {
            return -1;
        }
        const n = parseInt(String(text).trim(), 10);
        return isNaN(n) ? -1 : n;
    }

    /**
     * Erste Seite Mitglieder für Anzeigenamen (große Gruppen: nicht alle Zeilen laden).
     */
    async function fetchMembersFirstPage(token, groupId) {
        const path =
            '/groups/' +
            encodeURIComponent(groupId) +
            '/members?$select=' +
            encodeURIComponent(PERSON_SELECT) +
            '&$top=' +
            MEMBERS_FETCH_TOP;
        const data = await graphJson('GET', path, token, undefined);
        return { items: data.value || [], hasMore: !!data['@odata.nextLink'] };
    }

    async function mapWithConcurrency(items, limit, fn) {
        const results = new Array(items.length);
        let i = 0;
        async function worker() {
            while (i < items.length) {
                const idx = i++;
                results[idx] = await fn(items[idx], idx);
            }
        }
        const workers = [];
        const n = Math.min(limit, items.length || 1);
        for (let w = 0; w < n; w++) workers.push(worker());
        await Promise.all(workers);
        return results;
    }

    function formatDate(iso) {
        if (!iso) return '–';
        try {
            const d = new Date(iso);
            if (isNaN(d.getTime())) return String(iso);
            return d.toLocaleString(undefined, {
                dateStyle: 'medium',
                timeStyle: 'short'
            });
        } catch {
            return String(iso);
        }
    }

    function buildRowObjectFromMeta(meta, ownersRaw, memberCountExact, membersPage) {
        const ownerLabels = [];
        for (let oi = 0; oi < ownersRaw.length; oi++) {
            const lab = formatPerson(ownersRaw[oi]);
            if (lab) ownerLabels.push(lab);
        }
        ownerLabels.sort(function (a, b) {
            return a.localeCompare(b, 'de');
        });

        const membersRaw = membersPage.items;
        const memberLabels = [];
        for (let mi = 0; mi < membersRaw.length; mi++) {
            const lab = formatPerson(membersRaw[mi]);
            if (lab) memberLabels.push(lab);
        }
        memberLabels.sort(function (a, b) {
            return a.localeCompare(b, 'de');
        });

        let totalMembers = memberCountExact >= 0 ? memberCountExact : membersRaw.length;
        const countIsEstimate = memberCountExact < 0 && membersPage.hasMore;

        let truncated = membersPage.hasMore || memberLabels.length > MAX_MEMBER_NAMES_SHOWN;
        let shown = memberLabels;
        if (memberLabels.length > MAX_MEMBER_NAMES_SHOWN) {
            shown = memberLabels.slice(0, MAX_MEMBER_NAMES_SHOWN);
        }

        return {
            id: meta.id,
            displayName: meta.displayName,
            mail: meta.mail,
            created: meta.created,
            createdTs: meta.createdTs,
            owners: ownerLabels,
            ownerKey: ownerLabels.length ? ownerLabels[0] : '',
            members: shown,
            memberCount: totalMembers,
            memberCountIsEstimate: countIsEstimate,
            memberTruncated: truncated
        };
    }

    function odataEscape(s) {
        return String(s).replace(/'/g, "''");
    }

    function userRef(userId) {
        return 'https://graph.microsoft.com/v1.0/users/' + userId;
    }

    async function graphSearchUsers(token, query) {
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
        const path =
            '/users?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent(PERSON_SELECT) +
            '&$top=25';
        const data = await graphJson('GET', path, token, undefined);
        return data.value || [];
    }

    async function graphAddGroupMember(token, groupId, userId) {
        const body = { '@odata.id': userRef(userId) };
        await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/members/$ref', token, body);
    }

    async function graphRemoveGroupMember(token, groupId, memberId) {
        await graphJson(
            'DELETE',
            '/groups/' + encodeURIComponent(groupId) + '/members/' + encodeURIComponent(memberId) + '/$ref',
            token,
            undefined
        );
    }

    async function graphAddGroupOwner(token, groupId, userId) {
        const body = { '@odata.id': userRef(userId) };
        await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/owners/$ref', token, body);
    }

    async function graphRemoveGroupOwner(token, groupId, userId) {
        await graphJson(
            'DELETE',
            '/groups/' + encodeURIComponent(groupId) + '/owners/' + encodeURIComponent(userId) + '/$ref',
            token,
            undefined
        );
    }

    /**
     * Besitzer müssen oft bereits Mitglied sein – bei Fehler zuerst Mitglied, dann Besitzer.
     */
    async function graphAddOwnerWithMemberFallback(token, groupId, userId) {
        try {
            await graphAddGroupOwner(token, groupId, userId);
            return;
        } catch (e1) {
            try {
                await graphAddGroupMember(token, groupId, userId);
            } catch (e2) {
                const m2 = e2 && e2.message ? e2.message : String(e2);
                if (m2.indexOf('added object references already exist') === -1) {
                    throw e2;
                }
            }
            await graphAddGroupOwner(token, groupId, userId);
        }
    }

    async function refreshLoadedRow(token, groupId) {
        const idx = loadedRows.findIndex(function (r) {
            return r.id === groupId;
        });
        if (idx === -1) return;
        const meta = loadedRows[idx];
        const ownersRaw = await fetchOwnersForGroup(token, groupId);
        const memberCountExact = await fetchMemberCount(token, groupId);
        const membersPage = await fetchMembersFirstPage(token, groupId);
        loadedRows[idx] = buildRowObjectFromMeta(
            {
                id: meta.id,
                displayName: meta.displayName,
                mail: meta.mail,
                created: meta.created,
                createdTs: meta.createdTs
            },
            ownersRaw,
            memberCountExact,
            membersPage
        );
        renderTable();
    }

    async function fetchMembersForManage(token, groupId) {
        const out = [];
        let next =
            '/groups/' +
            encodeURIComponent(groupId) +
            '/members?$select=' +
            encodeURIComponent(PERSON_SELECT) +
            '&$top=200';
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
            return compareStrings(formatPerson(a), formatPerson(b));
        });
        return { items: out, truncated: !!next };
    }

    function renderManageLists(ownersRaw, membersResult) {
        const ownEl = document.getElementById('guOwnerList');
        const memEl = document.getElementById('guMemberList');
        if (ownEl) ownEl.replaceChildren();
        if (memEl) memEl.replaceChildren();

        if (ownEl) {
            for (let i = 0; i < ownersRaw.length; i++) {
                const o = ownersRaw[i];
                const row = document.createElement('div');
                row.style.display = 'flex';
                row.style.justifyContent = 'space-between';
                row.style.alignItems = 'flex-start';
                row.style.gap = '10px';
                row.style.padding = '6px 0';
                row.style.borderBottom = '1px solid #e9ecef';
                const txt = document.createElement('div');
                txt.style.lineHeight = '1.35';
                txt.style.fontSize = '0.92em';
                const line1 = formatPerson(o) || '–';
                const line2 = (o.userPrincipalName || o.mail || '').trim();
                txt.textContent = line1 + (line2 ? '\n' + line2 : '');
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn';
                btn.style.padding = '4px 10px';
                btn.style.fontSize = '0.8em';
                btn.textContent = 'Entfernen';
                btn.dataset.guRemoveOwner = o.id || '';
                row.appendChild(txt);
                row.appendChild(btn);
                ownEl.appendChild(row);
            }
            if (!ownersRaw.length) {
                const p = document.createElement('p');
                p.style.margin = '0';
                p.style.color = '#6c757d';
                p.textContent = 'Keine Besitzer.';
                ownEl.appendChild(p);
            }
        }

        const memItems = membersResult.items || [];
        if (memEl) {
            if (membersResult.truncated) {
                const note = document.createElement('p');
                note.style.margin = '0 0 8px';
                note.style.fontSize = '0.85em';
                note.style.color = '#856404';
                note.textContent =
                    'Hinweis: Es werden nur die ersten 2000 Mitglieder für diese Liste geladen (Performance).';
                memEl.appendChild(note);
            }
            for (let j = 0; j < memItems.length; j++) {
                const m = memItems[j];
                const row = document.createElement('div');
                row.style.display = 'flex';
                row.style.justifyContent = 'space-between';
                row.style.alignItems = 'flex-start';
                row.style.gap = '10px';
                row.style.padding = '6px 0';
                row.style.borderBottom = '1px solid #e9ecef';
                const txt = document.createElement('div');
                txt.style.lineHeight = '1.35';
                txt.style.fontSize = '0.92em';
                const t = m && m['@odata.type'] ? String(m['@odata.type']) : '';
                const typeHint =
                    t.indexOf('group') !== -1 ? ' (Gruppe)' : t.indexOf('user') !== -1 ? '' : ' (Objekt)';
                const line1 = (formatPerson(m) || '–') + typeHint;
                const line2 = (m.userPrincipalName || m.mail || '').trim();
                txt.textContent = line1 + (line2 ? '\n' + line2 : '');
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn';
                btn.style.padding = '4px 10px';
                btn.style.fontSize = '0.8em';
                btn.textContent = 'Entfernen';
                btn.dataset.guRemoveMember = m.id || '';
                row.appendChild(txt);
                row.appendChild(btn);
                memEl.appendChild(row);
            }
            if (!memItems.length) {
                const p = document.createElement('p');
                p.style.margin = '0';
                p.style.color = '#6c757d';
                p.textContent = 'Keine Mitglieder.';
                memEl.appendChild(p);
            }
        }
    }

    function setManageBusy(busy) {
        const ids = ['guBtnUserSearch', 'guBtnAddOwner', 'guBtnAddMember', 'guBtnRefreshManage'];
        for (let i = 0; i < ids.length; i++) {
            const el = document.getElementById(ids[i]);
            if (el) el.disabled = !!busy;
        }
    }

    async function loadManageDetail() {
        if (!selectedGroupId) return;
        const hint = document.getElementById('guManageHint');
        const panel = document.getElementById('guManagePanel');
        const title = document.getElementById('guManageTitle');
        const idEl = document.getElementById('guManageId');
        if (hint) hint.style.display = 'none';
        if (panel) panel.style.display = '';
        const row = loadedRows.find(function (r) {
            return r.id === selectedGroupId;
        });
        if (title) title.textContent = row ? row.displayName || '(ohne Name)' : '';
        if (idEl) idEl.textContent = selectedGroupId || '';

        setManageBusy(true);
        try {
            const token = await getGraphToken();
            const ownersRaw = await fetchOwnersForGroup(token, selectedGroupId);
            const membersResult = await fetchMembersForManage(token, selectedGroupId);
            renderManageLists(ownersRaw, membersResult);
        } catch (e) {
            appendLog('Verwaltung: ' + (e && e.message ? e.message : String(e)), 'err');
            toast('Verwaltung: ' + (e && e.message ? e.message : String(e)));
        } finally {
            setManageBusy(false);
        }
    }

    function selectGroup(groupId) {
        selectedGroupId = groupId || null;
        const sel = document.getElementById('guUserSearchResults');
        if (sel) sel.replaceChildren();
        const searchInp = document.getElementById('guUserSearch');
        if (searchInp) searchInp.value = '';
        if (!selectedGroupId) {
            const hint = document.getElementById('guManageHint');
            const panel = document.getElementById('guManagePanel');
            if (hint) hint.style.display = '';
            if (panel) panel.style.display = 'none';
            renderTable();
            return;
        }
        renderTable();
        loadManageDetail();
    }

    function norm(s) {
        return String(s || '').trim().toLowerCase();
    }

    function compareStrings(a, b) {
        return String(a || '').localeCompare(String(b || ''), 'de', { sensitivity: 'base' });
    }

    function getVisibleRows() {
        const filterInp = document.getElementById('guFilterText');
        const ownerInp = document.getElementById('guFilterOwner');
        const q = filterInp && filterInp.value ? norm(filterInp.value) : '';
        const oq = ownerInp && ownerInp.value ? norm(ownerInp.value) : '';

        let rows = loadedRows;

        if (q) {
            rows = rows.filter(function (r) {
                const a = norm(r.displayName);
                const b = norm(r.mail);
                return a.indexOf(q) !== -1 || b.indexOf(q) !== -1;
            });
        }

        if (oq) {
            rows = rows.filter(function (r) {
                const owners = Array.isArray(r.owners) ? r.owners.join(' ') : '';
                return norm(owners).indexOf(oq) !== -1;
            });
        }

        rows = rows.slice();
        const key = sortState && sortState.key ? sortState.key : 'displayName';
        const dir = sortState && sortState.dir === 'desc' ? -1 : 1;

        rows.sort(function (ra, rb) {
            if (key === 'memberCount') {
                return (Number(ra.memberCount || 0) - Number(rb.memberCount || 0)) * dir;
            }
            if (key === 'createdTs') {
                return (Number(ra.createdTs || 0) - Number(rb.createdTs || 0)) * dir;
            }
            if (key === 'ownerKey') {
                return compareStrings(ra.ownerKey || '', rb.ownerKey || '') * dir;
            }
            if (key === 'id') {
                return compareStrings(ra.id || '', rb.id || '') * dir;
            }
            if (key === 'mail') {
                return compareStrings(ra.mail || '', rb.mail || '') * dir;
            }
            return compareStrings(ra.displayName || '', rb.displayName || '') * dir;
        });

        return rows;
    }

    function updateSortIndicators() {
        const table = document.getElementById('guTable');
        if (!table) return;
        const ths = table.querySelectorAll('th[data-gu-sort]');
        for (let i = 0; i < ths.length; i++) {
            const th = ths[i];
            const key = th.getAttribute('data-gu-sort') || '';
            const ind = th.querySelector('.kt-sort-indicator');
            if (!ind) continue;
            if (key && sortState && sortState.key === key) {
                ind.textContent = sortState.dir === 'desc' ? '▼' : '▲';
            } else {
                ind.textContent = '';
            }
        }
    }

    function renderTable() {
        const tbody = document.getElementById('guTbody');
        if (!tbody) return;
        tbody.replaceChildren();

        updateSortIndicators();
        const rows = getVisibleRows();

        if (!rows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 7;
            td.style.color = '#6c757d';
            td.textContent = loadedRows.length ? 'Keine Treffer für den Filter.' : 'Noch keine Daten.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            return;
        }

        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const tr = document.createElement('tr');
            if (selectedGroupId && r.id === selectedGroupId) {
                tr.style.background = 'rgba(94, 114, 228, 0.08)';
            }

            const tdName = document.createElement('td');
            tdName.textContent = r.displayName || '–';

            const tdMail = document.createElement('td');
            tdMail.textContent = r.mail || '–';

            const tdId = document.createElement('td');
            tdId.textContent = r.id || '–';
            tdId.style.fontFamily = 'Consolas, monospace';
            tdId.style.fontSize = '0.85em';
            tdId.style.wordBreak = 'break-all';

            const tdCreated = document.createElement('td');
            tdCreated.textContent = r.created;

            const tdOwn = document.createElement('td');
            tdOwn.style.whiteSpace = 'pre-wrap';
            tdOwn.style.wordBreak = 'break-word';
            tdOwn.style.fontSize = '0.88em';
            tdOwn.textContent = r.owners.length ? r.owners.join(', ') : '–';

            const tdMem = document.createElement('td');
            tdMem.style.fontSize = '0.88em';
            const countLine = document.createElement('div');
            countLine.style.fontWeight = '600';
            let countText = String(r.memberCount);
            if (r.memberCountIsEstimate) countText += '+';
            countText += r.memberTruncated ? ' (Anzeige gekürzt)' : '';
            countLine.textContent = countText;
            tdMem.appendChild(countLine);
            if (r.members.length) {
                const names = document.createElement('div');
                names.style.marginTop = '6px';
                names.style.whiteSpace = 'pre-wrap';
                names.style.wordBreak = 'break-word';
                names.style.color = '#495057';
                names.textContent = r.members.join(', ');
                tdMem.appendChild(names);
            }

            const tdAct = document.createElement('td');
            const btnSel = document.createElement('button');
            btnSel.type = 'button';
            btnSel.className = 'btn';
            btnSel.style.padding = '6px 10px';
            btnSel.style.fontSize = '0.85em';
            btnSel.textContent = 'Auswählen';
            btnSel.dataset.guSelectGroup = r.id || '';
            tdAct.appendChild(btnSel);

            tr.appendChild(tdName);
            tr.appendChild(tdMail);
            tr.appendChild(tdId);
            tr.appendChild(tdCreated);
            tr.appendChild(tdOwn);
            tr.appendChild(tdMem);
            tr.appendChild(tdAct);
            tbody.appendChild(tr);
        }
    }

    async function loadGroups() {
        const btn = document.getElementById('guBtnLoad');
        const btnCsv = document.getElementById('guBtnCsv');
        const onlyTeams = document.getElementById('guOnlyTeams');
        const progress = document.getElementById('guProgress');
        if (btn) btn.disabled = true;
        if (btnCsv) btnCsv.disabled = true;
        clearLog();
        loadedRows = [];
        selectedGroupId = null;
        const hint = document.getElementById('guManageHint');
        const panel = document.getElementById('guManagePanel');
        if (hint) hint.style.display = '';
        if (panel) panel.style.display = 'none';

        try {
            const token = await getGraphToken();
            appendLog('Lade Microsoft 365-Gruppen …');

            const filter = encodeURIComponent("groupTypes/any(c:c eq 'Unified')");
            const initial =
                '/groups?$filter=' +
                filter +
                '&$select=' +
                encodeURIComponent(GROUP_LIST_SELECT) +
                '&$top=999';

            const groups = await fetchAllPages(token, initial);
            groups.sort(function (a, b) {
                const an = a && a.displayName ? String(a.displayName) : '';
                const bn = b && b.displayName ? String(b.displayName) : '';
                return an.localeCompare(bn, 'de');
            });
            appendLog('Gefunden: ' + groups.length + ' einheitliche Gruppe(n).', 'ok');

            let list = groups;
            if (onlyTeams && onlyTeams.checked) {
                list = groups.filter(groupHasTeamProvisioning);
                appendLog('Nach Teams-Filter: ' + list.length + ' Gruppe(n).', 'ok');
            }

            if (progress) {
                progress.textContent = 'Lade Besitzer und Mitglieder … 0 / ' + list.length;
            }

            const detailed = await mapWithConcurrency(list, 4, async function (g, idx) {
                if (progress && idx % 5 === 0) {
                    progress.textContent =
                        'Lade Besitzer und Mitglieder … ' + (idx + 1) + ' / ' + list.length;
                }
                const id = g.id;
                const ownersRaw = await fetchOwnersForGroup(token, id);
                const memberCountExact = await fetchMemberCount(token, id);
                const membersPage = await fetchMembersFirstPage(token, id);
                const meta = {
                    id: id,
                    displayName: g.displayName || '',
                    mail: g.mail || '',
                    created: formatDate(g.createdDateTime),
                    createdTs: g.createdDateTime ? new Date(g.createdDateTime).getTime() : 0
                };
                return buildRowObjectFromMeta(meta, ownersRaw, memberCountExact, membersPage);
            });

            if (progress) {
                progress.textContent =
                    'Fertig: ' + detailed.length + ' Gruppe(n) mit Besitzern und Mitgliedern.';
            }

            loadedRows = detailed;
            renderTable();
            appendLog('Tabelle aktualisiert.', 'ok');
            toast(detailed.length + ' Gruppe(n) geladen.');
        } catch (e) {
            const msg = e && e.message ? e.message : String(e);
            appendLog('Fehler: ' + msg, 'err');
            toast('Fehler: ' + msg);
            if (progress) progress.textContent = '';
        } finally {
            if (btn) btn.disabled = false;
            if (btnCsv) btnCsv.disabled = !loadedRows.length;
        }
    }

    function escapeCsvCell(v) {
        const s = String(v === undefined || v === null ? '' : v);
        if (/[",\r\n;]/.test(s)) {
            return '"' + s.replace(/"/g, '""') + '"';
        }
        return s;
    }

    function downloadCsv(filename, text) {
        const blob = new Blob([text], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(function () {
            URL.revokeObjectURL(url);
        }, 1500);
    }

    function exportCsv() {
        const rows = getVisibleRows();
        if (!rows.length) {
            toast('Keine Daten zum Export (Filter liefert keine Treffer).');
            return;
        }
        const header = [
            'Anzeigename',
            'E-Mail',
            'Gruppen-ID',
            'Erstellt',
            'Besitzer',
            'Mitgliederzahl',
            'Mitglieder (Anzeige)'
        ];
        const lines = [header.map(escapeCsvCell).join(';')];
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const owners = Array.isArray(r.owners) ? r.owners.join(', ') : '';
            const members = Array.isArray(r.members) ? r.members.join(', ') : '';
            lines.push(
                [
                    r.displayName || '',
                    r.mail || '',
                    r.id || '',
                    r.created || '',
                    owners,
                    String(r.memberCount || 0) + (r.memberCountIsEstimate ? '+' : ''),
                    members
                ]
                    .map(escapeCsvCell)
                    .join(';')
            );
        }
        const d = new Date();
        const stamp =
            d.getFullYear() +
            '-' +
            String(d.getMonth() + 1).padStart(2, '0') +
            '-' +
            String(d.getDate()).padStart(2, '0');
        downloadCsv('ms365-gruppen-' + stamp + '.csv', lines.join('\r\n'));
        toast(rows.length + ' Zeile(n) als CSV exportiert.');
    }

    async function onLogin() {
        const btn = document.getElementById('guBtnLogin');
        if (btn) btn.disabled = true;
        try {
            await getGraphToken();
            toast('Angemeldet – Sie können „Gruppen laden“ wählen.');
            appendLog('Anmeldung erfolgreich.', 'ok');
        } catch (e) {
            toast('Anmeldung: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    function toggleSort(key) {
        if (!key) return;
        if (sortState.key === key) {
            sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc';
        } else {
            sortState.key = key;
            sortState.dir = 'asc';
        }
        renderTable();
        toast('Sortierung: ' + (SORT_LABELS[key] || key) + ' (' + (sortState.dir === 'desc' ? 'absteigend' : 'aufsteigend') + ')');
    }

    function fillUserSearchSelect(users) {
        const sel = document.getElementById('guUserSearchResults');
        if (!sel) return;
        sel.replaceChildren();
        if (!users || !users.length) {
            const opt = document.createElement('option');
            opt.value = '';
            opt.textContent = '(keine Treffer)';
            sel.appendChild(opt);
            return;
        }
        for (let i = 0; i < users.length; i++) {
            const u = users[i];
            const opt = document.createElement('option');
            opt.value = u.id || '';
            const dn = u.displayName || '';
            const up = u.userPrincipalName || u.mail || '';
            opt.textContent = (dn || up || u.id) + (up && dn ? ' (' + up + ')' : '');
            sel.appendChild(opt);
        }
    }

    async function runUserSearch() {
        const inp = document.getElementById('guUserSearch');
        const q = inp && inp.value ? inp.value.trim() : '';
        if (!q) {
            toast('Bitte einen Suchbegriff eingeben.');
            return;
        }
        setManageBusy(true);
        try {
            const token = await getGraphToken();
            const users = await graphSearchUsers(token, q);
            fillUserSearchSelect(users);
            appendLog('Benutzersuche: ' + users.length + ' Treffer.', 'ok');
        } catch (e) {
            appendLog('Benutzersuche: ' + (e && e.message ? e.message : e), 'err');
            toast(String(e && e.message ? e.message : e));
        } finally {
            setManageBusy(false);
        }
    }

    function getSelectedUserIdFromSearch() {
        const sel = document.getElementById('guUserSearchResults');
        if (!sel || !sel.value) return '';
        return String(sel.value).trim();
    }

    async function runAddOwner() {
        if (!selectedGroupId) return;
        const userId = getSelectedUserIdFromSearch();
        if (!userId) {
            toast('Bitte einen Benutzer aus der Trefferliste auswählen.');
            return;
        }
        setManageBusy(true);
        try {
            const token = await getGraphToken();
            await graphAddOwnerWithMemberFallback(token, selectedGroupId, userId);
            appendLog('Besitzer hinzugefügt.', 'ok');
            await refreshLoadedRow(token, selectedGroupId);
            await loadManageDetail();
            toast('Besitzer hinzugefügt.');
        } catch (e) {
            appendLog('Besitzer: ' + (e && e.message ? e.message : e), 'err');
            toast(String(e && e.message ? e.message : e));
        } finally {
            setManageBusy(false);
        }
    }

    async function runAddMember() {
        if (!selectedGroupId) return;
        const userId = getSelectedUserIdFromSearch();
        if (!userId) {
            toast('Bitte einen Benutzer aus der Trefferliste auswählen.');
            return;
        }
        setManageBusy(true);
        try {
            const token = await getGraphToken();
            await graphAddGroupMember(token, selectedGroupId, userId);
            appendLog('Mitglied hinzugefügt.', 'ok');
            await refreshLoadedRow(token, selectedGroupId);
            await loadManageDetail();
            toast('Mitglied hinzugefügt.');
        } catch (e) {
            const msg = e && e.message ? e.message : String(e);
            if (msg.indexOf('added object references already exist') !== -1) {
                appendLog('Mitglied war bereits in der Gruppe.', 'warn');
                toast('War bereits Mitglied.');
            } else {
                appendLog('Mitglied: ' + msg, 'err');
                toast(msg);
            }
        } finally {
            setManageBusy(false);
        }
    }

    async function runRemoveOwner(userId) {
        if (!selectedGroupId || !userId) return;
        setManageBusy(true);
        try {
            const token = await getGraphToken();
            await graphRemoveGroupOwner(token, selectedGroupId, userId);
            appendLog('Besitzer entfernt.', 'ok');
            await refreshLoadedRow(token, selectedGroupId);
            await loadManageDetail();
            toast('Besitzer entfernt.');
        } catch (e) {
            appendLog('Besitzer entfernen: ' + (e && e.message ? e.message : e), 'err');
            toast(String(e && e.message ? e.message : e));
        } finally {
            setManageBusy(false);
        }
    }

    async function runRemoveMember(memberId) {
        if (!selectedGroupId || !memberId) return;
        setManageBusy(true);
        try {
            const token = await getGraphToken();
            await graphRemoveGroupMember(token, selectedGroupId, memberId);
            appendLog('Mitglied entfernt.', 'ok');
            await refreshLoadedRow(token, selectedGroupId);
            await loadManageDetail();
            toast('Mitglied entfernt.');
        } catch (e) {
            appendLog('Mitglied entfernen: ' + (e && e.message ? e.message : e), 'err');
            toast(String(e && e.message ? e.message : e));
        } finally {
            setManageBusy(false);
        }
    }

    function bind() {
        const btnL = document.getElementById('guBtnLogin');
        const btnLoad = document.getElementById('guBtnLoad');
        const btnCsv = document.getElementById('guBtnCsv');
        const filt = document.getElementById('guFilterText');
        const filtOwner = document.getElementById('guFilterOwner');
        const table = document.getElementById('guTable');
        const btnSearchUser = document.getElementById('guBtnUserSearch');
        const btnAddOwner = document.getElementById('guBtnAddOwner');
        const btnAddMember = document.getElementById('guBtnAddMember');
        const btnRefreshManage = document.getElementById('guBtnRefreshManage');
        const managePanel = document.getElementById('guManagePanel');
        if (btnL) btnL.addEventListener('click', () => onLogin());
        if (btnLoad) btnLoad.addEventListener('click', () => loadGroups());
        if (btnCsv) btnCsv.addEventListener('click', () => exportCsv());
        if (filt) filt.addEventListener('input', () => renderTable());
        if (filtOwner) filtOwner.addEventListener('input', () => renderTable());
        if (btnCsv) btnCsv.disabled = !loadedRows.length;

        if (btnSearchUser) btnSearchUser.addEventListener('click', () => runUserSearch());
        if (btnAddOwner) btnAddOwner.addEventListener('click', () => runAddOwner());
        if (btnAddMember) btnAddMember.addEventListener('click', () => runAddMember());
        if (btnRefreshManage) btnRefreshManage.addEventListener('click', () => loadManageDetail());

        if (table) {
            table.addEventListener('click', function (ev) {
                const t = ev.target;
                if (!t) return;
                const th = t.closest ? t.closest('th[data-gu-sort]') : null;
                if (!th) return;
                const key = th.getAttribute('data-gu-sort');
                toggleSort(key);
            });
            table.addEventListener('click', function (ev) {
                const t = ev.target;
                if (!t || !t.closest) return;
                const btn = t.closest('button[data-gu-select-group]');
                if (!btn) return;
                const gid = btn.getAttribute('data-gu-select-group');
                if (gid) selectGroup(gid);
            });
        }

        if (managePanel) {
            managePanel.addEventListener('click', function (ev) {
                const t = ev.target;
                if (!t || !t.closest) return;
                const bo = t.closest('button[data-gu-remove-owner]');
                if (bo) {
                    const uid = bo.getAttribute('data-gu-remove-owner');
                    if (uid) runRemoveOwner(uid);
                    return;
                }
                const bm = t.closest('button[data-gu-remove-member]');
                if (bm) {
                    const mid = bm.getAttribute('data-gu-remove-member');
                    if (mid) runRemoveMember(mid);
                }
            });
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bind);
    } else {
        bind();
    }
})();
