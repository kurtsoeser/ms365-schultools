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
    /** @type {{ id: string, displayName: string, mail: string, created: string, owners: string[], members: string[], memberCount: number, memberCountIsEstimate: boolean, memberTruncated: boolean }[]} */
    let loadedRows = [];

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

    function renderTable() {
        const tbody = document.getElementById('guTbody');
        const filterInp = document.getElementById('guFilterText');
        const q = filterInp && filterInp.value ? String(filterInp.value).trim().toLowerCase() : '';

        if (!tbody) return;
        tbody.replaceChildren();

        let rows = loadedRows;
        if (q) {
            rows = loadedRows.filter(function (r) {
                const a = (r.displayName || '').toLowerCase();
                const b = (r.mail || '').toLowerCase();
                return a.indexOf(q) !== -1 || b.indexOf(q) !== -1;
            });
        }

        if (!rows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 5;
            td.style.color = '#6c757d';
            td.textContent = loadedRows.length ? 'Keine Treffer für den Filter.' : 'Noch keine Daten.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            return;
        }

        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const tr = document.createElement('tr');

            const tdName = document.createElement('td');
            tdName.textContent = r.displayName || '–';

            const tdMail = document.createElement('td');
            tdMail.textContent = r.mail || '–';

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

            tr.appendChild(tdName);
            tr.appendChild(tdMail);
            tr.appendChild(tdCreated);
            tr.appendChild(tdOwn);
            tr.appendChild(tdMem);
            tbody.appendChild(tr);
        }
    }

    async function loadGroups() {
        const btn = document.getElementById('guBtnLoad');
        const onlyTeams = document.getElementById('guOnlyTeams');
        const progress = document.getElementById('guProgress');
        if (btn) btn.disabled = true;
        clearLog();
        loadedRows = [];

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
                const membersRaw = membersPage.items;

                const ownerLabels = [];
                for (let oi = 0; oi < ownersRaw.length; oi++) {
                    const lab = formatPerson(ownersRaw[oi]);
                    if (lab) ownerLabels.push(lab);
                }
                ownerLabels.sort(function (a, b) {
                    return a.localeCompare(b, 'de');
                });

                const memberLabels = [];
                for (let mi = 0; mi < membersRaw.length; mi++) {
                    const lab = formatPerson(membersRaw[mi]);
                    if (lab) memberLabels.push(lab);
                }
                memberLabels.sort(function (a, b) {
                    return a.localeCompare(b, 'de');
                });

                let totalMembers =
                    memberCountExact >= 0 ? memberCountExact : membersRaw.length;
                const countIsEstimate = memberCountExact < 0 && membersPage.hasMore;

                let truncated = membersPage.hasMore || memberLabels.length > MAX_MEMBER_NAMES_SHOWN;
                let shown = memberLabels;
                if (memberLabels.length > MAX_MEMBER_NAMES_SHOWN) {
                    shown = memberLabels.slice(0, MAX_MEMBER_NAMES_SHOWN);
                }

                return {
                    id: id,
                    displayName: g.displayName || '',
                    mail: g.mail || '',
                    created: formatDate(g.createdDateTime),
                    owners: ownerLabels,
                    members: shown,
                    memberCount: totalMembers,
                    memberCountIsEstimate: countIsEstimate,
                    memberTruncated: truncated
                };
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
        }
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

    function bind() {
        const btnL = document.getElementById('guBtnLogin');
        const btnLoad = document.getElementById('guBtnLoad');
        const filt = document.getElementById('guFilterText');
        if (btnL) btnL.addEventListener('click', () => onLogin());
        if (btnLoad) btnLoad.addEventListener('click', () => loadGroups());
        if (filt) filt.addEventListener('input', () => renderTable());
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bind);
    } else {
        bind();
    }
})();
