(function () {
    'use strict';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/TeamSettings.ReadWrite.All',
        'https://graph.microsoft.com/Group.ReadWrite.All'
    ];

    let msalMod = null;
    let pca = null;

    function toast(msg) {
        const el = document.getElementById('toast');
        if (el) {
            el.textContent = msg;
            el.classList.add('show');
            clearTimeout(toast._t);
            toast._t = setTimeout(() => el.classList.remove('show'), 3800);
        } else if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(msg);
        } else if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(msg);
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

    async function graphRequest(method, path, token, body, extraHeaders) {
        const url = path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path;
        let attempt = 0;
        while (true) {
            const headers = Object.assign({ Authorization: 'Bearer ' + token }, extraHeaders || {});
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
                typeof data === 'object' && data && data.error
                    ? JSON.stringify(data.error)
                    : text || String(res.status);
            throw new Error(method + ' ' + path + ': ' + msg);
        }
        return data || {};
    }

    function appendLog(msg, kind) {
        const el = document.getElementById('taLog');
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
        const el = document.getElementById('taLog');
        if (el) el.replaceChildren();
    }

    function parseTeamsOperationPathFromLocation(locationHeader) {
        if (!locationHeader) return null;
        let loc = String(locationHeader).trim();
        if (loc.indexOf('http') === 0) {
            try {
                const u = new URL(loc);
                loc = u.pathname.replace(/^\/v1\.0/i, '');
            } catch {
                return null;
            }
        }
        const m = loc.match(/\/teams\/([^/]+)\/operations\/([^/?\s]+)/i);
        if (m) return '/teams/' + m[1] + '/operations/' + m[2];
        const m2 = loc.match(/teams\('([^']+)'\)\/operations\('([^']+)'\)/i);
        if (m2) return '/teams/' + m2[1] + '/operations/' + m2[2];
        return null;
    }

    async function pollTeamsAsyncOperation(token, operationPath) {
        const maxAttempts = 90;
        for (let i = 0; i < maxAttempts; i++) {
            await sleep(2000);
            const data = await graphJson('GET', operationPath, token, undefined);
            const st = String(data.status || data.Status || '').toLowerCase();
            if (st === 'succeeded') {
                appendLog('Asynchrone Teams-Operation abgeschlossen.', 'ok');
                return;
            }
            if (st === 'failed') {
                const errMsg =
                    (data.error && (data.error.message || JSON.stringify(data.error))) || JSON.stringify(data);
                throw new Error('Teams-Operation fehlgeschlagen: ' + errMsg);
            }
            if (i > 0 && i % 10 === 0) {
                appendLog('Warte auf Teams-Operation … (' + i * 2 + ' s)', 'warn');
            }
        }
        throw new Error('Timeout: Teams-Operation nicht abgeschlossen.');
    }

    function normGuid(v) {
        const s = String(v || '').trim();
        if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(s)) {
            return s;
        }
        return '';
    }

    function odataEscape(s) {
        return String(s).replace(/'/g, "''");
    }

    function groupHasTeamProvisioning(g) {
        const opts = g && g.resourceProvisioningOptions;
        return Array.isArray(opts) && opts.indexOf('Team') !== -1;
    }

    /**
     * Prüft Teams-Ressource: zuerst resourceProvisioningOptions, sonst einmal GET …/team.
     */
    async function filterGroupsWithTeam(token, groups) {
        const out = [];
        for (const g of groups) {
            if (!g || !g.id) continue;
            if (groupHasTeamProvisioning(g)) {
                out.push(g);
                continue;
            }
            try {
                const team = await graphJson('GET', '/groups/' + encodeURIComponent(g.id) + '/team', token, undefined);
                if (team && (team.id || team.displayName !== undefined)) {
                    out.push(g);
                }
            } catch {
                // keine Team-Ressource
            }
        }
        return out;
    }

    const GROUP_SELECT =
        'id,displayName,mail,mailNickname,resourceProvisioningOptions';

    async function searchTeam(token, query) {
        const q = String(query || '').trim();
        if (!q) throw new Error('Suchbegriff fehlt.');

        if (normGuid(q)) {
            try {
                const g = await graphJson(
                    'GET',
                    '/groups/' + encodeURIComponent(q) + '?$select=' + GROUP_SELECT,
                    token,
                    undefined
                );
                const withTeam = await filterGroupsWithTeam(token, [g]);
                if (!withTeam.length) {
                    throw new Error('Diese GUID ist keine Microsoft-365-Gruppe mit Team.');
                }
                const x = withTeam[0];
                return { id: x.id, displayName: x.displayName || '', mail: x.mail || '' };
            } catch (e) {
                throw new Error('GUID nicht gefunden oder kein Team: ' + (e.message || e));
            }
        }

        let collection = [];
        if (q.indexOf('@') !== -1) {
            const filter = "mail eq '" + odataEscape(q) + "'";
            const data = await graphJson(
                'GET',
                '/groups?$filter=' + encodeURIComponent(filter) + '&$select=' + GROUP_SELECT,
                token,
                undefined
            );
            collection = data.value || [];
        } else {
            const filter = "startswith(displayName,'" + odataEscape(q) + "')";
            const data = await graphJson(
                'GET',
                '/groups?$filter=' +
                    encodeURIComponent(filter) +
                    '&$select=' +
                    GROUP_SELECT +
                    '&$top=15',
                token,
                undefined
            );
            collection = data.value || [];
            if (!collection.length) {
                const ex = "displayName eq '" + odataEscape(q) + "'";
                const data2 = await graphJson(
                    'GET',
                    '/groups?$filter=' +
                        encodeURIComponent(ex) +
                        '&$select=' +
                        GROUP_SELECT +
                        '&$top=15',
                    token,
                    undefined
                );
                collection = data2.value || [];
            }
        }

        const withTeams = await filterGroupsWithTeam(token, collection);
        if (!withTeams.length) {
            throw new Error('Kein Team zu diesen Suchkriterien gefunden (oder keine Teams-Ressource).');
        }
        if (withTeams.length > 1) {
            appendLog(
                'Hinweis: Mehrere Treffer – es wird das erste Team mit Teams-Ressource verwendet. Bitte ggf. die GUID direkt eintragen.',
                'warn'
            );
        }
        const g = withTeams[0];
        return { id: g.id, displayName: g.displayName || '', mail: g.mail || '' };
    }

    async function runArchiveOrUnarchive(archive) {
        const idInp = document.getElementById('taTeamId');
        const spo = document.getElementById('taSpoReadOnly');
        const teamId = normGuid(idInp && idInp.value);
        if (!teamId) {
            toast('Bitte zuerst eine gültige Team-/Gruppen-ID eintragen oder „Team suchen“ verwenden.');
            return;
        }

        const btnA = document.getElementById('taBtnArchive');
        const btnU = document.getElementById('taBtnUnarchive');
        if (btnA) btnA.disabled = true;
        if (btnU) btnU.disabled = true;

        try {
            const token = await getGraphToken();
            const path = '/teams/' + encodeURIComponent(teamId) + (archive ? '/archive' : '/unarchive');
            let body = undefined;
            if (archive && spo && spo.checked) {
                body = { shouldSetSpoSiteReadOnlyForMembers: true };
            }

            appendLog((archive ? 'Archivierung' : 'Aufheben der Archivierung') + ' starten …');
            const res = await graphRequest('POST', path, token, body);

            if (res.status !== 202 && res.status !== 200) {
                const t = await res.text();
                throw new Error('HTTP ' + res.status + ' ' + t);
            }

            const loc = res.headers.get('Location') || res.headers.get('Content-Location');
            const opPath = parseTeamsOperationPathFromLocation(loc);
            if (opPath) {
                appendLog('Asynchrone Verarbeitung (202) – Status wird abgefragt …', 'warn');
                await pollTeamsAsyncOperation(token, opPath);
            } else {
                appendLog('Keine Operation-URL in der Antwort – bitte Status im Admin Center prüfen.', 'warn');
            }

            appendLog(archive ? 'Team archiviert (Anfrage erfolgreich).' : 'Archivierung aufgehoben (Anfrage erfolgreich).', 'ok');
            toast(archive ? 'Archivierung ausgeführt.' : 'Archivierung aufgehoben.');
        } catch (e) {
            const msg = e && e.message ? e.message : String(e);
            appendLog('Fehler: ' + msg, 'err');
            toast('Fehler: ' + msg);
        } finally {
            if (btnA) btnA.disabled = false;
            if (btnU) btnU.disabled = false;
        }
    }

    // ---------------------------------------------------------------
    // Sammel-Archivierung: alle Kursteams eines Schuljahres (Suchbegriff)
    // ---------------------------------------------------------------

    /** @type {Array<{ id: string, displayName: string, mailNickname: string, selected: boolean }>} */
    let bulkResults = [];

    function appendBulkLog(msg, kind) {
        const el = document.getElementById('taBulkLog');
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

    function clearBulkLog() {
        const el = document.getElementById('taBulkLog');
        if (el) el.replaceChildren();
    }

    async function listTeamGroupsFromGraphFallback(token) {
        const collected = [];
        let nextPath = '/groups?$select=id,displayName,mailNickname,resourceProvisioningOptions&$top=999';
        while (nextPath) {
            const data = await graphJson('GET', nextPath, token, undefined);
            (data.value || []).forEach(function (g) {
                if (g && g.id && g.mailNickname && groupHasTeamProvisioning(g)) collected.push(g);
            });
            nextPath = data['@odata.nextLink'] || null;
        }
        return collected;
    }

    /** Alle M365-Gruppen mit Teams-Ressource (mailNickname gesetzt) – über alle Seiten. */
    async function listTeamGroupsFromGraph(token) {
        const collected = [];
        let nextPath =
            "/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=" +
            GROUP_SELECT +
            '&$top=999';
        try {
            while (nextPath) {
                const useEv = nextPath.indexOf('http') !== 0;
                const data = await graphJson(
                    'GET',
                    nextPath,
                    token,
                    undefined,
                    useEv ? { ConsistencyLevel: 'eventual' } : undefined
                );
                (data.value || []).forEach(function (g) {
                    if (g && g.id && g.mailNickname) collected.push(g);
                });
                nextPath = data['@odata.nextLink'] || null;
            }
            return collected;
        } catch {
            return listTeamGroupsFromGraphFallback(token);
        }
    }

    function renderBulkResults() {
        const wrap = document.getElementById('taBulkResultsWrap');
        const body = document.getElementById('taBulkResultsBody');
        const summary = document.getElementById('taBulkResultsSummary');
        if (!wrap || !body) return;

        if (!bulkResults.length) {
            wrap.style.display = 'none';
            body.replaceChildren();
            return;
        }
        wrap.style.display = '';
        const nSel = bulkResults.filter(function (r) {
            return r.selected;
        }).length;
        if (summary) {
            summary.textContent = bulkResults.length + ' Kursteam(s) gefunden, ' + nSel + ' ausgewählt.';
        }

        body.replaceChildren();
        bulkResults.forEach(function (r, idx) {
            const tr = document.createElement('tr');

            const tdSel = document.createElement('td');
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.checked = !!r.selected;
            cb.addEventListener('change', function () {
                bulkResults[idx].selected = cb.checked;
                renderBulkResults();
            });
            tdSel.appendChild(cb);

            const tdName = document.createElement('td');
            tdName.textContent = r.displayName || '(ohne Name)';

            const tdMail = document.createElement('td');
            tdMail.textContent = r.mailNickname || '';

            tr.append(tdSel, tdName, tdMail);
            body.appendChild(tr);
        });
    }

    function setAllBulkSelected(value) {
        bulkResults = bulkResults.map(function (r) {
            return Object.assign({}, r, { selected: value });
        });
        renderBulkResults();
    }

    async function onBulkSearch() {
        const inp = document.getElementById('taBulkPrefix');
        const term = inp && inp.value ? inp.value.trim() : '';
        if (!term) {
            toast('Bitte zuerst einen Suchbegriff eintragen (z. B. das Schuljahr-Präfix wie SJ25).');
            return;
        }
        const btn = document.getElementById('taBulkSearch');
        if (btn) btn.disabled = true;
        clearBulkLog();
        bulkResults = [];
        renderBulkResults();
        try {
            const token = await getGraphToken();
            appendBulkLog('Lade Team-Gruppen aus Microsoft 365 (Graph) …');
            const groups = await listTeamGroupsFromGraph(token);
            const F = window.ms365TeamsArchivBulkLogic;
            const matched = F && typeof F.filterAndSortGroupsByTerm === 'function'
                ? F.filterAndSortGroupsByTerm(groups, term)
                : [];
            bulkResults = matched.map(function (g) {
                return { id: g.id, displayName: g.displayName, mailNickname: g.mailNickname, selected: true };
            });
            renderBulkResults();
            appendBulkLog(
                'Gefunden: ' + bulkResults.length + ' von ' + groups.length + ' Team-Gruppe(n) im Tenant passen zu „' + term + '“.',
                bulkResults.length ? 'ok' : 'warn'
            );
            toast(bulkResults.length + ' Kursteam(s) gefunden.');
        } catch (e) {
            appendBulkLog('Fehler: ' + (e && e.message ? e.message : e), 'err');
            toast(String((e && e.message) || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    function isAlreadyInTargetStateError(msg) {
        return /already\s*archived/i.test(msg) || /not\s*.*archived/i.test(msg) || /isn't\s*archived/i.test(msg);
    }

    async function runBulkArchiveOrUnarchive(archive) {
        const selected = bulkResults.filter(function (r) {
            return r.selected;
        });
        if (!selected.length) {
            toast('Bitte zuerst Kursteams suchen und mindestens eines auswählen.');
            return;
        }

        const verb = archive ? 'archivieren' : 'ihre Archivierung aufheben';
        const question =
            selected.length +
            ' Kursteam(s) wirklich ' +
            verb +
            '?\n\nDie zugrunde liegenden Microsoft-365-Gruppen, Dateien und Chatverläufe bleiben erhalten – nur der Team-Status ändert sich.';
        const confirmed =
            typeof window.ms365AppDialogConfirm === 'function'
                ? await window.ms365AppDialogConfirm(question, {
                      title: archive ? 'Kursteams archivieren' : 'Archivierung aufheben',
                      okText: archive ? 'Archivieren' : 'Aufheben',
                      danger: !!archive
                  })
                : window.confirm(question);
        if (!confirmed) return;

        const btnA = document.getElementById('taBulkArchive');
        const btnU = document.getElementById('taBulkUnarchive');
        const btnS = document.getElementById('taBulkSearch');
        if (btnA) btnA.disabled = true;
        if (btnU) btnU.disabled = true;
        if (btnS) btnS.disabled = true;

        clearBulkLog();
        appendBulkLog(
            (archive ? 'Archivierung' : 'Aufheben der Archivierung') + ' für ' + selected.length + ' Team(s) starten …'
        );

        let token;
        try {
            token = await getGraphToken();
        } catch (e) {
            appendBulkLog('Anmeldung/Token: ' + (e.message || e), 'err');
            if (btnA) btnA.disabled = false;
            if (btnU) btnU.disabled = false;
            if (btnS) btnS.disabled = false;
            return;
        }

        let okCount = 0;
        let failCount = 0;
        for (let i = 0; i < selected.length; i++) {
            const item = selected[i];
            const label = item.displayName || item.mailNickname || item.id;
            appendBulkLog('[' + (i + 1) + '/' + selected.length + '] ' + label + ' …');
            try {
                const path = '/teams/' + encodeURIComponent(item.id) + (archive ? '/archive' : '/unarchive');
                const res = await graphRequest('POST', path, token, undefined);
                if (res.status !== 202 && res.status !== 200) {
                    const t = await res.text();
                    throw new Error('HTTP ' + res.status + ' ' + t);
                }
                appendBulkLog('  OK (Anfrage angenommen): ' + label, 'ok');
                okCount++;
            } catch (e) {
                const msg = e && e.message ? e.message : String(e);
                if (isAlreadyInTargetStateError(msg)) {
                    appendBulkLog('  Übersprungen (bereits im Zielstatus): ' + label, 'warn');
                    okCount++;
                } else {
                    appendBulkLog('  Fehler bei ' + label + ': ' + msg, 'err');
                    failCount++;
                }
            }
            await sleep(1200);
            try {
                token = await getGraphToken();
            } catch (e) {
                appendBulkLog('Token erneuern: ' + (e.message || e), 'err');
                break;
            }
        }

        appendBulkLog(
            'Fertig: ' + okCount + ' OK, ' + failCount + ' Fehler. Bereitstellung läuft im Hintergrund weiter (siehe Teams-Admin-Center).',
            failCount ? 'warn' : 'ok'
        );
        toast('Sammel-Archivierung: ' + okCount + ' OK' + (failCount ? ', ' + failCount + ' Fehler.' : '.'));

        if (btnA) btnA.disabled = false;
        if (btnU) btnU.disabled = false;
        if (btnS) btnS.disabled = false;
    }

    async function onSearch() {
        const searchInp = document.getElementById('taSearch');
        const idInp = document.getElementById('taTeamId');
        const summary = document.getElementById('taResolvedSummary');
        const q = searchInp && searchInp.value ? searchInp.value.trim() : '';
        if (!q) {
            toast('Bitte E-Mail oder Anzeigename (oder GUID) im Suchfeld eintragen.');
            return;
        }
        clearLog();
        try {
            const token = await getGraphToken();
            appendLog('Suche …');
            const found = await searchTeam(token, q);
            if (idInp) idInp.value = found.id;
            if (summary) {
                summary.style.display = '';
                summary.textContent =
                    'Gefunden: ' +
                    (found.displayName || '(ohne Name)') +
                    (found.mail ? ' · ' + found.mail : '') +
                    ' · ID: ' +
                    found.id;
            }
            appendLog('Team gefunden – ID wurde übernommen.', 'ok');
            toast('Team gefunden.');
        } catch (e) {
            appendLog('Suche: ' + (e && e.message ? e.message : e), 'err');
            toast(String(e && e.message ? e.message : e));
        }
    }

    async function onLogin() {
        const btn = document.getElementById('taBtnLogin');
        if (btn) btn.disabled = true;
        try {
            await getGraphToken();
            toast('Angemeldet – Sie können archivieren oder die Archivierung aufheben.');
        } catch (e) {
            toast('Anmeldung: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    function bind() {
        const btnL = document.getElementById('taBtnLogin');
        const btnS = document.getElementById('taBtnSearch');
        const btnA = document.getElementById('taBtnArchive');
        const btnU = document.getElementById('taBtnUnarchive');
        if (btnL) btnL.addEventListener('click', () => onLogin());
        if (btnS) btnS.addEventListener('click', () => onSearch());
        if (btnA) btnA.addEventListener('click', () => runArchiveOrUnarchive(true));
        if (btnU) btnU.addEventListener('click', () => runArchiveOrUnarchive(false));

        const btnBulkSearch = document.getElementById('taBulkSearch');
        const btnBulkAll = document.getElementById('taBulkSelectAll');
        const btnBulkNone = document.getElementById('taBulkSelectNone');
        const btnBulkArchive = document.getElementById('taBulkArchive');
        const btnBulkUnarchive = document.getElementById('taBulkUnarchive');
        if (btnBulkSearch) btnBulkSearch.addEventListener('click', () => onBulkSearch());
        if (btnBulkAll) btnBulkAll.addEventListener('click', () => setAllBulkSelected(true));
        if (btnBulkNone) btnBulkNone.addEventListener('click', () => setAllBulkSelected(false));
        if (btnBulkArchive) btnBulkArchive.addEventListener('click', () => runBulkArchiveOrUnarchive(true));
        if (btnBulkUnarchive) btnBulkUnarchive.addEventListener('click', () => runBulkArchiveOrUnarchive(false));
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bind);
    } else {
        bind();
    }
})();
