(function () {
    'use strict';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/Group.ReadWrite.All'
    ];

    const GROUP_SELECT = 'id,displayName,mail,mailNickname,resourceProvisioningOptions';
    const STEP_ORDER = [1, 2, 3, 4];

    let msalMod = null;
    let pca = null;
    /** @type {{ id: string, displayName: string, mail: string, mailNickname: string, resourceProvisioningOptions?: string[] }[]} */
    let loadedGroups = [];
    /** @type {Array<{ id: string, displayNameOld: string, displayNameNew: string, mailNicknameOld: string, mailNicknameNew: string, mailOld: string, hint: string, ok: boolean }>} */
    let previewRows = [];

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

    async function graphRequest(method, pathOrUrl, token, body) {
        const url =
            pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
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
        const el = document.getElementById('kuLog');
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
        const el = document.getElementById('kuLog');
        if (el) el.replaceChildren();
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

    function escapeRe(s) {
        return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    function odataEscape(s) {
        return String(s).replace(/'/g, "''");
    }

    function deriveMailNickname(displayName) {
        let s = displayName.trim().toLowerCase().replace(/\s+/g, '');
        s = s.replace(/[^a-z0-9\-]/g, '');
        if (s.length > 64) s = s.slice(0, 64);
        return s;
    }

    /**
     * Erwartetes Muster: "<Präfix> <Stufe><Kürzel>" z. B. "Klasse 1A", "Klasse 10HAK"
     */
    function computeNewDisplayName(displayName, prefix, fromGrade, toGrade) {
        const p = String(prefix || '').trim();
        const fg = String(fromGrade || '').trim();
        const tg = String(toGrade || '').trim();
        if (!p || !fg || !tg) return null;
        const re = new RegExp('^' + escapeRe(p) + '\\s+(\\d+)([A-Za-z0-9\\-]*)$', 'i');
        const m = String(displayName || '').trim().match(re);
        if (!m) return null;
        if (m[1] !== fg) return null;
        return p + ' ' + tg + (m[2] || '');
    }

    function getFilteredGroups() {
        const onlyTeams = document.getElementById('kuOnlyTeams');
        const filtInp = document.getElementById('kuFilterContains');
        const sub = filtInp && filtInp.value ? String(filtInp.value).trim().toLowerCase() : '';
        let list = loadedGroups.slice();
        if (onlyTeams && onlyTeams.checked) {
            list = list.filter(groupHasTeamProvisioning);
        }
        if (sub) {
            list = list.filter(function (g) {
                return String(g.displayName || '')
                    .toLowerCase()
                    .indexOf(sub) !== -1;
            });
        }
        return list;
    }

    async function checkMailNicknameConflict(token, mailNickname, excludeId) {
        if (!mailNickname) return 'Leerer Mail-Nickname.';
        const filter = "mailNickname eq '" + odataEscape(mailNickname) + "'";
        const path = '/groups?$filter=' + encodeURIComponent(filter) + '&$select=id,displayName';
        const data = await graphJson('GET', path, token, undefined);
        const v = data.value || [];
        if (!v.length) return null;
        if (v.length === 1 && v[0].id === excludeId) return null;
        return (
            'Mail-Nickname bereits vergeben (Gruppe: ' +
            (v[0].displayName || v[0].id) +
            ').'
        );
    }

    async function buildPreview() {
        const prefixInp = document.getElementById('kuPrefix');
        const fromInp = document.getElementById('kuFromGrade');
        const toInp = document.getElementById('kuToGrade');
        const updNick = document.getElementById('kuUpdateMailNick');
        const prefix = prefixInp && prefixInp.value ? prefixInp.value : 'Klasse';
        const fromG = fromInp && fromInp.value ? fromInp.value.trim() : '';
        const toG = toInp && toInp.value ? toInp.value.trim() : '';
        if (!fromG || !toG) {
            toast('Bitte „Stufe von“ und „Stufe nach“ ausfüllen.');
            return false;
        }
        if (fromG === toG) {
            toast('„Stufe von“ und „Stufe nach“ dürfen nicht gleich sein.');
            return false;
        }

        const token = await getGraphToken();
        const groups = getFilteredGroups();
        previewRows = [];

        for (let i = 0; i < groups.length; i++) {
            const g = groups[i];
            const oldDn = g.displayName || '';
            const newDn = computeNewDisplayName(oldDn, prefix, fromG, toG);
            const mailOld = g.mail || '';
            const nickOld = g.mailNickname || '';

            if (!newDn) {
                previewRows.push({
                    id: g.id,
                    displayNameOld: oldDn,
                    displayNameNew: '—',
                    mailNicknameOld: nickOld,
                    mailNicknameNew: '—',
                    mailOld: mailOld,
                    hint: 'Regel trifft nicht zu (Format oder andere Stufe).',
                    ok: false
                });
                continue;
            }

            let nickNew = nickOld;
            let hint = '';
            if (updNick && updNick.checked) {
                nickNew = deriveMailNickname(newDn);
                if (!nickNew) {
                    hint = 'Abgeleiteter Mail-Nickname leer.';
                }
            }

            const willChangeDn = newDn !== oldDn;
            const willChangeNick = !!(updNick && updNick.checked && nickNew !== nickOld);

            if (!hint && willChangeNick) {
                const conflict = await checkMailNicknameConflict(token, nickNew, g.id);
                if (conflict) hint = conflict;
            }

            if (!hint && !willChangeDn && !willChangeNick) {
                hint = 'Keine Änderung (Name und Mail-Nickname bereits passend).';
            }

            const ok = !hint && (willChangeDn || willChangeNick);

            previewRows.push({
                id: g.id,
                displayNameOld: oldDn,
                displayNameNew: newDn,
                mailNicknameOld: nickOld,
                mailNicknameNew: nickNew,
                mailOld: mailOld,
                hint: hint,
                ok: ok
            });
        }

        const okCount = previewRows.filter(function (r) {
            return r.ok;
        }).length;
        const summary = document.getElementById('kuPreviewSummary');
        if (summary) {
            summary.textContent =
                'Gefilterte Gruppen: ' +
                groups.length +
                '. Davon können ' +
                okCount +
                ' umbenannt werden (ohne Konflikt).';
        }

        renderPreviewTable();
        const next3 = document.getElementById('kuNext3');
        if (next3) next3.disabled = okCount === 0;
        return true;
    }

    function renderPreviewTable() {
        const tbody = document.getElementById('kuPreviewBody');
        if (!tbody) return;
        tbody.replaceChildren();
        for (let i = 0; i < previewRows.length; i++) {
            const r = previewRows[i];
            const tr = document.createElement('tr');
            if (r.hint) tr.style.background = 'rgba(255, 193, 7, 0.12)';
            function td(text) {
                const c = document.createElement('td');
                c.style.fontSize = '0.9em';
                c.style.wordBreak = 'break-word';
                c.textContent = text;
                return c;
            }
            tr.appendChild(td(r.displayNameOld || '–'));
            tr.appendChild(td(r.displayNameNew || '–'));
            tr.appendChild(td(r.mailNicknameOld || '–'));
            tr.appendChild(td(r.mailNicknameNew || '–'));
            tr.appendChild(td(r.hint || (r.ok ? 'OK' : '')));
            tbody.appendChild(tr);
        }
        if (!previewRows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 5;
            td.style.color = '#6c757d';
            td.textContent = 'Keine Gruppen – Schritt 1 prüfen.';
            tr.appendChild(td);
            tbody.appendChild(tr);
        }
    }

    function goToStep(step) {
        const n = Number(step);
        const contents = document.querySelectorAll('.ku-step-content');
        for (let i = 0; i < contents.length; i++) {
            const el = contents[i];
            const s = el.getAttribute('data-ku-step');
            if (String(s) === String(n)) el.classList.add('active');
            else el.classList.remove('active');
        }
        const steps = document.querySelectorAll('.ku-steps .step');
        for (let j = 0; j < steps.length; j++) {
            const st = steps[j];
            const s = st.getAttribute('data-ku-step');
            if (String(s) === String(n)) st.classList.add('active');
            else st.classList.remove('active');
        }
        const bar = document.getElementById('kuStepsBar');
        if (bar && typeof window.ms365ApplyStepProgress === 'function') {
            window.ms365ApplyStepProgress(bar, n, STEP_ORDER);
        }
    }

    async function loadGroups() {
        const status = document.getElementById('kuLoadStatus');
        const next1 = document.getElementById('kuNext1');
        if (next1) next1.disabled = true;
        if (status) status.textContent = 'Lade …';
        try {
            const token = await getGraphToken();
            const filter = encodeURIComponent("groupTypes/any(c:c eq 'Unified')");
            const initial =
                '/groups?$filter=' + filter + '&$select=' + encodeURIComponent(GROUP_SELECT) + '&$top=999';
            loadedGroups = await fetchAllPages(token, initial);
            loadedGroups.sort(function (a, b) {
                const an = a && a.displayName ? String(a.displayName) : '';
                const bn = b && b.displayName ? String(b.displayName) : '';
                return an.localeCompare(bn, 'de');
            });
            const n = getFilteredGroups().length;
            if (status) {
                status.textContent =
                    'Gesamt ' +
                    loadedGroups.length +
                    ' einheitliche Gruppe(n). Nach Filter: ' +
                    n +
                    ' sichtbar für die nächsten Schritte.';
            }
            if (next1) next1.disabled = loadedGroups.length === 0;
            toast('Gruppen geladen.');
        } catch (e) {
            if (status) status.textContent = '';
            toast(String(e && e.message ? e.message : e));
        }
    }

    async function runRename() {
        const updNick = document.getElementById('kuUpdateMailNick');
        clearLog();
        const toPatch = previewRows.filter(function (r) {
            return r.ok;
        });
        if (!toPatch.length) {
            toast('Keine gültigen Zeilen zum Umbenennen.');
            return;
        }
        appendLog('Start: ' + toPatch.length + ' Gruppe(n).');
        let ok = 0;
        let fail = 0;
        try {
            const token = await getGraphToken();
            for (let i = 0; i < toPatch.length; i++) {
                const r = toPatch[i];
                const body = {};
                if (r.displayNameNew && r.displayNameNew !== r.displayNameOld) {
                    body.displayName = r.displayNameNew;
                }
                if (updNick && updNick.checked && r.mailNicknameNew && r.mailNicknameNew !== r.mailNicknameOld) {
                    body.mailNickname = r.mailNicknameNew;
                }
                if (!Object.keys(body).length) {
                    appendLog('Übersprungen (keine Felder): ' + r.displayNameOld, 'warn');
                    continue;
                }
                try {
                    await graphJson('PATCH', '/groups/' + encodeURIComponent(r.id), token, body);
                    appendLog('OK: ' + r.displayNameOld + ' → ' + r.displayNameNew, 'ok');
                    ok++;
                } catch (e) {
                    fail++;
                    appendLog(
                        'Fehler ' +
                            r.displayNameOld +
                            ': ' +
                            (e && e.message ? e.message : e),
                        'err'
                    );
                }
            }
            appendLog('Fertig. Erfolg: ' + ok + ', Fehler: ' + fail + '.', fail ? 'warn' : 'ok');
            toast('Umbenennung abgeschlossen.');
        } catch (e) {
            appendLog(String(e && e.message ? e.message : e), 'err');
        }
    }

    async function onLogin() {
        try {
            await getGraphToken();
            toast('Angemeldet.');
        } catch (e) {
            toast(String(e && e.message ? e.message : e));
        }
    }

    function bind() {
        const btnLogin = document.getElementById('kuBtnLogin');
        const btnLoad = document.getElementById('kuBtnLoad');
        const next1 = document.getElementById('kuNext1');
        const next2 = document.getElementById('kuNext2');
        const next3 = document.getElementById('kuNext3');
        const back2 = document.getElementById('kuBack2');
        const back3 = document.getElementById('kuBack3');
        const back4 = document.getElementById('kuBack4');
        const btnRun = document.getElementById('kuBtnRun');
        const filt = document.getElementById('kuFilterContains');
        const onlyTeams = document.getElementById('kuOnlyTeams');

        if (btnLogin) btnLogin.addEventListener('click', () => onLogin());
        if (btnLoad) btnLoad.addEventListener('click', () => loadGroups());
        function refreshFilterStatus() {
            const status = document.getElementById('kuLoadStatus');
            const next1b = document.getElementById('kuNext1');
            if (!loadedGroups.length) return;
            const n = getFilteredGroups().length;
            if (status) {
                status.textContent =
                    'Gesamt ' +
                    loadedGroups.length +
                    ' Gruppe(n). Nach Filter: ' +
                    n +
                    ' sichtbar.';
            }
            if (next1b) next1b.disabled = false;
        }

        if (filt)
            filt.addEventListener('input', function () {
                if (loadedGroups.length) refreshFilterStatus();
            });
        if (onlyTeams)
            onlyTeams.addEventListener('change', function () {
                if (loadedGroups.length) refreshFilterStatus();
            });

        if (next1) {
            next1.addEventListener('click', function () {
                if (!loadedGroups.length) {
                    toast('Bitte zuerst Gruppen laden.');
                    return;
                }
                goToStep(2);
            });
        }
        if (back2) back2.addEventListener('click', () => goToStep(1));
        if (next2) {
            next2.addEventListener('click', async function () {
                try {
                    const ok = await buildPreview();
                    if (ok) goToStep(3);
                } catch (e) {
                    toast(String(e && e.message ? e.message : e));
                }
            });
        }
        if (back3) back3.addEventListener('click', () => goToStep(2));
        if (next3) {
            next3.addEventListener('click', function () {
                goToStep(4);
            });
        }
        if (back4) back4.addEventListener('click', () => goToStep(3));
        if (btnRun) btnRun.addEventListener('click', () => runRename());

        goToStep(1);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bind);
    } else {
        bind();
    }
})();
