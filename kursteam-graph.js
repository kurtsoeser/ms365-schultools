(function () {
    'use strict';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.ReadWrite.All',
        'https://graph.microsoft.com/User.Read.All',
        /** POST /teams (Kursteam wie New-Team -Template EDU_Class → teamsTemplates educationClass) */
        'https://graph.microsoft.com/Team.Create'
    ];

    let msalMod = null;
    let pca = null;

    function toast(msg) {
        if (typeof window.ms365ShowToast === 'function') {
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
        if (!cfg) {
            cfg = {};
        }
        let id = String(cfg.clientId || '').trim();
        if (!id) {
            const meta = document.querySelector('meta[name="ms365-graph-client-id"]');
            const fromMeta = meta && meta.getAttribute('content') ? meta.getAttribute('content').trim() : '';
            if (fromMeta) {
                id = fromMeta;
            }
        }
        if (!id) {
            throw new Error(
                'Keine clientId: ms365-config.js fehlt/leer oder blockiert. Seite mit Strg+F5 neu laden; im Netzwerk-Tab prüfen, ob ms365-config.js mit 200 lädt. Alternativ meta ms365-graph-client-id in ms365-schooltool.html setzen (Entra-Anwendungs-ID).'
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

    async function graphRequest(method, path, token, body) {
        const url = path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path;
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

    async function graphJson(method, path, token, body) {
        const res = await graphRequest(method, path, token, body);
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

    function isGraphDuplicateRefError(err) {
        const msg = String(err && err.message ? err.message : err);
        return /already exist/i.test(msg) || /already exists/i.test(msg);
    }

    function appendLog(msg, kind) {
        const el = document.getElementById('kursteamOnlineLog');
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
        const el = document.getElementById('kursteamOnlineLog');
        if (el) el.replaceChildren();
    }

    /**
     * Microsoft: PowerShell New-Team -Template "EDU_Class" entspricht Graph teamsTemplates('educationClass').
     * Richtiger Weg: POST /teams mit group@odata.bind (siehe team-post, Beispiel 4/6), nicht nur PUT …/team.
     */
    function parseTeamsOperationPath(locationHeader) {
        if (!locationHeader) return null;
        const loc = String(locationHeader).trim();
        const m = loc.match(/teams\('([^']+)'\)\/operations\('([^']+)'\)/i);
        if (m) return '/teams/' + m[1] + '/operations/' + m[2];
        const m2 = loc.match(/\/teams\/([^/]+)\/operations\/([^/?\s]+)/i);
        if (m2) return '/teams/' + m2[1] + '/operations/' + m2[2];
        return null;
    }

    async function pollTeamsAsyncOperation(token, operationPath, appendLog) {
        const maxAttempts = 120;
        for (let i = 0; i < maxAttempts; i++) {
            await sleep(2000);
            const data = await graphJson('GET', operationPath, token, undefined);
            const st = String(data.status || data.Status || '').toLowerCase();
            if (st === 'succeeded') {
                appendLog('  Teams: Bereitstellung abgeschlossen (Template educationClass).', 'ok');
                return;
            }
            if (st === 'failed') {
                const errMsg =
                    (data.error && (data.error.message || JSON.stringify(data.error))) ||
                    JSON.stringify(data);
                throw new Error('Team-Bereitstellung fehlgeschlagen: ' + errMsg);
            }
            if (i > 0 && i % 8 === 0) {
                appendLog('  Teams: Warte auf Bereitstellung … (' + i * 2 + ' s)', 'warn');
            }
        }
        throw new Error('Timeout: Team-Bereitstellung (async) nicht abgeschlossen.');
    }

    /** Fallback wie Jahrgang/ARGE: Team an bestehende Gruppe hängen (ohne educationClass-Template). */
    function buildPutTeamBody(includeSpecialization) {
        const base = {
            memberSettings: {
                allowCreatePrivateChannels: true,
                allowCreateUpdateChannels: true
            },
            messagingSettings: {
                allowUserEditMessages: true,
                allowUserDeleteMessages: true
            },
            funSettings: {
                allowGiphy: true,
                giphyContentRating: 'moderate'
            },
            guestSettings: {
                allowCreateUpdateChannels: false
            }
        };
        if (includeSpecialization) {
            base.specialization = 'educationClass';
        }
        return base;
    }

    /**
     * Primär: POST /teams mit teamsTemplates('educationClass') + group@odata.bind — entspricht New-Team -Template EDU_Class.
     * Hinweis: POST /education/classes (educationClass) ist laut Microsoft Learn für delegierte Auth nicht vorgesehen und legt
     * ohnehin kein Team an (nur Klassen-Gruppe); siehe https://learn.microsoft.com/de-de/graph/api/educationclass-post
     * Fallback: PUT /groups/{id}/team (Replikations-Wiederholungen).
     */
    async function provisionKursteamTeam(token, gid, appendLog) {
        const postBody = {
            'template@odata.bind':
                'https://graph.microsoft.com/v1.0/teamsTemplates(\'educationClass\')',
            'group@odata.bind': 'https://graph.microsoft.com/v1.0/groups(\'' + gid + '\')'
        };

        let lastPostErr = null;
        for (let attempt = 0; attempt < 3; attempt++) {
            try {
                const res = await graphRequest('POST', '/teams', token, postBody);
                const text = await res.text();
                if (res.status === 202 || res.status === 200) {
                    const loc = res.headers.get('Location') || res.headers.get('Content-Location');
                    const opPath = parseTeamsOperationPath(loc);
                    if (opPath) {
                        appendLog('  Teams: educationClass-Anlage gestartet (POST /teams) …', 'warn');
                        await pollTeamsAsyncOperation(token, opPath, appendLog);
                    } else {
                        appendLog(
                            '  Teams: POST /teams angenommen (keine Operation-URL – ggf. im Admin prüfen).',
                            'warn'
                        );
                    }
                    return await getGraphToken();
                }
                if (res.status === 404 && attempt < 2) {
                    appendLog(
                        '  Teams: 404 nach Gruppenerstellung – Replikation, Warte 10 s …',
                        'warn'
                    );
                    await sleep(10000);
                    token = await getGraphToken();
                    continue;
                }
                lastPostErr = new Error('POST /teams: ' + res.status + ' ' + (text || ''));
                break;
            } catch (e) {
                lastPostErr = e;
                if (attempt < 2 && /404/.test(String(e.message))) {
                    appendLog('  Teams: Wiederholung nach Wartezeit (404) …', 'warn');
                    await sleep(10000);
                    token = await getGraphToken();
                    continue;
                }
                break;
            }
        }

        appendLog(
            '  Teams: POST /teams (educationClass) nicht möglich: ' +
                (lastPostErr && lastPostErr.message ? lastPostErr.message : '') +
                ' – Fallback PUT …/team.',
            'warn'
        );

        const teamUri = '/groups/' + gid + '/team';
        let useEducationSpec = true;
        for (let ti = 0; ti < 8; ti++) {
            try {
                await graphJson('PUT', teamUri, token, buildPutTeamBody(useEducationSpec));
                appendLog('  Teams: Team per PUT bereitgestellt (Fallback).', 'ok');
                return await getGraphToken();
            } catch (e) {
                if (useEducationSpec) {
                    useEducationSpec = false;
                    appendLog('  Teams: PUT ohne specialization educationClass …', 'warn');
                    ti--;
                    continue;
                }
                if (ti < 7) {
                    appendLog('  Teams: Warte auf Replikation (' + (ti + 1) + '/8) …', 'warn');
                    await sleep(10000);
                    token = await getGraphToken();
                } else {
                    throw e;
                }
            }
        }
    }

    async function runKursteamOnline() {
        const snapshotFn = window.ms365GetKursteamSnapshotForGraph;
        if (typeof snapshotFn !== 'function') {
            appendLog('Interner Fehler: Kursteam-Daten nicht verfügbar.', 'err');
            return;
        }
        const pack = snapshotFn();
        if (!pack || !pack.teams || !pack.teams.length) {
            appendLog('Keine gültigen Teams – bitte in Schritt „Teams konfigurieren“ generieren und prüfen.', 'err');
            return;
        }
        const missing = pack.teams.filter(function (t) {
            return !t.besitzer;
        });
        if (missing.length) {
            appendLog('Bitte für alle Teams einen gültigen Besitzer (E-Mail / UPN) im Mandanten eintragen.', 'err');
            return;
        }

        const btnLogin = document.getElementById('kursteamOnlineLogin');
        const btnRun = document.getElementById('kursteamOnlineRun');
        if (btnRun) btnRun.disabled = true;
        if (btnLogin) btnLogin.disabled = true;

        clearLog();
        appendLog('Start – Microsoft Graph (Browser), Kursteams …');
        appendLog(
            'Hinweis: Wie PowerShell New-Team -Template EDU_Class → Graph teamsTemplates(\'educationClass\') + POST /teams (siehe Microsoft Learn: Create team, Beispiel 4/6).',
            'warn'
        );

        let token;
        try {
            token = await getGraphToken();
        } catch (e) {
            appendLog('Anmeldung/Token: ' + (e.message || e), 'err');
            if (btnRun) btnRun.disabled = false;
            if (btnLogin) btnLogin.disabled = false;
            return;
        }

        const total = pack.teams.length;
        let i = 0;
        for (const t of pack.teams) {
            i++;
            try {
                appendLog('[' + i + '/' + total + '] ' + t.teamName + ' …');

                const owner = await graphJson(
                    'GET',
                    '/users/' + encodeURIComponent(t.besitzer),
                    token,
                    undefined
                );
                const ownerId = owner.id;

                const groupBody = {
                    displayName: t.teamName,
                    description: 'Kursteam (WebUntis / MS365-Schulverwaltung)',
                    mailNickname: t.gruppenmail,
                    mailEnabled: true,
                    securityEnabled: false,
                    groupTypes: ['Unified'],
                    visibility: 'Private'
                };

                const group = await graphJson('POST', '/groups', token, groupBody);
                const gid = group.id;

                await sleep(2000);

                try {
                    await graphJson('POST', '/groups/' + gid + '/owners/$ref', token, {
                        '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + ownerId
                    });
                } catch (e) {
                    if (isGraphDuplicateRefError(e)) {
                        appendLog(
                            '  Besitzer: bereits gesetzt (häufig, wenn gleicher Admin wie angemeldeter Benutzer).',
                            'warn'
                        );
                    } else {
                        throw e;
                    }
                }

                try {
                    await graphJson('POST', '/groups/' + gid + '/members/$ref', token, {
                        '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + ownerId
                    });
                } catch (e) {
                    if (isGraphDuplicateRefError(e)) {
                        appendLog('  Mitglied: bereits gesetzt.', 'warn');
                    } else {
                        appendLog('  Hinweis (Besitzer als Mitglied): ' + e.message, 'warn');
                    }
                }

                token = await provisionKursteamTeam(token, gid, appendLog);

                appendLog('OK [' + i + '/' + total + '] ' + t.teamName + ' → ' + t.gruppenmail, 'ok');
            } catch (e) {
                appendLog('Fehler [' + i + '/' + total + '] ' + t.teamName + ': ' + (e.message || e), 'err');
            }

            await sleep(2000);
            try {
                token = await getGraphToken();
            } catch (e) {
                appendLog('Token erneuern: ' + (e.message || e), 'err');
                break;
            }
        }

        appendLog('Fertig.', 'ok');
        if (btnRun) btnRun.disabled = false;
        if (btnLogin) btnLogin.disabled = false;
    }

    async function loginOnly() {
        const btnLogin = document.getElementById('kursteamOnlineLogin');
        if (btnLogin) btnLogin.disabled = true;
        try {
            await getGraphToken();
            toast('Microsoft angemeldet – Sie können jetzt Kursteams anlegen.');
        } catch (e) {
            toast('Anmeldung: ' + (e.message || e));
        } finally {
            if (btnLogin) btnLogin.disabled = false;
        }
    }

    window.ms365KursteamGraphLogin = loginOnly;
    window.ms365KursteamGraphRun = runKursteamOnline;
})();
