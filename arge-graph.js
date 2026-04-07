(function () {
    'use strict';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/Group.ReadWrite.All',
        'https://graph.microsoft.com/User.Read.All'
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

    async function getPca() {
        const m = await loadMsal();
        const PublicClientApplication = m.PublicClientApplication || (m.default && m.default.PublicClientApplication);
        if (!PublicClientApplication) {
            throw new Error('MSAL: PublicClientApplication nicht gefunden (Import).');
        }
        const cfg = window.MS365_MSAL_CONFIG;
        if (!cfg || !String(cfg.clientId || '').trim()) {
            throw new Error('Bitte clientId in ms365-config.js eintragen (Entra-App-Registrierung).');
        }
        if (!pca) {
            pca = new PublicClientApplication({
                auth: {
                    clientId: cfg.clientId.trim(),
                    authority: cfg.authority || 'https://login.microsoftonline.com/organizations',
                    redirectUri: cfg.redirectUri || window.location.href.split('#')[0]
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

    function appendLog(msg, kind) {
        const el = document.getElementById('argeOnlineLog');
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
        const el = document.getElementById('argeOnlineLog');
        if (el) el.replaceChildren();
    }

    async function runArgeOnline() {
        const snapshot = window.ms365GetArgeSnapshotForGraph;
        if (typeof snapshot !== 'function') {
            appendLog('Interner Fehler: ARGE-Daten nicht verfügbar.', 'err');
            return;
        }
        const pack = snapshot();
        if (!pack || !pack.rows || !pack.rows.length) {
            appendLog('Keine ARGE-Zeilen – bitte Schritte 1–3 abschließen.', 'err');
            return;
        }
        const missing = pack.rows.filter(function (r) {
            return !r.owner;
        });
        if (missing.length) {
            appendLog('Bitte für alle ARGEs einen Besitzer (UPN) eintragen.', 'err');
            return;
        }
        if (pack.exchangeSmtp) {
            appendLog(
                'Hinweis: „Primäre SMTP per Exchange“ ist aktiv – die Online-Ausführung setzt das nicht (nur PowerShell/CMD). Graph legt den Mail-Nickname an.',
                'warn'
            );
        }

        const btnLogin = document.getElementById('argeOnlineLogin');
        const btnRun = document.getElementById('argeOnlineRun');
        if (btnRun) btnRun.disabled = true;
        if (btnLogin) btnLogin.disabled = true;

        clearLog();
        appendLog('Start – Microsoft Graph (Browser) …');

        let token;
        try {
            token = await getGraphToken();
        } catch (e) {
            appendLog('Anmeldung/Token: ' + (e.message || e), 'err');
            if (btnRun) btnRun.disabled = false;
            if (btnLogin) btnLogin.disabled = false;
            return;
        }

        const teamBody = {
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

        const total = pack.rows.length;
        let i = 0;
        for (const r of pack.rows) {
            i++;
            try {
                appendLog('[' + i + '/' + total + '] ' + r.displayName + ' …');

                const owner = await graphJson(
                    'GET',
                    '/users/' + encodeURIComponent(r.owner),
                    token,
                    undefined
                );
                const ownerId = owner.id;

                const groupBody = {
                    displayName: r.displayName,
                    description: r.description,
                    mailNickname: r.mailNick,
                    mailEnabled: true,
                    securityEnabled: false,
                    groupTypes: ['Unified'],
                    visibility: 'Private'
                };

                const group = await graphJson('POST', '/groups', token, groupBody);
                const gid = group.id;

                await sleep(2000);

                await graphJson('POST', '/groups/' + gid + '/owners/$ref', token, {
                    '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + ownerId
                });

                try {
                    await graphJson('POST', '/groups/' + gid + '/members/$ref', token, {
                        '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + ownerId
                    });
                } catch (e) {
                    appendLog('  Hinweis (Besitzer als Mitglied): ' + e.message, 'warn');
                }

                if (pack.createTeams) {
                    const teamUri = '/groups/' + gid + '/team';
                    let teamOk = false;
                    for (let ti = 0; ti < 8; ti++) {
                        try {
                            await graphJson('PUT', teamUri, token, teamBody);
                            appendLog('  Teams: Team bereitgestellt.', 'ok');
                            teamOk = true;
                            break;
                        } catch (e) {
                            if (ti < 7) {
                                appendLog('  Teams: Warte auf Replikation (' + (ti + 1) + '/8) …', 'warn');
                                await sleep(10000);
                                token = await getGraphToken();
                            } else {
                                appendLog('  Teams: ' + e.message, 'err');
                            }
                        }
                    }
                    if (!teamOk) {
                        /* Fehler bereits protokolliert */
                    }
                }

                appendLog('OK [' + i + '/' + total + '] ' + r.displayName + ' → ' + r.mailNick, 'ok');
            } catch (e) {
                appendLog('Fehler [' + i + '/' + total + '] ' + r.displayName + ': ' + (e.message || e), 'err');
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
        const btnLogin = document.getElementById('argeOnlineLogin');
        if (btnLogin) btnLogin.disabled = true;
        try {
            await getGraphToken();
            toast('Microsoft angemeldet – Sie können jetzt ausführen.');
        } catch (e) {
            toast('Anmeldung: ' + (e.message || e));
        } finally {
            if (btnLogin) btnLogin.disabled = false;
        }
    }

    window.ms365ArgeGraphLogin = loginOnly;
    window.ms365ArgeGraphRun = runArgeOnline;
})();
