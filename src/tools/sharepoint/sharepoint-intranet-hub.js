(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    if (!G) return;

    const SCOPES_GRAPH_SITE = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.Read.All',
        'https://graph.microsoft.com/Sites.Create.All'
    ];

    function $(id) {
        return document.getElementById(id);
    }

    function toast(msg) {
        if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(msg);
        } else {
            window.alert(msg);
        }
    }

    function slugify(s) {
        return String(s || '')
            .trim()
            .toLowerCase()
            .replace(/[^a-z0-9-]/g, '-')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '');
    }

    function adminHostFromSpoHost(h) {
        const host = String(h || '').trim().toLowerCase();
        if (!host || host.indexOf('.') === -1) return '';
        const i = host.indexOf('.sharepoint.com');
        if (i === -1) return '';
        const name = host.slice(0, i);
        return name + '-admin.sharepoint.com';
    }

    function buildWebUrl(host, slug) {
        const h = String(host || '').trim().toLowerCase();
        const s = slugify(slug);
        if (!h || !s) return '';
        return 'https://' + h + '/sites/' + s;
    }

    function normalizeLocale(locale) {
        const raw = String(locale || '').trim() || 'de-DE';
        const m = raw.match(/^([a-zA-Z]{2,3})(?:[_-]([a-zA-Z]{2}))?$/);
        if (!m) return raw;
        return m[2] ? m[1].toLowerCase() + '-' + m[2].toUpperCase() : m[1].toLowerCase();
    }

    function signedInUpn() {
        try {
            if (typeof window.ms365AuthGetUserPrincipalName === 'function') {
                const u = String(window.ms365AuthGetUserPrincipalName() || '').trim();
                if (u) return u;
            }
            if (typeof window.ms365AuthGetAccountInfo === 'function') {
                const info = window.ms365AuthGetAccountInfo();
                const u = info && (info.upn || info.username) ? String(info.upn || info.username).trim() : '';
                if (u) return u;
            }
        } catch {
            /* ignore */
        }
        return '';
    }

    function fillOwnerFromAccount() {
        const el = $('fOwner');
        if (!el || String(el.value || '').trim()) return;
        const u = signedInUpn();
        if (u) el.value = u;
    }

    function explainCreateSiteError(status, resText, owner, host) {
        let code = '';
        let message = '';
        try {
            const j = JSON.parse(resText || '{}');
            code = j && j.error && j.error.code ? String(j.error.code) : '';
            message = j && j.error && j.error.message ? String(j.error.message) : '';
        } catch {
            /* ignore */
        }
        const blob = (code + ' ' + message + ' ' + (resText || '')).toLowerCase();
        if (status === 404 || /itemnotfound|item not found|resource could not be found/.test(blob)) {
            return (
                'Site-Erstellung: HTTP ' +
                status +
                ' (itemNotFound).\n\n' +
                'Häufige Ursachen bei neuen Mandanten:\n' +
                '1) SharePoint ist noch nicht bereit – einmal https://' +
                (host || 'TENANT') +
                ' und das SharePoint Admin Center öffnen, 1–2 Min. warten, dann erneut versuchen.\n' +
                '2) Besitzer „' +
                owner +
                '“ existiert nicht in diesem Tenant (UPN/E-Mail exakt wie in Entra ID).\n' +
                '3) SharePoint-Host falsch (Hostname ermitteln nutzen).\n\n' +
                'Rohantwort: ' +
                (resText || '')
            );
        }
        if (status === 403 || /accessdenied|forbidden|authorization_requestdenied/.test(blob)) {
            return (
                'Site-Erstellung: HTTP ' +
                status +
                ' – keine Berechtigung (Sites.Create.All / Admin-Zustimmung fehlt?).\n' +
                (resText || '')
            );
        }
        return 'Site-Erstellung: HTTP ' + status + ' ' + (resText || '');
    }

    /**
     * Prüft Root-Site + Besitzer, bevor Graph die Site anlegt.
     * @returns {Promise<{ ownerEmail: string }>}
     */
    async function preflightCreateSite(token, host, owner) {
        try {
            await G.graphJson('GET', '/sites/root?$select=id,webUrl,displayName', token, undefined, 'v1.0');
        } catch (e) {
            const msg = String(e && e.message ? e.message : e);
            const st = e && e.status;
            if (st === 404 || /itemNotFound|not found|nicht gefunden/i.test(msg)) {
                throw new Error(
                    'SharePoint ist in diesem Mandanten noch nicht bereit (Root-Site fehlt).\n' +
                        'Bitte einmal https://' +
                        host +
                        ' und https://' +
                        String(host || '').replace(/\.sharepoint\.com$/i, '-admin.sharepoint.com') +
                        ' im Browser öffnen (als Global/SharePoint-Admin), 1–2 Minuten warten, dann „Hostname ermitteln“ und erneut anlegen.\n' +
                        'Details: ' +
                        msg
                );
            }
            throw new Error('SharePoint-Root nicht lesbar: ' + msg);
        }

        const ownerEmail = String(owner || '').trim();
        if (!ownerEmail) throw new Error('Besitzer (UPN/E-Mail) ist erforderlich.');

        const meUpn = signedInUpn().toLowerCase();
        if (meUpn && ownerEmail.toLowerCase() === meUpn) {
            try {
                await G.graphJson('GET', '/me?$select=id,userPrincipalName,mail', token, undefined, 'v1.0');
                return { ownerEmail: ownerEmail };
            } catch {
                /* weiter mit /users */
            }
        }

        try {
            await G.graphJson(
                'GET',
                '/users/' + encodeURIComponent(ownerEmail) + '?$select=id,userPrincipalName,mail',
                token,
                undefined,
                'v1.0'
            );
        } catch (e) {
            const msg = String(e && e.message ? e.message : e);
            const st = e && e.status;
            if (st === 404 || /itemNotFound|Resource.*not found|does not exist/i.test(msg)) {
                throw new Error(
                    'Besitzer „' +
                        ownerEmail +
                        '“ wurde in Entra ID nicht gefunden.\n' +
                        'UPN exakt wie im Microsoft 365 Admin Center eintragen (oft onmicrosoft.com oder die Schul-Domain). Angemeldet: ' +
                        (signedInUpn() || '–')
                );
            }
            if (st === 403 || /Authorization_RequestDenied|Insufficient privileges/i.test(msg)) {
                /* User.Read.All fehlt oft – Create trotzdem versuchen */
                return { ownerEmail: ownerEmail };
            }
            throw new Error('Besitzer-Prüfung fehlgeschlagen: ' + msg);
        }
        return { ownerEmail: ownerEmail };
    }

    function adminCenterSitesUrl(host) {
        const admin = adminHostFromSpoHost(host);
        if (!admin) return 'https://admin.microsoft.com/sharepoint';
        return 'https://' + admin + '/_layouts/15/online/AdminHome.aspx#/siteManagement';
    }

    function setPsScript(siteUrl) {
        let host = String($('fHost').value || '').trim();
        if (!host) {
            try {
                host = new URL(siteUrl).hostname;
            } catch {
                host = '';
            }
        }
        const admin = adminHostFromSpoHost(host);
        const siteQ = JSON.stringify(siteUrl);
        const ps =
            '# Hub-Website registrieren (SharePoint Online PowerShell)\n' +
            '# Einmalig: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser\n' +
            'Connect-SPOService -Url https://' +
            admin +
            '\n' +
            'Register-SPOHubSite -Site ' +
            siteQ +
            '\n' +
            '\n' +
            '# Alternativ mit PnP.PowerShell:\n' +
            '# Install-Module PnP.PowerShell -Scope CurrentUser\n' +
            '# Connect-PnPOnline -Url https://' +
            admin +
            ' -Interactive\n' +
            '# Register-PnPHubSite -Site ' +
            siteQ +
            '\n';
        if ($('fPsHub')) $('fPsHub').value = ps;
        const adminLink = $('fAdminHubLink');
        if (adminLink) {
            adminLink.href = adminCenterSitesUrl(host);
            adminLink.textContent = 'SharePoint Admin Center – Aktive Sites';
        }
    }

    function revealHubFallback() {
        const det = $('ihHubTech');
        if (det) det.open = true;
        const ps = $('fPsHub');
        if (ps) {
            try {
                ps.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            } catch {
                /* ignore */
            }
        }
    }

    async function detectHost() {
        $('fLog').textContent = 'Ermittle SharePoint-Host …';
        const token = await G.getGraphToken(SCOPES_GRAPH_SITE);
        const host = await G.getSharePointHostname(token);
        if (!host) throw new Error('Hostname konnte nicht ermittelt werden.');
        $('fHost').value = host;
        $('fLog').textContent = 'SharePoint-Host: ' + host;
        toast('Hostname gesetzt.');
    }

    async function createSite() {
        fillOwnerFromAccount();
        const host = String($('fHost').value || '').trim();
        if (!host) {
            await detectHost();
        }
        const h2 = String($('fHost').value || '').trim();
        const slug = $('fSlug').value;
        const webUrl = buildWebUrl(h2, slug);
        if (!webUrl) throw new Error('Website-Adresse ungültig (Host + Kurzname prüfen).');

        const title = String($('fTitle').value || '').trim() || 'Intranet';
        const description = String($('fDesc').value || '').trim();
        const locale = normalizeLocale($('fLocale').value || 'de-DE');
        if ($('fLocale')) $('fLocale').value = locale;
        let owner = String($('fOwner').value || '').trim();
        if (!owner) {
            owner = signedInUpn();
            if (owner && $('fOwner')) $('fOwner').value = owner;
        }
        if (!owner) throw new Error('Besitzer (UPN/E-Mail) ist erforderlich.');

        $('fLog').textContent = 'Prüfe SharePoint-Root und Besitzer …';
        const token = await G.getGraphToken(SCOPES_GRAPH_SITE);
        const pre = await preflightCreateSite(token, h2, owner);
        owner = pre.ownerEmail;

        $('fLog').textContent = 'Erstelle Kommunikationswebsite (Graph beta) …\nURL: ' + webUrl + '\nBesitzer: ' + owner;
        const body = {
            name: title,
            description: description,
            webUrl: webUrl,
            locale: locale,
            shareByEmailEnabled: $('fShareByMail').checked,
            template: 'sitepagepublishing',
            ownerIdentityToResolve: { email: owner }
        };

        const res = await G.graphRequest('POST', '/sites', token, body, 'beta');
        const resText = await res.text();
        let resJson = null;
        if (resText) {
            try {
                resJson = JSON.parse(resText);
            } catch {
                resJson = { raw: resText };
            }
        }

        if (res.status !== 202 && res.status !== 200) {
            throw new Error(explainCreateSiteError(res.status, resText, owner, h2));
        }

        let opUrl = res.headers.get('Location') || res.headers.get('location') || '';
        if (!opUrl && resJson && resJson.location) opUrl = resJson.location;
        if (opUrl && opUrl.indexOf('http') !== 0) {
            opUrl = 'https://graph.microsoft.com' + (opUrl.indexOf('/') === 0 ? '' : '/') + opUrl;
        }
        if (!opUrl) {
            $('fJson').textContent = JSON.stringify(resJson, null, 2);
            throw new Error('Keine Operation-URL (Location-Header). Antwort siehe JSON.');
        }

        $('fLog').textContent = 'Vorgang gestartet, warte auf Abschluss …\n' + opUrl;
        const done = await G.pollRichLongRunningOperation(opUrl, token, { siteWebUrl: webUrl });
        $('fJson').textContent = JSON.stringify({ createResponse: resJson, operation: done }, null, 2);
        $('fLastSiteUrl').value = webUrl;
        if ($('fManualUrl')) $('fManualUrl').value = webUrl;
        setPsScript(webUrl);
        $('fLog').textContent = 'Site bereit (laut Vorgang): ' + webUrl;
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({ intranetSiteUrl: webUrl, schoolIntranetSiteUrl: webUrl });
            }
        } catch {
            /* ignore */
        }
        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'sharepoint',
                action: 'create-site',
                target: webUrl,
                summary: 'Kommunikationssite angelegt'
            });
        }
        toast('Kommunikationswebsite erstellt.');
        await maybeCreateStartpaket(webUrl);
        return webUrl;
    }

    function packLog(msg) {
        const el = $('fLog');
        if (!el) return;
        el.textContent += (el.textContent ? '\n' : '') + msg;
    }

    async function maybeCreateStartpaket(webUrl) {
        const wantLehrer = $('fPackLehrer') && $('fPackLehrer').checked;
        const wantTermine = $('fPackTermine') && $('fPackTermine').checked;
        const wantSchularbeiten = $('fPackSchularbeiten') && $('fPackSchularbeiten').checked;
        const wantProjektwochen = $('fPackProjektwochen') && $('fPackProjektwochen').checked;
        const wantStammdaten = $('fPackStammdaten') && $('fPackStammdaten').checked;
        if (!wantLehrer && !wantTermine && !wantSchularbeiten && !wantProjektwochen && !wantStammdaten) return;
        if (wantLehrer && window.ms365SpoLehrerListe && typeof window.ms365SpoLehrerListe.createList === 'function') {
            packLog('Startpaket: Lehrerliste …');
            try {
                await window.ms365SpoLehrerListe.createList(webUrl, 'Lehrerinnen', packLog);
                packLog('Lehrerliste fertig.');
            } catch (e) {
                packLog('Lehrerliste: ' + (e && e.message ? e.message : e));
            }
        }
        if (wantTermine && window.ms365SpoSchultermine && typeof window.ms365SpoSchultermine.createList === 'function') {
            packLog('Startpaket: Schultermine …');
            try {
                await window.ms365SpoSchultermine.createList(webUrl, 'Schultermine', packLog);
                packLog('Schultermine-Liste fertig.');
            } catch (e) {
                packLog('Schultermine: ' + (e && e.message ? e.message : e));
            }
        }
        if (
            wantStammdaten &&
            window.ms365SpoStammdatenListen &&
            typeof window.ms365SpoStammdatenListen.createSelectedLists === 'function'
        ) {
            packLog('Startpaket: Stammdaten-Listen (Schülerinnen, Fächer) …');
            try {
                await window.ms365SpoStammdatenListen.syncSelectedLists(
                    webUrl,
                    { schueler: true, faecher: true, fachgruppen: true, arges: true, klassen: false },
                    packLog,
                    { syncMode: true, removeOrphans: false }
                );
                packLog('Stammdaten-Listen fertig.');
            } catch (e) {
                packLog('Stammdaten-Listen: ' + (e && e.message ? e.message : e));
            }
        }
        if (
            wantSchularbeiten &&
            window.ms365SpoSchularbeiten &&
            typeof window.ms365SpoSchularbeiten.createLists === 'function'
        ) {
            packLog('Startpaket: Schularbeiten …');
            try {
                await window.ms365SpoSchularbeiten.createLists(webUrl, packLog);
                packLog('Schularbeiten-Paket fertig.');
            } catch (e) {
                packLog('Schularbeiten: ' + (e && e.message ? e.message : e));
            }
        }
        if (
            wantProjektwochen &&
            window.ms365SpoProjektwochen &&
            typeof window.ms365SpoProjektwochen.createLists === 'function'
        ) {
            packLog('Startpaket: Projektwochen …');
            try {
                await window.ms365SpoProjektwochen.createLists(webUrl, packLog);
                packLog('Projektwochen-Paket fertig.');
            } catch (e) {
                packLog('Projektwochen: ' + (e && e.message ? e.message : e));
            }
        }
    }

    async function registerHub() {
        const siteUrl = String($('fLastSiteUrl').value || $('fManualUrl').value || '').trim();
        if (!siteUrl) throw new Error('Zuerst Site erstellen oder Site-URL eintragen.');
        if ($('fManualUrl') && !$('fManualUrl').value) $('fManualUrl').value = siteUrl;

        let host = String($('fHost').value || '').trim();
        if (!host) {
            try {
                host = new URL(siteUrl).hostname;
            } catch {
                host = '';
            }
        }
        if (!host) throw new Error('SharePoint-Host fehlt (Hostname ermitteln oder vollständige Site-URL eintragen).');

        setPsScript(siteUrl);
        $('fLog').textContent = 'Hole SharePoint-Token und versuche Hub-Registrierung (REST) …';
        const spoScope = 'https://' + host + '/Sites.FullControl.All';
        let spoToken;
        try {
            spoToken = await G.getGraphToken([spoScope]);
        } catch (e) {
            revealHubFallback();
            $('fLog').textContent =
                'SharePoint-Token fehlgeschlagen (App braucht Zustimmung „Office 365 SharePoint Online“ / Sites.FullControl.All).\n' +
                'Das ist im Browser oft nicht eingerichtet – bitte PowerShell unten oder Admin Center nutzen.\n' +
                String(e && e.message ? e.message : e);
            toast('Hub: bitte PowerShell oder Admin Center');
            return;
        }

        try {
            const hubJson = await G.registerHubSiteViaSpoRest(siteUrl, spoToken);
            $('fHubJson').textContent = JSON.stringify(hubJson, null, 2);
            $('fLog').textContent = 'Hub-Registrierung über SharePoint REST erfolgreich.\n' + siteUrl;
            toast('Als Hub-Website registriert.');
            try {
                if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                    window.ms365AppDataV2.patchSetup({
                        intranetSiteUrl: siteUrl,
                        schoolIntranetSiteUrl: siteUrl,
                        intranetHubAt: new Date().toISOString()
                    });
                }
            } catch {
                /* ignore */
            }
            if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                window.ms365ActionLog.append({
                    tool: 'sharepoint',
                    action: 'register-hub',
                    target: siteUrl,
                    summary: 'Hub-Website registriert'
                });
            }
        } catch (e) {
            const detail = String(e && e.message ? e.message : e);
            const corsLikely =
                /Failed to fetch|NetworkError|CORS|TypeError|Load failed|Access-Control/i.test(detail) ||
                detail === 'TypeError: Failed to fetch';
            revealHubFallback();
            $('fHubJson').textContent = detail;
            $('fLog').textContent =
                (corsLikely
                    ? 'Browser-REST an SharePoint ist blockiert (CORS) – das ist normal und kein Rechteproblem.\n'
                    : 'REST-Registrierung fehlgeschlagen.\n') +
                'Hub-Registrierung gibt es nicht über Microsoft Graph. Bitte eine der Alternativen unten:\n' +
                '• PowerShell-Skript kopieren und ausführen\n' +
                '• SharePoint Admin Center → Site → Hub → Als Hub-Website registrieren\n\n' +
                'Detail: ' +
                detail;
            toast('Hub: PowerShell oder Admin Center verwenden');
        }
    }

    $('btnHost').addEventListener('click', function () {
        detectHost().catch(function (e) {
            $('fLog').textContent = String(e && e.message ? e.message : e);
            toast('Fehler: ' + (e && e.message ? e.message : e));
        });
    });

    $('btnCreate').addEventListener('click', function () {
        createSite().catch(function (e) {
            $('fLog').textContent = String(e && e.message ? e.message : e);
            toast('Fehler: ' + (e && e.message ? e.message : e));
        });
    });

    $('btnHub').addEventListener('click', function () {
        registerHub().catch(function (e) {
            $('fLog').textContent += '\n' + String(e && e.message ? e.message : e);
        });
    });

    $('btnPsCopy').addEventListener('click', function () {
        const t = $('fPsHub').value;
        if (!t) return;
        navigator.clipboard.writeText(t).then(
            function () {
                toast('PowerShell kopiert.');
            },
            function () {
                toast('Kopieren fehlgeschlagen.');
            }
        );
    });

    function refreshPreview() {
        const u = buildWebUrl($('fHost').value, $('fSlug').value);
        $('fPreviewUrl').textContent = u || '–';
    }
    ['fHost', 'fSlug'].forEach(function (id) {
        const el = $(id);
        if (!el) return;
        el.addEventListener('input', refreshPreview);
    });
    refreshPreview();
    fillOwnerFromAccount();
    try {
        const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
        const saved =
            setup && setup.schoolIntranetSiteUrl
                ? String(setup.schoolIntranetSiteUrl).trim()
                : setup && setup.intranetSiteUrl
                  ? String(setup.intranetSiteUrl).trim()
                  : '';
        if (saved) {
            if ($('fLastSiteUrl')) $('fLastSiteUrl').value = saved;
            if ($('fManualUrl') && !$('fManualUrl').value) $('fManualUrl').value = saved;
            setPsScript(saved);
        }
    } catch {
        /* ignore */
    }
    window.addEventListener('ms365-auth-widget-ready', fillOwnerFromAccount);
    window.addEventListener('ms365-auth-state-changed', fillOwnerFromAccount);
})();
