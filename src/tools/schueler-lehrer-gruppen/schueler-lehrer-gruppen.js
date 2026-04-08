(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-schueler-lehrer-gruppen-v1';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/Group.ReadWrite.All'
    ];

    const PERSON_SELECT = 'id,displayName,mail,userPrincipalName';

    let msalMod = null;
    let pca = null;
    let slgCurrentStep = 1;

    /** @type {string | null} */
    let resolvedSchuelerId = null;
    /** @type {string | null} */
    let resolvedLehrerId = null;

    function toast(msg) {
        const el = document.getElementById('toast');
        if (el) {
            el.textContent = msg;
            el.classList.add('show');
            clearTimeout(toast._t);
            toast._t = setTimeout(function () {
                el.classList.remove('show');
            }, 3800);
        } else if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(msg);
        } else {
            window.alert(msg);
        }
    }

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normEmail(v) {
        return normStr(v).toLowerCase();
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
                typeof data === 'object' && data && data.error ? JSON.stringify(data.error) : text || String(res.status);
            throw new Error(method + ' ' + path + ': ' + msg);
        }
        return data || {};
    }

    function odataEscape(s) {
        return String(s).replace(/'/g, "''");
    }

    function guidLooksValid(s) {
        return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(
            String(s || '').trim()
        );
    }

    function sanitizeMailNickname(name) {
        let n = String(name || '')
            .replace(/[^0-9a-zA-Z]/g, '')
            .slice(0, 60);
        if (!n) n = 'group';
        return n.toLowerCase();
    }

    function isUnifiedGroup(g) {
        const gt = g && g.groupTypes;
        return Array.isArray(gt) && gt.indexOf('Unified') !== -1;
    }

    function userRef(userId) {
        return 'https://graph.microsoft.com/v1.0/users/' + userId;
    }

    function isDuplicateMemberError(e) {
        const m = String((e && e.message) || e || '');
        return (
            m.indexOf('added object references already exist') !== -1 ||
            m.indexOf('One or more added object references already exist') !== -1 ||
            m.indexOf('already exist') !== -1
        );
    }

    function loadTenantSettings() {
        if (typeof window.ms365TenantSettingsLoad !== 'function') {
            return null;
        }
        return window.ms365TenantSettingsLoad();
    }

    function getSchoolDomainNoAt() {
        const s = loadTenantSettings();
        const d = s && s.domain ? normStr(s.domain) : '';
        if (d) return d;
        if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
            return normStr(window.ms365GetSchoolDomainNoAt());
        }
        return '';
    }

    function collectStudentEmails(settings) {
        const out = [];
        const seen = new Set();
        const students = settings && Array.isArray(settings.students) ? settings.students : [];
        students.forEach(function (row) {
            const em = normEmail(row && row.email);
            if (!em || em.indexOf('@') === -1) return;
            if (seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    function collectTeacherEmails(settings) {
        const out = [];
        const seen = new Set();
        const teachers = settings && Array.isArray(settings.teachers) ? settings.teachers : [];
        teachers.forEach(function (row) {
            const em = normEmail(row && row.email);
            if (!em || em.indexOf('@') === -1) return;
            if (seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    function countStudentsWithAnyData(settings) {
        const students = settings && Array.isArray(settings.students) ? settings.students : [];
        let n = 0;
        students.forEach(function (row) {
            const klasse = normStr(row && (row.klasse || row.class));
            const name = normStr(row && row.name);
            const email = normEmail(row && row.email);
            if (klasse || name || email) n++;
        });
        return n;
    }

    function countTeachers(settings) {
        const teachers = settings && Array.isArray(settings.teachers) ? settings.teachers : [];
        return teachers.length;
    }

    function refreshStep1Ui() {
        const settings = loadTenantSettings();
        const domain = getSchoolDomainNoAt();
        const elDom = document.getElementById('slgDomainPreview');
        const elMS = document.getElementById('slgMailSchuelerPreview');
        const elML = document.getElementById('slgMailLehrerPreview');
        const st = document.getElementById('slgStatStudents');
        const stM = document.getElementById('slgStatStudentsMail');
        const tt = document.getElementById('slgStatTeachers');
        const ttM = document.getElementById('slgStatTeachersMail');
        const warn = document.getElementById('slgTenantWarn');

        if (elDom) elDom.textContent = domain || '(keine Domain in den Tenant‑Einstellungen)';
        if (elMS) elMS.textContent = domain ? 'schueler@' + domain : 'schueler@…';
        if (elML) elML.textContent = domain ? 'lehrer@' + domain : 'lehrer@…';

        const studEmails = collectStudentEmails(settings);
        const teachEmails = collectTeacherEmails(settings);
        if (st) st.textContent = String(countStudentsWithAnyData(settings));
        if (stM) stM.textContent = String(studEmails.length);
        if (tt) tt.textContent = String(countTeachers(settings));
        if (ttM) ttM.textContent = String(teachEmails.length);

        if (warn) {
            const lines = [];
            if (!domain) lines.push('Bitte in den Tenant‑Einstellungen eine Schul‑Domain eintragen (für die Adress‑Vorschau).');
            if (!studEmails.length) lines.push('Keine Schüler:innen mit E‑Mail in der Liste – Schritt 3 kann dort nichts übernehmen.');
            if (!teachEmails.length) lines.push('Keine Lehrer:innen mit E‑Mail in der Liste – Schritt 3 kann dort nichts übernehmen.');
            warn.style.display = lines.length ? 'block' : 'none';
            warn.innerHTML = lines.length ? '<strong>Hinweis:</strong> ' + lines.join(' ') : '';
        }
    }

    function slgStepNum(el) {
        const raw = el.getAttribute('data-slg-step');
        const n = parseFloat(String(raw || '').trim());
        return Number.isFinite(n) ? n : NaN;
    }

    function goToSlgStep(step) {
        slgCurrentStep = step;
        document.querySelectorAll('.slg-step-content').forEach(function (el) {
            el.classList.toggle('active', slgStepNum(el) === step);
        });
        document.querySelectorAll('.slg-steps .step').forEach(function (el) {
            const s = slgStepNum(el);
            el.classList.toggle('active', s === step);
            el.classList.toggle('completed', s < step);
        });
        if (typeof window.ms365ApplyStepProgress === 'function') {
            window.ms365ApplyStepProgress(document.querySelector('.slg-steps'), step, [1, 2, 3, 4]);
        }
        if (step === 4) {
            updateEntraLinks();
        }
    }

    function toggleModeBlocks() {
        const schuelerNew = document.querySelector('input[name="slgSchuelerMode"][value="new"]');
        const schuelerIsNew = schuelerNew && schuelerNew.checked;
        const b1 = document.getElementById('slgSchuelerNewBlock');
        const b2 = document.getElementById('slgSchuelerExistBlock');
        if (b1) b1.style.display = schuelerIsNew ? 'block' : 'none';
        if (b2) b2.style.display = schuelerIsNew ? 'none' : 'block';

        const lehrerNew = document.querySelector('input[name="slgLehrerMode"][value="new"]');
        const lehrerIsNew = lehrerNew && lehrerNew.checked;
        const l1 = document.getElementById('slgLehrerNewBlock');
        const l2 = document.getElementById('slgLehrerExistBlock');
        if (l1) l1.style.display = lehrerIsNew ? 'block' : 'none';
        if (l2) l2.style.display = lehrerIsNew ? 'none' : 'block';
    }

    function setSummary(kind, html, show) {
        const id = kind === 'schueler' ? 'slgSchuelerSummary' : 'slgLehrerSummary';
        const el = document.getElementById(id);
        if (!el) return;
        el.style.display = show ? 'block' : 'none';
        el.innerHTML = html || '';
    }

    async function fetchGroup(token, id) {
        const path =
            '/groups/' +
            encodeURIComponent(id) +
            '?$select=' +
            encodeURIComponent('id,displayName,mail,mailNickname,groupTypes,mailEnabled,securityEnabled');
        return graphJson('GET', path, token, undefined);
    }

    async function findGroupsByMailNickname(token, nickname) {
        const esc = odataEscape(nickname);
        const filter = "mailNickname eq '" + esc + "'";
        const path =
            '/groups?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent('id,displayName,mail,mailNickname,groupTypes') +
            '&$top=15';
        const data = await graphJson('GET', path, token, undefined);
        return data.value || [];
    }

    async function createUnifiedGroup(token, displayName, mailNickname, description) {
        const nick = sanitizeMailNickname(mailNickname);
        const body = {
            displayName: String(displayName).trim(),
            description: description || 'MS365-Schulverwaltung – Schüler:innen/Lehrer:innen',
            mailNickname: nick,
            mailEnabled: true,
            securityEnabled: false,
            groupTypes: ['Unified'],
            visibility: 'Private'
        };
        const group = await graphJson('POST', '/groups', token, body);
        const gid = group.id;
        await sleep(1500);
        try {
            const me = await graphJson('GET', '/me', token, undefined);
            const meId = me && me.id;
            if (meId) {
                try {
                    await graphJson('POST', '/groups/' + gid + '/owners/$ref', token, {
                        '@odata.id': userRef(meId)
                    });
                } catch (e) {
                    if (!isDuplicateMemberError(e)) throw e;
                }
                try {
                    await graphJson('POST', '/groups/' + gid + '/members/$ref', token, {
                        '@odata.id': userRef(meId)
                    });
                } catch (e) {
                    if (!isDuplicateMemberError(e)) throw e;
                }
            }
        } catch (e) {
            /* Besitzer optional */
        }
        return group;
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

    async function graphAddMember(token, groupId, userId) {
        await graphJson(
            'POST',
            '/groups/' + encodeURIComponent(groupId) + '/members/$ref',
            token,
            {
                '@odata.id': userRef(userId)
            }
        );
    }

    function appendSyncLog(msg, kind) {
        const el = document.getElementById('slgSyncLog');
        if (!el) return;
        const line = document.createElement('div');
        line.textContent = new Date().toLocaleTimeString() + '  ' + msg;
        if (kind === 'err') line.style.color = '#b00020';
        else if (kind === 'ok') line.style.color = '#0d8050';
        else if (kind === 'warn') line.style.color = '#856404';
        el.appendChild(line);
        el.scrollTop = el.scrollHeight;
    }

    function clearSyncLog() {
        const el = document.getElementById('slgSyncLog');
        if (el) el.replaceChildren();
    }

    function updateEntraLinks() {
        const a1 = document.getElementById('slgLinkSchuelerEntra');
        const a2 = document.getElementById('slgLinkLehrerEntra');
        const sep = document.getElementById('slgLinkSep');
        const base = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Members/groupId/';
        if (a1 && resolvedSchuelerId) {
            a1.href = base + encodeURIComponent(resolvedSchuelerId);
            a1.style.display = 'inline';
        } else if (a1) {
            a1.style.display = 'none';
        }
        if (a2 && resolvedLehrerId) {
            a2.href = base + encodeURIComponent(resolvedLehrerId);
            a2.style.display = 'inline';
        } else if (a2) {
            a2.style.display = 'none';
        }
        if (sep) {
            sep.style.display = resolvedSchuelerId && resolvedLehrerId ? 'inline' : 'none';
        }
    }

    function persistResolvedIds(kind, group) {
        if (kind === 'schueler') {
            resolvedSchuelerId = group && group.id ? String(group.id) : null;
        } else {
            resolvedLehrerId = group && group.id ? String(group.id) : null;
        }
    }

    function formatGroupSummary(g) {
        if (!g || !g.id) return '';
        const unified = isUnifiedGroup(g) ? 'Microsoft 365‑Gruppe (Unified)' : 'Keine Unified‑Gruppe';
        const mail = normStr(g.mail) || '–';
        const nick = normStr(g.mailNickname) || '–';
        return (
            '<strong>OK:</strong> ' +
            normStr(g.displayName) +
            '<br>Object‑ID: <code>' +
            g.id +
            '</code><br>Mail‑Nickname: <code>' +
            nick +
            '</code> · SMTP: ' +
            mail +
            '<br><span style="color:#084298;">' +
            unified +
            '</span>'
        );
    }

    async function handleCreateSchueler() {
        const dn = document.getElementById('slgSchuelerDisplayName');
        const nn = document.getElementById('slgSchuelerMailNick');
        const displayName = dn ? dn.value : 'Schüler:innen';
        const mailNick = nn ? nn.value : 'schueler';
        try {
            const token = await getGraphToken();
            const g = await createUnifiedGroup(
                token,
                displayName,
                mailNick,
                'Alle Schüler:innen (MS365-Schulverwaltung / Tenant‑Liste)'
            );
            persistResolvedIds('schueler', g);
            setSummary('schueler', formatGroupSummary(g), true);
            toast('Schüler:innen‑Gruppe angelegt.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleCreateLehrer() {
        const dn = document.getElementById('slgLehrerDisplayName');
        const nn = document.getElementById('slgLehrerMailNick');
        const displayName = dn ? dn.value : 'Lehrer:innen';
        const mailNick = nn ? nn.value : 'lehrer';
        try {
            const token = await getGraphToken();
            const g = await createUnifiedGroup(
                token,
                displayName,
                mailNick,
                'Alle Lehrer:innen (MS365-Schulverwaltung / Tenant‑Liste)'
            );
            persistResolvedIds('lehrer', g);
            setSummary('lehrer', formatGroupSummary(g), true);
            toast('Lehrer:innen‑Gruppe angelegt.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleResolveSchueler() {
        const inp = document.getElementById('slgSchuelerGroupId');
        const id = inp ? normStr(inp.value) : '';
        if (!guidLooksValid(id)) {
            toast('Bitte eine gültige Object‑ID (GUID) eintragen.');
            return;
        }
        try {
            const token = await getGraphToken();
            const g = await fetchGroup(token, id);
            if (!isUnifiedGroup(g)) {
                setSummary(
                    'schueler',
                    '<strong>Warnung:</strong> Diese Gruppe ist keine Microsoft 365‑Gruppe (Unified). Bitte eine passende Gruppe wählen.',
                    true
                );
                persistResolvedIds('schueler', g);
                return;
            }
            persistResolvedIds('schueler', g);
            setSummary('schueler', formatGroupSummary(g), true);
            toast('Schüler:innen‑Gruppe geladen.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleResolveLehrer() {
        const inp = document.getElementById('slgLehrerGroupId');
        const id = inp ? normStr(inp.value) : '';
        if (!guidLooksValid(id)) {
            toast('Bitte eine gültige Object‑ID (GUID) eintragen.');
            return;
        }
        try {
            const token = await getGraphToken();
            const g = await fetchGroup(token, id);
            if (!isUnifiedGroup(g)) {
                setSummary(
                    'lehrer',
                    '<strong>Warnung:</strong> Diese Gruppe ist keine Microsoft 365‑Gruppe (Unified). Bitte eine passende Gruppe wählen.',
                    true
                );
                persistResolvedIds('lehrer', g);
                return;
            }
            persistResolvedIds('lehrer', g);
            setSummary('lehrer', formatGroupSummary(g), true);
            toast('Lehrer:innen‑Gruppe geladen.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleFindSchuelerNick() {
        try {
            const token = await getGraphToken();
            const nn = document.getElementById('slgSchuelerMailNick');
            const nick = sanitizeMailNickname(nn ? nn.value : 'schueler') || 'schueler';
            const list = await findGroupsByMailNickname(token, nick);
            if (!list.length) {
                toast('Keine Gruppe mit Mail‑Nickname „' + nick + '“ gefunden.');
                return;
            }
            const g = list[0];
            const inp = document.getElementById('slgSchuelerGroupId');
            if (inp) inp.value = g.id;
            persistResolvedIds('schueler', g);
            if (!isUnifiedGroup(g)) {
                setSummary(
                    'schueler',
                    '<strong>Warnung:</strong> Gefundene Gruppe ist keine Unified‑Gruppe. ' + formatGroupSummary(g),
                    true
                );
                return;
            }
            setSummary('schueler', formatGroupSummary(g), true);
            toast(list.length > 1 ? 'Mehrere Treffer – erste Gruppe übernommen.' : 'Gruppe gefunden.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleFindLehrerNick() {
        try {
            const token = await getGraphToken();
            const nn = document.getElementById('slgLehrerMailNick');
            const nick = sanitizeMailNickname(nn ? nn.value : 'lehrer') || 'lehrer';
            const list = await findGroupsByMailNickname(token, nick);
            if (!list.length) {
                toast('Keine Gruppe mit Mail‑Nickname „' + nick + '“ gefunden.');
                return;
            }
            const g = list[0];
            const inp = document.getElementById('slgLehrerGroupId');
            if (inp) inp.value = g.id;
            persistResolvedIds('lehrer', g);
            if (!isUnifiedGroup(g)) {
                setSummary(
                    'lehrer',
                    '<strong>Warnung:</strong> Gefundene Gruppe ist keine Unified‑Gruppe. ' + formatGroupSummary(g),
                    true
                );
                return;
            }
            setSummary('lehrer', formatGroupSummary(g), true);
            toast(list.length > 1 ? 'Mehrere Treffer – erste Gruppe übernommen.' : 'Gruppe gefunden.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function syncEmailsToGroup(token, groupId, emails, label) {
        let ok = 0;
        let skip = 0;
        let fail = 0;
        for (let i = 0; i < emails.length; i++) {
            const em = emails[i];
            try {
                const u = await resolveUserByEmail(token, em);
                if (!u || !u.id) {
                    appendSyncLog(label + ': Kein Benutzer für ' + em, 'warn');
                    fail++;
                    continue;
                }
                try {
                    await graphAddMember(token, groupId, u.id);
                    ok++;
                    appendSyncLog(label + ': ' + em + ' → Mitglied', 'ok');
                } catch (e) {
                    if (isDuplicateMemberError(e)) {
                        skip++;
                        appendSyncLog(label + ': ' + em + ' (war schon Mitglied)', 'warn');
                    } else {
                        fail++;
                        appendSyncLog(label + ': ' + em + ' — ' + (e.message || e), 'err');
                    }
                }
            } catch (e) {
                fail++;
                appendSyncLog(label + ': ' + em + ' — ' + (e.message || e), 'err');
            }
            if ((i + 1) % 8 === 0) {
                await sleep(120);
            }
        }
        return { ok: ok, skip: skip, fail: fail };
    }

    async function handleSyncStudents() {
        const settings = loadTenantSettings();
        const emails = collectStudentEmails(settings);
        if (!emails.length) {
            toast('Keine Schüler:innen‑E‑Mails in den Tenant‑Einstellungen.');
            return;
        }
        if (!resolvedSchuelerId) {
            toast('Zuerst in Schritt 2 die Schüler:innen‑Gruppe anlegen oder auswählen.');
            return;
        }
        clearSyncLog();
        appendSyncLog('Start: Schüler:innen (' + emails.length + ' Adressen) …', '');
        try {
            let token = await getGraphToken();
            const r = await syncEmailsToGroup(token, resolvedSchuelerId, emails, 'Schüler');
            appendSyncLog(
                'Fertig Schüler:innen: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.',
                'ok'
            );
            toast('Synchronisation Schüler:innen abgeschlossen.');
        } catch (e) {
            appendSyncLog('Abbruch: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleSyncTeachers() {
        const settings = loadTenantSettings();
        const emails = collectTeacherEmails(settings);
        if (!emails.length) {
            toast('Keine Lehrer:innen‑E‑Mails in den Tenant‑Einstellungen.');
            return;
        }
        if (!resolvedLehrerId) {
            toast('Zuerst in Schritt 2 die Lehrer:innen‑Gruppe anlegen oder auswählen.');
            return;
        }
        clearSyncLog();
        appendSyncLog('Start: Lehrer:innen (' + emails.length + ' Adressen) …', '');
        try {
            const token = await getGraphToken();
            const r = await syncEmailsToGroup(token, resolvedLehrerId, emails, 'Lehrer');
            appendSyncLog(
                'Fertig Lehrer:innen: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.',
                'ok'
            );
            toast('Synchronisation Lehrer:innen abgeschlossen.');
        } catch (e) {
            appendSyncLog('Abbruch: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        }
    }

    function parseManualLines(text) {
        const raw = String(text || '').split(/\r?\n/);
        const out = [];
        const seen = new Set();
        raw.forEach(function (line) {
            const p = normStr(line);
            if (!p || p.indexOf('@') === -1) return;
            const em = normEmail(p);
            if (seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    async function handleManualAdd() {
        const sel = document.getElementById('slgManualTarget');
        const ta = document.getElementById('slgManualLines');
        const outEl = document.getElementById('slgManualResult');
        const kind = sel && sel.value === 'lehrer' ? 'lehrer' : 'schueler';
        const gid = kind === 'lehrer' ? resolvedLehrerId : resolvedSchuelerId;
        if (!gid) {
            toast('Keine ' + (kind === 'lehrer' ? 'Lehrer:innen' : 'Schüler:innen') + '‑Gruppe – bitte Schritt 2.');
            return;
        }
        const emails = parseManualLines(ta ? ta.value : '');
        if (!emails.length) {
            toast('Bitte mindestens eine gültige E‑Mail‑Zeile eintragen.');
            return;
        }
        try {
            const token = await getGraphToken();
            let ok = 0;
            let skip = 0;
            let fail = 0;
            const lines = [];
            for (let i = 0; i < emails.length; i++) {
                const em = emails[i];
                try {
                    const u = await resolveUserByEmail(token, em);
                    if (!u || !u.id) {
                        fail++;
                        lines.push(em + ' → nicht gefunden');
                        continue;
                    }
                    try {
                        await graphAddMember(token, gid, u.id);
                        ok++;
                        lines.push(em + ' → hinzugefügt');
                    } catch (e) {
                        if (isDuplicateMemberError(e)) {
                            skip++;
                            lines.push(em + ' → war schon Mitglied');
                        } else {
                            fail++;
                            lines.push(em + ' → Fehler: ' + (e.message || e));
                        }
                    }
                } catch (e) {
                    fail++;
                    lines.push(em + ' → ' + (e.message || e));
                }
            }
            if (outEl) {
                outEl.style.display = 'block';
                outEl.innerHTML =
                    '<strong>Ergebnis:</strong> neu ' +
                    ok +
                    ', übersprungen ' +
                    skip +
                    ', Fehler ' +
                    fail +
                    '.<br>' +
                    lines.map(function (x) {
                        return escapeHtml(x);
                    }).join('<br>');
            }
            toast('Manuelle Zuordnung abgeschlossen.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function buildStateObject() {
        return {
            kind: 'ms365-schueler-lehrer-gruppen-v1',
            savedAt: new Date().toISOString(),
            step: slgCurrentStep,
            resolvedSchuelerId: resolvedSchuelerId,
            resolvedLehrerId: resolvedLehrerId,
            slgSchuelerDisplayName: document.getElementById('slgSchuelerDisplayName')
                ? document.getElementById('slgSchuelerDisplayName').value
                : '',
            slgLehrerDisplayName: document.getElementById('slgLehrerDisplayName')
                ? document.getElementById('slgLehrerDisplayName').value
                : '',
            slgSchuelerMailNick: document.getElementById('slgSchuelerMailNick')
                ? document.getElementById('slgSchuelerMailNick').value
                : 'schueler',
            slgLehrerMailNick: document.getElementById('slgLehrerMailNick')
                ? document.getElementById('slgLehrerMailNick').value
                : 'lehrer',
            slgSchuelerGroupId: document.getElementById('slgSchuelerGroupId')
                ? document.getElementById('slgSchuelerGroupId').value
                : '',
            slgLehrerGroupId: document.getElementById('slgLehrerGroupId')
                ? document.getElementById('slgLehrerGroupId').value
                : '',
            slgSchuelerMode: document.querySelector('input[name="slgSchuelerMode"]:checked')
                ? document.querySelector('input[name="slgSchuelerMode"]:checked').value
                : 'new',
            slgLehrerMode: document.querySelector('input[name="slgLehrerMode"]:checked')
                ? document.querySelector('input[name="slgLehrerMode"]:checked').value
                : 'new',
            slgManualLines: document.getElementById('slgManualLines') ? document.getElementById('slgManualLines').value : ''
        };
    }

    function applyStateObject(o) {
        if (!o || typeof o !== 'object') return;
        if (o.step !== undefined) {
            const s = parseInt(String(o.step), 10);
            if (s >= 1 && s <= 4) goToSlgStep(s);
            else goToSlgStep(1);
        } else {
            goToSlgStep(1);
        }
        resolvedSchuelerId = o.resolvedSchuelerId ? String(o.resolvedSchuelerId) : null;
        resolvedLehrerId = o.resolvedLehrerId ? String(o.resolvedLehrerId) : null;

        function setVal(id, v) {
            const el = document.getElementById(id);
            if (el && v !== undefined) el.value = String(v);
        }
        setVal('slgSchuelerDisplayName', o.slgSchuelerDisplayName);
        setVal('slgLehrerDisplayName', o.slgLehrerDisplayName);
        setVal('slgSchuelerMailNick', o.slgSchuelerMailNick);
        setVal('slgLehrerMailNick', o.slgLehrerMailNick);
        setVal('slgSchuelerGroupId', o.slgSchuelerGroupId);
        setVal('slgLehrerGroupId', o.slgLehrerGroupId);
        setVal('slgManualLines', o.slgManualLines);

        if (o.slgSchuelerMode) {
            const r = document.querySelector('input[name="slgSchuelerMode"][value="' + o.slgSchuelerMode + '"]');
            if (r) r.checked = true;
        }
        if (o.slgLehrerMode) {
            const r = document.querySelector('input[name="slgLehrerMode"][value="' + o.slgLehrerMode + '"]');
            if (r) r.checked = true;
        }
        toggleModeBlocks();
        updateEntraLinks();
    }

    function saveState() {
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(buildStateObject()));
            toast('Zwischenstand gespeichert.');
        } catch (e) {
            toast('Speichern fehlgeschlagen: ' + (e.message || e));
        }
    }

    function loadState() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) {
                toast('Kein gespeicherter Stand.');
                return;
            }
            const o = JSON.parse(raw);
            applyStateObject(o);
            toast('Stand geladen.');
        } catch (e) {
            toast('Laden fehlgeschlagen: ' + (e.message || e));
        }
    }

    function clearStorage() {
        try {
            localStorage.removeItem(STORAGE_KEY);
            resolvedSchuelerId = null;
            resolvedLehrerId = null;
            setSummary('schueler', '', false);
            setSummary('lehrer', '', false);
            toast('Lokaler Speicher gelöscht.');
        } catch (e) {
            toast('Löschen fehlgeschlagen: ' + (e.message || e));
        }
    }

    function wire() {
        document.getElementById('slgBtnNext1') &&
            document.getElementById('slgBtnNext1').addEventListener('click', function () {
                goToSlgStep(2);
            });
        document.getElementById('slgBtnBack2') &&
            document.getElementById('slgBtnBack2').addEventListener('click', function () {
                goToSlgStep(1);
            });
        document.getElementById('slgBtnNext2') &&
            document.getElementById('slgBtnNext2').addEventListener('click', function () {
                goToSlgStep(3);
            });
        document.getElementById('slgBtnBack3') &&
            document.getElementById('slgBtnBack3').addEventListener('click', function () {
                goToSlgStep(2);
            });
        document.getElementById('slgBtnNext3') &&
            document.getElementById('slgBtnNext3').addEventListener('click', function () {
                goToSlgStep(4);
            });
        document.getElementById('slgBtnBack4') &&
            document.getElementById('slgBtnBack4').addEventListener('click', function () {
                goToSlgStep(3);
            });

        document.getElementById('slgBtnLogin') &&
            document.getElementById('slgBtnLogin').addEventListener('click', async function () {
                try {
                    await getGraphToken();
                    toast('Angemeldet.');
                } catch (e) {
                    toast('Anmeldung: ' + (e.message || e));
                }
            });

        document.getElementById('slgBtnCreateSchueler') &&
            document.getElementById('slgBtnCreateSchueler').addEventListener('click', function () {
                handleCreateSchueler();
            });
        document.getElementById('slgBtnCreateLehrer') &&
            document.getElementById('slgBtnCreateLehrer').addEventListener('click', function () {
                handleCreateLehrer();
            });
        document.getElementById('slgBtnResolveSchueler') &&
            document.getElementById('slgBtnResolveSchueler').addEventListener('click', function () {
                handleResolveSchueler();
            });
        document.getElementById('slgBtnResolveLehrer') &&
            document.getElementById('slgBtnResolveLehrer').addEventListener('click', function () {
                handleResolveLehrer();
            });
        document.getElementById('slgBtnFindSchuelerNick') &&
            document.getElementById('slgBtnFindSchuelerNick').addEventListener('click', function () {
                handleFindSchuelerNick();
            });
        document.getElementById('slgBtnFindLehrerNick') &&
            document.getElementById('slgBtnFindLehrerNick').addEventListener('click', function () {
                handleFindLehrerNick();
            });

        document.querySelectorAll('input[name="slgSchuelerMode"]').forEach(function (r) {
            r.addEventListener('change', toggleModeBlocks);
        });
        document.querySelectorAll('input[name="slgLehrerMode"]').forEach(function (r) {
            r.addEventListener('change', toggleModeBlocks);
        });

        document.getElementById('slgBtnSyncStudents') &&
            document.getElementById('slgBtnSyncStudents').addEventListener('click', function () {
                handleSyncStudents();
            });
        document.getElementById('slgBtnSyncTeachers') &&
            document.getElementById('slgBtnSyncTeachers').addEventListener('click', function () {
                handleSyncTeachers();
            });

        document.getElementById('slgBtnManualAdd') &&
            document.getElementById('slgBtnManualAdd').addEventListener('click', function () {
                handleManualAdd();
            });

        document.getElementById('slgBtnSaveState') &&
            document.getElementById('slgBtnSaveState').addEventListener('click', saveState);
        document.getElementById('slgBtnLoadState') &&
            document.getElementById('slgBtnLoadState').addEventListener('click', loadState);
        document.getElementById('slgBtnClearStorage') &&
            document.getElementById('slgBtnClearStorage').addEventListener('click', clearStorage);
    }

    function init() {
        wire();
        refreshStep1Ui();
        toggleModeBlocks();
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (raw) {
                const o = JSON.parse(raw);
                applyStateObject(o);
            } else {
                goToSlgStep(1);
            }
        } catch {
            goToSlgStep(1);
        }
        updateEntraLinks();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
