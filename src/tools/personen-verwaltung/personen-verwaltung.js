(function () {
    'use strict';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/User.ReadWrite.All',
        'https://graph.microsoft.com/Group.ReadWrite.All',
        'https://graph.microsoft.com/Organization.Read.All'
    ];

    const USER_LIST_SELECT =
        'id,displayName,givenName,surname,mail,mailNickname,userPrincipalName,jobTitle,department,' +
        'officeLocation,mobilePhone,businessPhones,companyName,preferredLanguage,accountEnabled,' +
        'streetAddress,city,postalCode,country,createdDateTime,userType,assignedLicenses,usageLocation';

    const USER_REFRESH_SELECT = USER_LIST_SELECT;

    const GROUP_MEMBEROF_SELECT = 'id,displayName,mail,mailNickname,groupTypes,securityEnabled,mailEnabled';

    let msalMod = null;
    let pca = null;
    /** @type {Record<string, any>[]} */
    let loadedUsers = [];
    /** @type {string | null} */
    let selectedUserId = null;
    let pendingTabAfterSelect = '';
    /** @type {'profil' | 'lizenzen' | 'gruppen'} */
    let activeTab = 'profil';
    /** @type {Record<string, any>[] | null} */
    let cachedGroupsForSelection = null;
    /** @type {boolean} */
    let profileEditMode = false;
    /** @type {Record<string, any>[]} */
    let subscribedSkus = [];
    /** @type {boolean} */
    let subscribedSkusOk = false;
    /** @type {boolean} */
    let licenseBusy = false;
    /** @type {boolean} */
    let groupBusy = false;

    const SESSION_CACHE_KEY = 'ms365-pv-users-cache-v1';
    const SESSION_CACHE_MAX_AGE_MS = 30 * 60 * 1000; // 30 Minuten

    function saveUsersToSession() {
        try {
            sessionStorage.setItem(
                SESSION_CACHE_KEY,
                JSON.stringify({
                    savedAt: Date.now(),
                    users: loadedUsers,
                    skus: subscribedSkus,
                    skusOk: subscribedSkusOk
                })
            );
        } catch {
            // sessionStorage voll oder nicht verfügbar – ignorieren
        }
    }

    function loadUsersFromSession() {
        try {
            const raw = sessionStorage.getItem(SESSION_CACHE_KEY);
            if (!raw) return false;
            const obj = JSON.parse(raw);
            if (!obj || !Array.isArray(obj.users) || !obj.users.length) return false;
            const age = Date.now() - (obj.savedAt || 0);
            if (age > SESSION_CACHE_MAX_AGE_MS) return false;
            loadedUsers = obj.users;
            subscribedSkus = Array.isArray(obj.skus) ? obj.skus : [];
            subscribedSkusOk = !!obj.skusOk;
            return obj.savedAt || null;
        } catch {
            return false;
        }
    }

    function clearUsersFromSession() {
        try { sessionStorage.removeItem(SESSION_CACHE_KEY); } catch { /* ignore */ }
    }

    function showCacheBanner(savedAt) {
        const banner = document.getElementById('pvCacheBanner');
        if (!banner) return;
        const d = new Date(savedAt);
        const time = d.toLocaleTimeString('de-AT', { hour: '2-digit', minute: '2-digit' });
        banner.style.display = '';
        banner.textContent = 'Daten aus dem Sitzungs-Cache (eingelesen um ' + time + '). ';
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn small-btn';
        btn.style.marginLeft = '8px';
        btn.textContent = 'Jetzt neu einlesen';
        btn.addEventListener('click', function () {
            clearUsersFromSession();
            loadUsers();
        });
        banner.appendChild(btn);
    }

    function hideCacheBanner() {
        const banner = document.getElementById('pvCacheBanner');
        if (banner) banner.style.display = 'none';
    }

    function dlgConfirm(msg, opts) {
        if (typeof window.ms365AppDialogConfirm === 'function') {
            return window.ms365AppDialogConfirm(msg, opts);
        }
        return Promise.resolve(window.confirm(msg));
    }

    function dlgPrompt(msg, def, opts) {
        if (typeof window.ms365AppDialogPrompt === 'function') {
            return window.ms365AppDialogPrompt(msg, def, opts);
        }
        return Promise.resolve(window.prompt(msg, def));
    }

    function graphErrorFriendly(e) {
        const raw = String(e && e.message ? e.message : e);
        const idx = raw.indexOf('{');
        if (idx !== -1) {
            try {
                const obj = JSON.parse(raw.slice(idx));
                const inner = obj.error || obj;
                if (inner && inner.message) return String(inner.message);
            } catch {
                // ignore
            }
        }
        return raw;
    }

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

    async function graphRequest(method, pathOrUrl, token, body, extraHeaders) {
        const url =
            pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
        let attempt = 0;
        while (true) {
            const headers = { Authorization: 'Bearer ' + token };
            if (extraHeaders && typeof extraHeaders === 'object') {
                Object.assign(headers, extraHeaders);
            }
            if (body !== undefined && method !== 'GET' && method !== 'DELETE') {
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

    async function graphJson(method, pathOrUrl, token, body, extraHeaders) {
        const res = await graphRequest(method, pathOrUrl, token, body, extraHeaders);
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

    async function graphDelete(path, token) {
        const res = await graphRequest('DELETE', path, token, undefined);
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
            throw new Error('DELETE ' + path + ': ' + msg);
        }
    }

    function appendLog(msg, kind) {
        const el = document.getElementById('pvLog');
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
        const el = document.getElementById('pvLog');
        if (el) el.replaceChildren();
    }

    async function fetchAllPages(token, initialPath, onProgress) {
        const out = [];
        let next = initialPath;
        let page = 0;
        while (next) {
            page++;
            const data = await graphJson('GET', next, token, undefined);
            const vals = data.value;
            if (Array.isArray(vals)) {
                for (let i = 0; i < vals.length; i++) out.push(vals[i]);
            }
            next = data['@odata.nextLink'] || null;
            if (onProgress) onProgress(out.length, page, !!next);
        }
        return out;
    }

    function norm(s) {
        return String(s || '').trim().toLowerCase();
    }

    function compareStrings(a, b) {
        return String(a || '').localeCompare(String(b || ''), 'de', { sensitivity: 'base' });
    }

    function readSortFromSelect() {
        const sel = document.getElementById('pvSortKey');
        const raw = sel && sel.value ? String(sel.value) : 'displayName:asc';
        const parts = raw.split(':');
        const key = parts[0] || 'displayName';
        const dir = parts[1] === 'desc' ? 'desc' : 'asc';
        return { key: key, dir: dir };
    }

    function formatPhones(u) {
        const m = u && u.mobilePhone ? String(u.mobilePhone).trim() : '';
        const bp = u && Array.isArray(u.businessPhones) ? u.businessPhones.filter(Boolean).join(', ') : '';
        if (m && bp) return m + ' · ' + bp;
        return m || bp || '';
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

    function groupTypeLabel(g) {
        if (!g || typeof g !== 'object') return '–';
        const types = g.groupTypes;
        if (Array.isArray(types) && types.indexOf('Unified') !== -1) return 'Microsoft 365 (Unified)';
        if (g.securityEnabled && !g.mailEnabled) return 'Sicherheitsgruppe';
        if (g.mailEnabled && !g.securityEnabled) return 'Verteilerliste';
        if (g.securityEnabled && g.mailEnabled) return 'Mail-aktivierte Sicherheitsgruppe';
        return 'Gruppe';
    }

    function userTypeLabel(ut) {
        const t = String(ut || '').toLowerCase();
        if (t === 'guest') return 'Gast';
        if (t === 'member') return 'Mitglied';
        return ut ? String(ut) : '–';
    }

    function getSelectedUser() {
        if (!selectedUserId) return null;
        return loadedUsers.find(function (x) {
            return x.id === selectedUserId;
        }) || null;
    }

    function updateDetailActionButtons() {
        const save = document.getElementById('pvBtnSave');
        const saveBottom = document.getElementById('pvBtnSaveBottom');
        const cancel = document.getElementById('pvBtnCancelEdit');
        const del = document.getElementById('pvBtnDelete');
        const hasSel = !!selectedUserId;
        if (save) {
            save.style.display = hasSel ? '' : 'none';
            save.disabled = !hasSel;
        }
        if (saveBottom) saveBottom.disabled = !hasSel;
        if (cancel) {
            cancel.style.display = hasSel ? '' : 'none';
            cancel.disabled = !hasSel;
        }
        if (del) {
            del.style.display = hasSel ? '' : 'none';
            del.disabled = !hasSel;
        }
    }

    function Lic() {
        return window.ms365GraphLicenses || null;
    }

    function userLicenseSummary(u) {
        const api = Lic();
        if (!api || typeof api.summarizeUserLicenses !== 'function') return null;
        return api.summarizeUserLicenses(u);
    }

    function getVisibleRows() {
        const filterInp = document.getElementById('pvFilterText');
        const q = filterInp && filterInp.value ? norm(filterInp.value) : '';

        const typeSel = document.getElementById('pvFilterUserType');
        const typeVal = typeSel && typeSel.value ? String(typeSel.value) : '';

        const accSel = document.getElementById('pvFilterAccount');
        const accVal = accSel && accSel.value !== '' ? String(accSel.value) : '';

        const depSel = document.getElementById('pvFilterDepartment');
        const depVal = depSel && depSel.value ? String(depSel.value) : '';

        const licSel = document.getElementById('pvFilterLicense');
        const licVal = licSel && licSel.value ? String(licSel.value) : '';

        let rows = loadedUsers.slice();

        if (typeVal) {
            rows = rows.filter(function (u) {
                return String(u.userType || '') === typeVal;
            });
        }

        if (accVal === '1') {
            rows = rows.filter(function (u) {
                return u.accountEnabled === true;
            });
        } else if (accVal === '0') {
            rows = rows.filter(function (u) {
                return u.accountEnabled === false;
            });
        }

        if (depVal) {
            rows = rows.filter(function (u) {
                return String(u.department || '').trim() === depVal;
            });
        }

        if (licVal) {
            const api = Lic();
            if (api && typeof api.userMatchesLicenseFilter === 'function') {
                rows = rows.filter(function (u) {
                    return api.userMatchesLicenseFilter(u, licVal);
                });
            }
        }

        if (q) {
            rows = rows.filter(function (u) {
                const blob = [
                    u.displayName,
                    u.givenName,
                    u.surname,
                    u.mail,
                    u.userPrincipalName,
                    u.department,
                    u.jobTitle,
                    u.id,
                    u.officeLocation,
                    u.companyName
                ]
                    .map(function (x) {
                        return norm(x);
                    })
                    .join(' ');
                const sum = userLicenseSummary(u);
                const lic = sum && sum.primaryLabel ? norm(sum.primaryLabel) : '';
                return blob.indexOf(q) !== -1 || (lic && lic.indexOf(q) !== -1);
            });
        }

        const sortState = readSortFromSelect();
        const key = sortState.key || 'displayName';
        const dir = sortState.dir === 'desc' ? -1 : 1;

        rows.sort(function (ua, ub) {
            return compareStrings(ua[key] || '', ub[key] || '') * dir;
        });

        return rows;
    }

    function refreshDepartmentFilter() {
        const sel = document.getElementById('pvFilterDepartment');
        if (!sel) return;
        const current = sel.value;
        const set = new Set();
        for (let i = 0; i < loadedUsers.length; i++) {
            const d = loadedUsers[i].department;
            if (d && String(d).trim()) set.add(String(d).trim());
        }
        const list = Array.from(set).sort(function (a, b) {
            return compareStrings(a, b);
        });
        sel.replaceChildren();
        const o0 = document.createElement('option');
        o0.value = '';
        o0.textContent = '(alle)';
        sel.appendChild(o0);
        for (let j = 0; j < list.length; j++) {
            const o = document.createElement('option');
            o.value = list[j];
            o.textContent = list[j];
            sel.appendChild(o);
        }
        if (current && set.has(current)) sel.value = current;
    }

    function refreshLicenseFilter() {
        const sel = document.getElementById('pvFilterLicense');
        if (!sel) return;
        const api = Lic();
        const current = sel.value;
        sel.replaceChildren();
        const opts =
            api && typeof api.buildLicenseFilterOptions === 'function'
                ? api.buildLicenseFilterOptions(loadedUsers)
                : [{ value: '', label: '(alle Lizenzen)' }];
        for (let i = 0; i < opts.length; i++) {
            const o = document.createElement('option');
            o.value = opts[i].value;
            o.textContent = opts[i].label;
            sel.appendChild(o);
        }
        const values = {};
        for (let j = 0; j < opts.length; j++) values[opts[j].value] = true;
        if (current && values[current]) sel.value = current;
    }

    function updateStatsPanel() {
        const total = loadedUsers.length;
        let members = 0;
        let guests = 0;
        let active = 0;
        for (let i = 0; i < loadedUsers.length; i++) {
            const u = loadedUsers[i];
            if (String(u.userType || '').toLowerCase() === 'guest') guests++;
            else members++;
            if (u.accountEnabled === true) active++;
        }
        const el = function (id, val) {
            const n = document.getElementById(id);
            if (n) n.textContent = val;
        };
        el('pvStatTotal', total ? String(total) : '–');
        el('pvStatMember', total ? String(members) : '–');
        el('pvStatGuest', total ? String(guests) : '–');
        el('pvStatActive', total ? String(active) : '–');
    }

    function updateProgressLine() {
        const progress = document.getElementById('pvProgress');
        if (!progress) return;
        if (!loadedUsers.length) {
            progress.textContent = '';
            return;
        }
        const visible = getVisibleRows();
        const base = 'Geladen: ' + loadedUsers.length + ' Person(en).';
        if (visible.length !== loadedUsers.length) {
            progress.textContent = base + ' Angezeigt: ' + visible.length + ' Treffer.';
        } else {
            progress.textContent = base;
        }
    }

    function renderUserTree() {
        const tree = document.getElementById('pvTree');
        if (!tree) return;
        tree.replaceChildren();
        const rows = getVisibleRows();

        if (!rows.length) {
            const li = document.createElement('li');
            const p = document.createElement('p');
            p.className = 'muted';
            p.style.margin = '0';
            p.style.padding = '14px 12px';
            p.textContent = loadedUsers.length ? 'Keine Treffer für die Filter.' : 'Noch keine Daten – „Personen einlesen“ wählen.';
            li.appendChild(p);
            tree.appendChild(li);
            updateProgressLine();
            return;
        }

        for (let i = 0; i < rows.length; i++) {
            const u = rows[i];
            const li = document.createElement('li');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'pv-tree-row';
            btn.dataset.pvSelectUser = u.id || '';
            btn.setAttribute('aria-current', selectedUserId && u.id === selectedUserId ? 'true' : 'false');

            const isGuest = String(u.userType || '').toLowerCase() === 'guest';
            const iconWrap = document.createElement('span');
            iconWrap.className = 'pv-tree-icon';
            const icon = document.createElement('i');
            icon.className = isGuest ? 'bi bi-person-badge' : 'bi bi-person-fill';
            icon.setAttribute('aria-hidden', 'true');
            iconWrap.appendChild(icon);

            const main = document.createElement('div');
            main.className = 'pv-tree-main';
            const title = document.createElement('div');
            title.className = 'pv-tree-title';
            title.textContent = u.displayName || u.userPrincipalName || u.mail || '(ohne Namen)';
            const sub = document.createElement('div');
            sub.className = 'pv-tree-sub';
            sub.textContent = u.userPrincipalName || u.mail || u.id || '';
            main.appendChild(title);
            main.appendChild(sub);

            const meta = document.createElement('div');
            meta.className = 'pv-tree-meta';
            const pillType = document.createElement('span');
            pillType.className = 'pill' + (isGuest ? '' : ' ok');
            pillType.textContent = userTypeLabel(u.userType);
            meta.appendChild(pillType);
            if (u.accountEnabled === false) {
                const pillOff = document.createElement('span');
                pillOff.className = 'pill err';
                pillOff.textContent = 'Inaktiv';
                meta.appendChild(pillOff);
            }
            if (u.department) {
                const pillDep = document.createElement('span');
                pillDep.className = 'pill muted-pill';
                pillDep.textContent = String(u.department).trim();
                meta.appendChild(pillDep);
            }

            btn.appendChild(iconWrap);
            btn.appendChild(main);
            btn.appendChild(meta);
            li.appendChild(btn);
            tree.appendChild(li);
        }
        updateProgressLine();
    }

    function dispVal(v) {
        if (v === undefined || v === null || v === '') return '–';
        return String(v);
    }

    function addProfileTextField(root, label, fieldKey, value, editable, fullWidth) {
        const wrap = document.createElement('div');
        wrap.className =
            'field ' + (editable ? 'field-editable' : 'field-readonly') + (fullWidth ? ' field-full' : '');
        const lab = document.createElement('label');
        lab.setAttribute('for', 'pv_f_' + fieldKey);
        lab.textContent = label;
        const inp = document.createElement('input');
        inp.type = 'text';
        inp.id = 'pv_f_' + fieldKey;
        inp.dataset.pvField = fieldKey;
        inp.readOnly = !editable;
        inp.autocomplete = 'off';
        if (editable && (value === undefined || value === null || value === '')) {
            inp.value = '';
            inp.placeholder = '–';
        } else {
            inp.value = dispVal(value);
        }
        wrap.appendChild(lab);
        wrap.appendChild(inp);
        root.appendChild(wrap);
    }

    function addProfileAccountEnabled(root, u, editable) {
        const wrap = document.createElement('div');
        wrap.className = 'field ' + (editable ? 'field-editable' : 'field-readonly');
        const lab = document.createElement('label');
        lab.setAttribute('for', 'pv_f_accountEnabled');
        lab.textContent = 'Konto aktiv';
        wrap.appendChild(lab);
        if (editable) {
            const row = document.createElement('div');
            row.className = 'pv-checkbox-row';
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.id = 'pv_f_accountEnabled';
            cb.dataset.pvField = 'accountEnabled';
            cb.checked = u.accountEnabled !== false;
            const l2 = document.createElement('label');
            l2.htmlFor = 'pv_f_accountEnabled';
            l2.style.margin = '0';
            l2.style.fontWeight = '600';
            l2.textContent = 'Konto ist aktiviert';
            row.appendChild(cb);
            row.appendChild(l2);
            wrap.appendChild(row);
        } else {
            const inp = document.createElement('input');
            inp.type = 'text';
            inp.readOnly = true;
            inp.value = u.accountEnabled === false ? 'Nein' : u.accountEnabled === true ? 'Ja' : '–';
            if (inp.value === '–') inp.style.color = 'var(--muted)';
            wrap.appendChild(inp);
        }
        root.appendChild(wrap);
    }

    function renderProfileTab(u, editable) {
        const root = document.getElementById('pvProfileFields');
        if (!root) return;
        root.replaceChildren();
        if (!u) return;

        addProfileTextField(root, 'Anzeigename', 'displayName', u.displayName, editable, false);
        addProfileTextField(root, 'Alias (Mail-Nickname)', 'mailNickname', u.mailNickname, editable, false);
        addProfileTextField(root, 'Vorname', 'givenName', u.givenName, editable, false);
        addProfileTextField(root, 'Nachname', 'surname', u.surname, editable, false);
        addProfileTextField(
            root,
            'Benutzername (UPN)' + (String(u.userType).toLowerCase() === 'guest' ? ' (Gast)' : ''),
            'userPrincipalName',
            u.userPrincipalName,
            editable,
            false
        );
        addProfileTextField(root, 'E-Mail (SMTP)', 'mail', u.mail, editable, false);
        addProfileTextField(root, 'Position', 'jobTitle', u.jobTitle, editable, false);
        addProfileTextField(root, 'Abteilung', 'department', u.department, editable, false);
        addProfileTextField(root, 'Firma', 'companyName', u.companyName, editable, false);
        addProfileTextField(root, 'Bürostandort', 'officeLocation', u.officeLocation, editable, false);
        addProfileTextField(root, 'Straße', 'streetAddress', u.streetAddress, editable, true);
        addProfileTextField(root, 'PLZ', 'postalCode', u.postalCode, editable, false);
        addProfileTextField(root, 'Ort', 'city', u.city, editable, false);
        addProfileTextField(root, 'Land', 'country', u.country, editable, false);
        addProfileTextField(root, 'Mobiltelefon', 'mobilePhone', u.mobilePhone, editable, false);
        const bp0 = u.businessPhones && u.businessPhones[0] ? u.businessPhones[0] : '';
        addProfileTextField(root, 'Geschäftstelefon (1. Zeile)', 'businessPhone0', bp0, editable, false);
        addProfileTextField(root, 'Sprache (z. B. de-AT)', 'preferredLanguage', u.preferredLanguage, editable, false);
        addProfileAccountEnabled(root, u, editable);
        addProfileTextField(root, 'Kontotyp', '_userType', userTypeLabel(u.userType), false, false);
        addProfileTextField(root, 'Objekt-ID', '_id', u.id, false, true);
        addProfileTextField(root, 'Erstellt', '_created', formatDate(u.createdDateTime), false, false);
        const licSum = userLicenseSummary(u);
        const licText = licSum
            ? licSum.hasAny
                ? (licSum.licenses || [])
                      .map(function (l) {
                          return l.name || l.shortLabel;
                      })
                      .filter(Boolean)
                      .join(', ')
                : 'Keine'
            : '–';
        addProfileTextField(root, 'Lizenzen (Übersicht)', '_licenses', licText, false, true);
    }

    function readInputTrim(el) {
        if (!el) return '';
        return String(el.value || '').trim();
    }

    function buildPatchFromForm(u) {
        const root = document.getElementById('pvProfileFields');
        if (!root) return null;

        function get(field) {
            const el = root.querySelector('[data-pv-field="' + field + '"]');
            if (!el) return undefined;
            if (el.type === 'checkbox') return !!el.checked;
            const t = readInputTrim(el);
            return t === '' || t === '–' ? '' : t;
        }

        const patch = {};
        const strFields = [
            'displayName',
            'givenName',
            'surname',
            'userPrincipalName',
            'mail',
            'mailNickname',
            'jobTitle',
            'department',
            'companyName',
            'officeLocation',
            'streetAddress',
            'city',
            'postalCode',
            'country',
            'mobilePhone',
            'preferredLanguage'
        ];
        for (let i = 0; i < strFields.length; i++) {
            const k = strFields[i];
            let nv = get(k);
            if (nv === undefined) continue;
            if (k === 'mailNickname' && nv) {
                nv = sanitizeMailNickname(nv, '');
            }
            const ov = u[k] == null ? '' : String(u[k]);
            if (String(nv) !== ov) {
                patch[k] = nv === '' ? null : nv;
            }
        }

        const accEl = root.querySelector('[data-pv-field="accountEnabled"]');
        if (accEl && accEl.type === 'checkbox') {
            const nv = !!accEl.checked;
            const ov = u.accountEnabled !== false;
            if (nv !== ov) patch.accountEnabled = nv;
        }

        const bpNew = get('businessPhone0');
        if (bpNew !== undefined) {
            const ov = u.businessPhones && u.businessPhones[0] ? String(u.businessPhones[0]) : '';
            if (String(bpNew) !== ov) {
                patch.businessPhones = bpNew === '' ? [] : [bpNew];
            }
        }

        return patch;
    }

    async function refreshUserFromGraph(token, userId) {
        const path =
            '/users/' + encodeURIComponent(userId) + '?$select=' + encodeURIComponent(USER_REFRESH_SELECT);
        return graphJson('GET', path, token, undefined);
    }

    function mergeUserIntoList(updated) {
        const idx = loadedUsers.findIndex(function (x) {
            return x.id === updated.id;
        });
        if (idx === -1) {
            loadedUsers.push(updated);
        } else {
            loadedUsers[idx] = updated;
        }
        refreshDepartmentFilter();
        refreshLicenseFilter();
        updateStatsPanel();
    }

    async function saveProfilePatch() {
        const u = getSelectedUser();
        if (!u) return;
        const root = document.getElementById('pvProfileFields');
        const dnEl = root && root.querySelector('[data-pv-field="displayName"]');
        if (dnEl && readInputTrim(dnEl) === '') {
            toast('Anzeigename darf nicht leer sein.');
            return;
        }
        const upnEl = root && root.querySelector('[data-pv-field="userPrincipalName"]');
        if (upnEl && readInputTrim(upnEl) === '') {
            toast('UPN darf nicht leer sein.');
            return;
        }

        const patch = buildPatchFromForm(u);
        if (!patch || Object.keys(patch).length === 0) {
            toast('Keine Änderungen.');
            return;
        }
        const saveBtns = [
            document.getElementById('pvBtnSave'),
            document.getElementById('pvBtnSaveBottom')
        ].filter(Boolean);
        saveBtns.forEach(function (b) {
            b.disabled = true;
        });
        try {
            const token = await getGraphToken();
            await graphJson('PATCH', '/users/' + encodeURIComponent(u.id), token, patch);
            const fresh = await refreshUserFromGraph(token, u.id);
            mergeUserIntoList(fresh);
            appendLog('Profil gespeichert (PATCH).', 'ok');
            toast('Gespeichert.');
            profileEditMode = true;
            updateDetailActionButtons();
            renderProfileTab(fresh, true);
            renderUserTree();
        } catch (e) {
            const msg = graphErrorFriendly(e);
            appendLog('PATCH: ' + msg, 'err');
            toast(msg);
        } finally {
            saveBtns.forEach(function (b) {
                b.disabled = false;
            });
        }
    }

    async function resetProfileFromGraph() {
        const u = getSelectedUser();
        if (!u) return;
        try {
            const token = await getGraphToken();
            const fresh = await refreshUserFromGraph(token, u.id);
            mergeUserIntoList(fresh);
            profileEditMode = true;
            renderProfileTab(fresh, true);
            renderUserTree();
            appendLog('Profil neu geladen.', 'ok');
            toast('Profil neu geladen.');
        } catch (e) {
            const msg = graphErrorFriendly(e);
            appendLog('Profil laden: ' + msg, 'err');
            toast(msg);
            renderProfileTab(u, true);
        }
    }

    function sanitizeMailNickname(raw, fallbackFromUpn) {
        let s = String(raw || '').trim();
        if (!s && fallbackFromUpn) {
            const at = fallbackFromUpn.indexOf('@');
            s = at > 0 ? fallbackFromUpn.slice(0, at) : fallbackFromUpn;
        }
        s = s.split('@')[0].replace(/[^a-zA-Z0-9._-]/g, '');
        return s;
    }

    function openCreateModal() {
        const bd = document.getElementById('pvModalCreateBackdrop');
        if (!bd) return;
        const ids = ['pvCreateUpn', 'pvCreateDisplayName', 'pvCreateMailNick', 'pvCreatePassword', 'pvCreateGiven', 'pvCreateSurname', 'pvCreateMail'];
        for (let i = 0; i < ids.length; i++) {
            const el = document.getElementById(ids[i]);
            if (el) el.value = '';
        }
        const f = document.getElementById('pvCreateForcePw');
        if (f) f.checked = true;
        const e = document.getElementById('pvCreateEnabled');
        if (e) e.checked = true;
        bd.classList.add('active');
        bd.setAttribute('aria-hidden', 'false');
    }

    function closeCreateModal() {
        const bd = document.getElementById('pvModalCreateBackdrop');
        if (!bd) return;
        bd.classList.remove('active');
        bd.setAttribute('aria-hidden', 'true');
    }

    async function submitCreateUser() {
        const upn = readInputTrim(document.getElementById('pvCreateUpn'));
        const displayName = readInputTrim(document.getElementById('pvCreateDisplayName'));
        let mailNick = readInputTrim(document.getElementById('pvCreateMailNick'));
        const password = String(document.getElementById('pvCreatePassword')?.value || '');
        const givenName = readInputTrim(document.getElementById('pvCreateGiven'));
        const surname = readInputTrim(document.getElementById('pvCreateSurname'));
        const mail = readInputTrim(document.getElementById('pvCreateMail'));
        const forcePw = document.getElementById('pvCreateForcePw') ? document.getElementById('pvCreateForcePw').checked : true;
        const enabled = document.getElementById('pvCreateEnabled') ? document.getElementById('pvCreateEnabled').checked : true;

        if (!upn || !displayName || !password) {
            toast('UPN, Anzeigename und Kennwort sind Pflichtfelder.');
            return;
        }
        mailNick = sanitizeMailNickname(mailNick, upn);
        if (!mailNick) {
            toast('Mail-Nickname ungültig oder leer.');
            return;
        }

        const body = {
            accountEnabled: enabled,
            displayName: displayName,
            mailNickname: mailNick,
            userPrincipalName: upn,
            passwordProfile: {
                password: password,
                forceChangePasswordNextSignIn: !!forcePw
            }
        };
        if (givenName) body.givenName = givenName;
        if (surname) body.surname = surname;
        if (mail) body.mail = mail;

        const sub = document.getElementById('pvModalCreateSubmit');
        if (sub) sub.disabled = true;
        try {
            const token = await getGraphToken();
            const created = await graphJson('POST', '/users', token, body);
            const id = created && created.id ? created.id : null;
            appendLog('Benutzer angelegt: ' + (created.userPrincipalName || upn), 'ok');
            toast('Benutzer angelegt. Als Nächstes: Nutzungsort und Lizenz zuweisen.');
            closeCreateModal();
            if (id) {
                pendingTabAfterSelect = 'lizenzen';
                try {
                    const fresh = await refreshUserFromGraph(token, id);
                    mergeUserIntoList(fresh);
                    selectUser(id);
                } catch {
                    loadedUsers.push(created);
                    refreshDepartmentFilter();
                    updateStatsPanel();
                    if (id) selectUser(id);
                }
            }
            renderUserTree();
        } catch (e) {
            const msg = e && e.message ? e.message : String(e);
            appendLog('Anlegen: ' + msg, 'err');
            toast(msg);
        } finally {
            if (sub) sub.disabled = false;
        }
    }

    function updateDeleteModalUi() {
        const hard = document.getElementById('pvDeleteHard') && document.getElementById('pvDeleteHard').checked;
        const warn = document.getElementById('pvDeleteHardWarn');
        const intro = document.getElementById('pvDeleteSoftIntro');
        const sub = document.getElementById('pvModalDeleteSubmit');
        if (warn) warn.style.display = hard ? 'block' : 'none';
        if (intro) intro.style.opacity = hard ? '0.55' : '1';
        if (sub) {
            sub.textContent = hard ? 'Endgültig löschen' : 'Konto deaktivieren';
            sub.className = hard ? 'btn btn-danger' : 'btn btn-success';
        }
    }

    function openDeleteModal() {
        const u = getSelectedUser();
        if (!u) return;
        const bd = document.getElementById('pvModalDeleteBackdrop');
        const echo = document.getElementById('pvDeleteUpnEcho');
        const inp = document.getElementById('pvDeleteConfirmInput');
        const sub = document.getElementById('pvModalDeleteSubmit');
        const hardChk = document.getElementById('pvDeleteHard');
        if (hardChk) hardChk.checked = false;
        updateDeleteModalUi();
        if (echo) echo.textContent = u.userPrincipalName || u.mail || u.id;
        if (inp) inp.value = '';
        if (sub) sub.disabled = true;
        if (bd) {
            bd.classList.add('active');
            bd.setAttribute('aria-hidden', 'false');
        }
    }

    function closeDeleteModal() {
        const bd = document.getElementById('pvModalDeleteBackdrop');
        if (!bd) return;
        bd.classList.remove('active');
        bd.setAttribute('aria-hidden', 'true');
    }

    function syncDeleteConfirmButton() {
        const u = getSelectedUser();
        const inp = document.getElementById('pvDeleteConfirmInput');
        const sub = document.getElementById('pvModalDeleteSubmit');
        if (!sub || !inp || !u) return;
        const ok = readInputTrim(inp) === String(u.userPrincipalName || '').trim();
        sub.disabled = !ok;
    }

    async function submitDeleteUser() {
        const u = getSelectedUser();
        if (!u) return;
        const inp = document.getElementById('pvDeleteConfirmInput');
        if (!inp || readInputTrim(inp) !== String(u.userPrincipalName || '').trim()) {
            toast('UPN stimmt nicht überein.');
            return;
        }
        const hard = document.getElementById('pvDeleteHard') && document.getElementById('pvDeleteHard').checked;
        const sub = document.getElementById('pvModalDeleteSubmit');
        if (sub) sub.disabled = true;
        try {
            const token = await getGraphToken();
            if (!hard) {
                if (u.accountEnabled === false) {
                    toast('Konto ist bereits deaktiviert.');
                    closeDeleteModal();
                    return;
                }
                await graphJson('PATCH', '/users/' + encodeURIComponent(u.id), token, { accountEnabled: false });
                const fresh = await refreshUserFromGraph(token, u.id);
                mergeUserIntoList(fresh);
                appendLog('Konto deaktiviert: ' + (fresh.userPrincipalName || u.id), 'ok');
                toast('Konto deaktiviert.');
                closeDeleteModal();
                profileEditMode = false;
                cachedGroupsForSelection = null;
                selectUser(fresh.id);
                updateStatsPanel();
                return;
            }

            await graphDelete('/users/' + encodeURIComponent(u.id), token);
            loadedUsers = loadedUsers.filter(function (x) {
                return x.id !== u.id;
            });
            appendLog('Benutzer gelöscht (DELETE): ' + (u.userPrincipalName || u.id), 'ok');
            toast('Dauerhaft gelöscht.');
            closeDeleteModal();
            selectedUserId = null;
            cachedGroupsForSelection = null;
            profileEditMode = false;
            const hint = document.getElementById('pvHint');
            const detail = document.getElementById('pvDetail');
            if (hint) hint.style.display = '';
            if (detail) detail.style.display = 'none';
            updateDetailActionButtons();
            refreshDepartmentFilter();
            updateStatsPanel();
            renderUserTree();
        } catch (e) {
            const msg = e && e.message ? e.message : String(e);
            appendLog((hard ? 'Löschen: ' : 'Deaktivieren: ') + msg, 'err');
            toast(msg);
        } finally {
            if (sub) sub.disabled = false;
            syncDeleteConfirmButton();
        }
    }

    function renderGroupsTable(groups) {
        const tbody = document.getElementById('pvGroupsTbody');
        if (!tbody) return;
        tbody.replaceChildren();

        if (!groups || !groups.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 4;
            td.style.color = '#6c757d';
            td.textContent = 'Keine direkten Gruppenmitgliedschaften gefunden.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            return;
        }

        const sorted = groups.slice().sort(function (a, b) {
            return compareStrings(a.displayName, b.displayName);
        });

        for (let i = 0; i < sorted.length; i++) {
            const g = sorted[i];
            const tr = document.createElement('tr');
            const tdN = document.createElement('td');
            tdN.textContent = g.displayName || '–';
            const tdM = document.createElement('td');
            tdM.textContent = g.mail || g.mailNickname || '–';
            tdM.style.wordBreak = 'break-all';
            tdM.style.fontSize = '0.9em';
            const tdT = document.createElement('td');
            tdT.textContent = groupTypeLabel(g);
            tdT.style.fontSize = '0.88em';
            const tdAct = document.createElement('td');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn small-btn';
            btn.setAttribute('data-pv-group-remove', g.id || '');
            btn.textContent = 'Entfernen';
            tdAct.appendChild(btn);
            tr.appendChild(tdN);
            tr.appendChild(tdM);
            tr.appendChild(tdT);
            tr.appendChild(tdAct);
            tbody.appendChild(tr);
        }
    }

    function odataEscape(s) {
        return String(s || '').replace(/'/g, "''");
    }

    function isGuid(s) {
        return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(s || '').trim());
    }

    function isDuplicateMemberError(e) {
        const m = String((e && e.message) || e || '');
        return (
            m.indexOf('added object references already exist') !== -1 ||
            m.indexOf('One or more added object references already exist') !== -1 ||
            m.indexOf('already exist') !== -1
        );
    }

    function memberGroupIds() {
        const set = new Set();
        (cachedGroupsForSelection || []).forEach(function (g) {
            const id = String((g && g.id) || '').toLowerCase();
            if (id) set.add(id);
        });
        return set;
    }

    function fillGroupSearchResults(groups) {
        const container = document.getElementById('pvGroupSearchResults');
        if (!container) return;
        container.replaceChildren();
        const already = memberGroupIds();
        const filtered = (groups || []).filter(function (g) {
            return g && g.id && !already.has(String(g.id).toLowerCase());
        });
        if (!filtered.length) {
            const hint = document.createElement('span');
            hint.className = 'pv-group-checklist-hint';
            hint.textContent = groups && groups.length ? '(alle Treffer bereits Mitglied)' : '(keine Treffer)';
            container.appendChild(hint);
            return;
        }
        // Alle auswählen-Zeile
        const selAllRow = document.createElement('label');
        selAllRow.className = 'pv-group-checklist-selectall';
        const selAllCb = document.createElement('input');
        selAllCb.type = 'checkbox';
        selAllCb.style.width = '16px';
        selAllCb.style.height = '16px';
        selAllCb.style.cursor = 'pointer';
        selAllCb.style.accentColor = '#5e72e4';
        selAllCb.setAttribute('aria-label', 'Alle auswählen');
        selAllRow.appendChild(selAllCb);
        const selAllTxt = document.createElement('span');
        selAllTxt.textContent = 'Alle auswählen (' + filtered.length + ')';
        selAllRow.appendChild(selAllTxt);
        container.appendChild(selAllRow);

        filtered.forEach(function (g) {
            const mail = g.mail || g.mailNickname || '';
            const label = document.createElement('label');
            label.className = 'pv-group-checklist-item';
            label.setAttribute('data-pv-group-id', g.id);
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = g.id;
            cb.addEventListener('change', function () {
                label.classList.toggle('is-checked', cb.checked);
                // Alle-auswählen-Checkbox aktualisieren
                const allCbs = container.querySelectorAll('.pv-group-checklist-item input[type="checkbox"]');
                const checkedCount = container.querySelectorAll('.pv-group-checklist-item input[type="checkbox"]:checked').length;
                selAllCb.indeterminate = checkedCount > 0 && checkedCount < allCbs.length;
                selAllCb.checked = checkedCount === allCbs.length;
            });
            label.appendChild(cb);
            const txt = document.createElement('span');
            txt.textContent = (g.displayName || g.id) + (mail ? ' · ' + mail : '') + ' · ' + groupTypeLabel(g);
            label.appendChild(txt);
            container.appendChild(label);
        });

        selAllCb.addEventListener('change', function () {
            const allCbs = container.querySelectorAll('.pv-group-checklist-item input[type="checkbox"]');
            allCbs.forEach(function (cb) {
                cb.checked = selAllCb.checked;
                cb.closest('.pv-group-checklist-item').classList.toggle('is-checked', selAllCb.checked);
            });
        });
    }

    async function searchDirectoryGroups(token, queryRaw) {
        const q = String(queryRaw || '').trim();
        if (!q) return [];
        const select = 'id,displayName,mail,mailNickname,groupTypes,securityEnabled,mailEnabled';
        if (isGuid(q)) {
            try {
                const g = await graphJson(
                    'GET',
                    '/groups/' + encodeURIComponent(q) + '?$select=' + encodeURIComponent(select),
                    token
                );
                return g && g.id ? [g] : [];
            } catch {
                return [];
            }
        }
        try {
            const phrase = q.replace(/"/g, '\\"').replace(/\r?\n/g, ' ').trim();
            const aqs =
                '(displayName:' + phrase + ' OR mail:' + phrase + ' OR mailNickname:' + phrase + ')';
            const path =
                '/groups?$search=' +
                encodeURIComponent('"' + aqs + '"') +
                '&$select=' +
                encodeURIComponent(select) +
                '&$top=25';
            const data = await graphJson('GET', path, token, undefined, { ConsistencyLevel: 'eventual' });
            return Array.isArray(data.value) ? data.value : [];
        } catch {
            // Fallback ohne $search
        }
        const esc = odataEscape(q);
        const filter =
            "startswith(displayName,'" +
            esc +
            "') or startswith(mailNickname,'" +
            esc +
            "') or startswith(mail,'" +
            esc +
            "')";
        const path =
            '/groups?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent(select) +
            '&$top=25';
        const data = await graphJson('GET', path, token);
        return Array.isArray(data.value) ? data.value : [];
    }

    async function fetchUserGroups(token, userId) {
        const path =
            '/users/' +
            encodeURIComponent(userId) +
            '/memberOf/microsoft.graph.group?$select=' +
            encodeURIComponent(GROUP_MEMBEROF_SELECT) +
            '&$top=999';
        return fetchAllPages(token, path, undefined);
    }

    async function loadGroupsForSelected() {
        const prog = document.getElementById('pvGroupsProgress');
        if (!selectedUserId) return;
        if (prog) prog.textContent = 'Lade Gruppen …';

        const tbody = document.getElementById('pvGroupsTbody');
        if (tbody) {
            tbody.replaceChildren();
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 4;
            td.style.color = '#6c757d';
            td.textContent = 'Lade …';
            tr.appendChild(td);
            tbody.appendChild(tr);
        }

        try {
            const token = await getGraphToken();
            const groups = await fetchUserGroups(token, selectedUserId);
            cachedGroupsForSelection = groups;
            renderGroupsTable(groups);
            if (prog) prog.textContent = groups.length ? groups.length + ' Gruppe(n).' : 'Keine Einträge.';
            appendLog('Gruppen für ausgewählte Person: ' + groups.length, 'ok');
        } catch (e) {
            cachedGroupsForSelection = [];
            renderGroupsTable([]);
            const msg = graphErrorFriendly(e);
            if (prog) prog.textContent = 'Fehler: ' + msg;
            appendLog('Gruppen laden: ' + msg, 'err');
            toast('Gruppen: ' + msg);
        }
    }

    async function searchGroupsForAdd() {
        const inp = document.getElementById('pvGroupSearch');
        const q = inp && inp.value ? String(inp.value).trim() : '';
        if (!q) {
            toast('Bitte einen Gruppennamen, Alias oder eine ID eingeben.');
            return;
        }
        const btn = document.getElementById('pvGroupSearchBtn');
        const status = document.getElementById('pvGroupsProgress');
        if (btn) btn.disabled = true;
        try {
            const token = await getGraphToken();
            const list = await searchDirectoryGroups(token, q);
            fillGroupSearchResults(list);
            if (status) {
                status.textContent = list.length
                    ? 'Suche: ' + list.length + ' Treffer.'
                    : 'Suche: keine Treffer.';
            }
        } catch (e) {
            const msg = graphErrorFriendly(e);
            fillGroupSearchResults([]);
            if (status) status.textContent = 'Suche: ' + msg;
            toast(msg);
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    async function addSelectedUserToGroup() {
        const u = getSelectedUser();
        const container = document.getElementById('pvGroupSearchResults');
        const checkedBoxes = container
            ? Array.from(container.querySelectorAll('.pv-group-checklist-item input[type="checkbox"]:checked'))
            : [];
        if (!u || !checkedBoxes.length) {
            toast('Bitte zuerst mindestens eine Gruppe aus den Treffern auswählen.');
            return;
        }
        if (groupBusy) return;
        const groups = checkedBoxes.map(function (cb) {
            const lbl = cb.closest('.pv-group-checklist-item');
            return { id: cb.value, label: lbl ? (lbl.querySelector('span') ? lbl.querySelector('span').textContent : cb.value) : cb.value };
        });
        const groupNames = groups.map(function (g) { return '· ' + g.label; }).join('\n');
        if (
            !(await dlgConfirm(
                'Diese Person zu ' + groups.length + ' Gruppe(n) hinzufügen?\n\n' +
                    (u.displayName || u.userPrincipalName || '') +
                    '\n\n' + groupNames,
                { title: 'Zu Gruppen hinzufügen', okText: 'Hinzufügen' }
            ))
        ) {
            return;
        }
        const btn = document.getElementById('pvGroupAddBtn');
        groupBusy = true;
        if (btn) btn.disabled = true;
        try {
            const token = await getGraphToken();
            let ok = 0, fail = 0;
            for (const g of groups) {
                try {
                    await graphJson('POST', '/groups/' + encodeURIComponent(g.id) + '/members/$ref', token, {
                        '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + u.id
                    });
                    appendLog('Mitglied hinzugefügt: ' + (u.displayName || u.id) + ' → ' + g.label, 'ok');
                    ok++;
                } catch (e) {
                    if (isDuplicateMemberError(e)) {
                        appendLog('Bereits Mitglied: ' + g.label, 'warn');
                        ok++;
                    } else {
                        appendLog('Fehler bei ' + g.label + ': ' + graphErrorFriendly(e), 'err');
                        fail++;
                    }
                }
            }
            toast(ok + ' Gruppe(n) hinzugefügt' + (fail ? ', ' + fail + ' Fehler.' : '.'));
            cachedGroupsForSelection = null;
            if (container) {
                container.replaceChildren();
                const hint = document.createElement('span');
                hint.className = 'pv-group-checklist-hint';
                hint.textContent = '(zuerst suchen)';
                container.appendChild(hint);
            }
            await loadGroupsForSelected();
        } catch (e) {
            const msg = graphErrorFriendly(e);
            appendLog('Gruppe hinzufügen: ' + msg, 'err');
            toast(msg);
        } finally {
            groupBusy = false;
            if (btn) btn.disabled = false;
        }
    }

    async function removeUserFromGroup(groupIdRaw) {
        const u = getSelectedUser();
        const groupId = String(groupIdRaw || '').trim();
        if (!u || !groupId || groupBusy) return;
        const g = (cachedGroupsForSelection || []).find(function (x) {
            return x && x.id === groupId;
        });
        const label = (g && (g.displayName || g.mail)) || groupId;
        if (
            !(await dlgConfirm(
                'Mitgliedschaft entfernen?\n\n' +
                    (u.displayName || u.userPrincipalName || '') +
                    '\n← ' +
                    label,
                { title: 'Aus Gruppe entfernen', okText: 'Entfernen', danger: true }
            ))
        ) {
            return;
        }
        groupBusy = true;
        try {
            const token = await getGraphToken();
            await graphJson(
                'DELETE',
                '/groups/' + encodeURIComponent(groupId) + '/members/' + encodeURIComponent(u.id) + '/$ref',
                token
            );
            appendLog('Mitglied entfernt: ' + (u.displayName || u.id) + ' ← ' + label, 'ok');
            toast('Aus der Gruppe entfernt.');
            cachedGroupsForSelection = null;
            await loadGroupsForSelected();
        } catch (e) {
            const msg = graphErrorFriendly(e);
            appendLog('Gruppe entfernen: ' + msg, 'err');
            toast(msg);
        } finally {
            groupBusy = false;
        }
    }

    function assignedSkuIdsOfUser(u) {
        const list = u && Array.isArray(u.assignedLicenses) ? u.assignedLicenses : [];
        return list
            .map(function (l) {
                return String((l && l.skuId) || '').toLowerCase();
            })
            .filter(Boolean);
    }

    function skuLookupFromSubscribed() {
        const map = new Map();
        subscribedSkus.forEach(function (s) {
            const id = String((s && s.skuId) || '').toLowerCase();
            if (!id) return;
            map.set(id, { skuId: id, skuPartNumber: String((s && s.skuPartNumber) || '') });
        });
        return map;
    }

    function renderLicenseTab() {
        const u = getSelectedUser();
        const hint = document.getElementById('pvLicHint');
        const usage = document.getElementById('pvLicUsageLocation');
        const tbody = document.getElementById('pvLicAssignedBody');
        const sel = document.getElementById('pvLicAssignSelect');
        const status = document.getElementById('pvLicStatus');
        if (!u) return;
        if (usage) usage.value = String(u.usageLocation || '').toUpperCase();
        if (hint) {
            hint.textContent = subscribedSkusOk
                ? 'Zuweisen und Entziehen über Microsoft Graph (assignLicense). Nutzungsort ist ein zweistelliger Ländercode (Österreich: AT).'
                : 'Mandanten-SKUs konnten nicht gelesen werden (Organization.Read.All). Zuweisen über den Education-Katalog ist möglich; Graph lehnt unbekannte SKUs ab.';
        }
        if (status) status.textContent = '';

        const api = Lic();
        const lookup = skuLookupFromSubscribed();
        const sum = api && typeof api.summarizeUserLicenses === 'function' ? api.summarizeUserLicenses(u, lookup) : null;
        const licenses = sum && Array.isArray(sum.licenses) ? sum.licenses : [];

        if (tbody) {
            tbody.replaceChildren();
            if (!licenses.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 3;
                td.style.color = '#6c757d';
                td.textContent = 'Keine Lizenz zugewiesen.';
                tr.appendChild(td);
                tbody.appendChild(tr);
            } else {
                licenses.forEach(function (lic) {
                    const tr = document.createElement('tr');
                    const tdName = document.createElement('td');
                    tdName.textContent = lic.name || lic.shortLabel || lic.skuId;
                    const tdSku = document.createElement('td');
                    const code = document.createElement('code');
                    code.textContent = String(lic.skuPartNumber || lic.skuId || '').slice(0, 42);
                    tdSku.appendChild(code);
                    const tdAct = document.createElement('td');
                    const btn = document.createElement('button');
                    btn.type = 'button';
                    btn.className = 'btn small-btn';
                    btn.setAttribute('data-pv-lic-remove', lic.skuId);
                    btn.textContent = 'Entziehen';
                    tdAct.appendChild(btn);
                    tr.appendChild(tdName);
                    tr.appendChild(tdSku);
                    tr.appendChild(tdAct);
                    tbody.appendChild(tr);
                });
            }
        }

        if (sel) {
            const assigned = assignedSkuIdsOfUser(u);
            const opts =
                api && typeof api.buildAssignableSkuOptions === 'function'
                    ? api.buildAssignableSkuOptions(subscribedSkus, assigned, {
                          fallbackCatalog: !subscribedSkusOk
                      })
                    : [];
            sel.replaceChildren();
            const o0 = document.createElement('option');
            o0.value = '';
            o0.textContent = opts.length ? '(Lizenz wählen)' : '(keine freie Lizenz)';
            sel.appendChild(o0);
            opts.forEach(function (o) {
                const opt = document.createElement('option');
                opt.value = o.skuId;
                const rest = o.remaining == null ? '' : ' · ' + o.remaining + ' frei';
                opt.textContent = (o.name || o.shortLabel) + rest;
                opt.disabled = !!o.disabled;
                sel.appendChild(opt);
            });
        }
    }

    async function loadSubscribedSkus(token) {
        try {
            const data = await graphJson(
                'GET',
                '/subscribedSkus?$select=' +
                    encodeURIComponent('skuId,skuPartNumber,prepaidUnits,consumedUnits,capabilityStatus'),
                token
            );
            subscribedSkus = Array.isArray(data.value) ? data.value : [];
            subscribedSkusOk = true;
            appendLog('Mandanten-Lizenzen: ' + subscribedSkus.length + ' SKU(s).', 'ok');
        } catch (e) {
            subscribedSkus = [];
            subscribedSkusOk = false;
            appendLog('Mandanten-Lizenzen nicht lesbar: ' + graphErrorFriendly(e), 'warn');
        }
    }

    async function ensureUsageLocation(token, u, locationHint) {
        const cur = String((u && u.usageLocation) || '').trim().toUpperCase();
        if (/^[A-Z]{2}$/.test(cur)) return cur;
        let next = String(locationHint || '').trim().toUpperCase();
        if (!/^[A-Z]{2}$/.test(next)) {
            const asked = await dlgPrompt(
                'Für die Lizenzzuweisung braucht das Konto einen Nutzungsort (Ländercode, z. B. AT).',
                'AT',
                { title: 'Nutzungsort', inputLabel: 'Ländercode', okText: 'Setzen' }
            );
            if (asked == null) return '';
            next = String(asked).trim().toUpperCase();
        }
        if (!/^[A-Z]{2}$/.test(next)) {
            toast('Ungültiger Ländercode (zwei Buchstaben, z. B. AT).');
            return '';
        }
        await graphJson('PATCH', '/users/' + encodeURIComponent(u.id), token, { usageLocation: next });
        u.usageLocation = next;
        appendLog('Nutzungsort gesetzt: ' + next, 'ok');
        return next;
    }

    async function assignSelectedLicense() {
        const u = getSelectedUser();
        const sel = document.getElementById('pvLicAssignSelect');
        const skuId = sel && sel.value ? String(sel.value).toLowerCase() : '';
        if (!u || !skuId) {
            toast('Bitte eine Lizenz wählen.');
            return;
        }
        if (licenseBusy) return;
        const optLabel = sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].textContent : skuId;
        const isGuest = String(u.userType || '').toLowerCase() === 'guest';
        if (
            isGuest &&
            !(await dlgConfirm(
                'Gäste erhalten selten Education-Lizenzen. Trotzdem zuweisen?\n\n' + optLabel,
                { title: 'Lizenz zuweisen', okText: 'Zuweisen' }
            ))
        ) {
            return;
        }
        if (
            !isGuest &&
            !(await dlgConfirm('Lizenz zuweisen?\n\n' + optLabel + '\n\n' + (u.displayName || u.userPrincipalName || ''), {
                title: 'Lizenz zuweisen',
                okText: 'Zuweisen'
            }))
        ) {
            return;
        }
        const usageEl = document.getElementById('pvLicUsageLocation');
        const btn = document.getElementById('pvLicAssignBtn');
        licenseBusy = true;
        if (btn) btn.disabled = true;
        try {
            const token = await getGraphToken();
            const loc = await ensureUsageLocation(token, u, usageEl ? usageEl.value : '');
            if (!loc) return;
            await graphJson('POST', '/users/' + encodeURIComponent(u.id) + '/assignLicense', token, {
                addLicenses: [{ skuId: skuId, disabledPlans: [] }],
                removeLicenses: []
            });
            const fresh = await refreshUserFromGraph(token, u.id);
            mergeUserIntoList(fresh);
            appendLog('Lizenz zugewiesen: ' + optLabel, 'ok');
            toast('Lizenz zugewiesen.');
            renderProfileTab(fresh, profileEditMode);
            renderUserTree();
            await loadSubscribedSkus(token);
            renderLicenseTab();
        } catch (e) {
            const msg = graphErrorFriendly(e);
            appendLog('Lizenz zuweisen: ' + msg, 'err');
            toast(msg);
        } finally {
            licenseBusy = false;
            if (btn) btn.disabled = false;
        }
    }

    async function removeLicense(skuIdRaw) {
        const u = getSelectedUser();
        const skuId = String(skuIdRaw || '').toLowerCase();
        if (!u || !skuId || licenseBusy) return;
        const api = Lic();
        const lookup = skuLookupFromSubscribed();
        const info = api && typeof api.resolveSku === 'function' ? api.resolveSku(skuId) : { name: skuId };
        const lookupPart = lookup.get(skuId);
        const label =
            api && lookupPart
                ? api.resolveSku(skuId, lookupPart.skuPartNumber).name
                : info.name || skuId;
        if (
            !(await dlgConfirm('Lizenz entziehen?\n\n' + label + '\n\n' + (u.displayName || u.userPrincipalName || ''), {
                title: 'Lizenz entziehen',
                okText: 'Entziehen',
                danger: true
            }))
        ) {
            return;
        }
        licenseBusy = true;
        try {
            const token = await getGraphToken();
            await graphJson('POST', '/users/' + encodeURIComponent(u.id) + '/assignLicense', token, {
                addLicenses: [],
                removeLicenses: [skuId]
            });
            const fresh = await refreshUserFromGraph(token, u.id);
            mergeUserIntoList(fresh);
            appendLog('Lizenz entzogen: ' + label, 'ok');
            toast('Lizenz entzogen.');
            renderProfileTab(fresh, profileEditMode);
            renderUserTree();
            await loadSubscribedSkus(token);
            renderLicenseTab();
        } catch (e) {
            const msg = graphErrorFriendly(e);
            appendLog('Lizenz entziehen: ' + msg, 'err');
            toast(msg);
        } finally {
            licenseBusy = false;
        }
    }

    async function saveUsageLocation() {
        const u = getSelectedUser();
        const inp = document.getElementById('pvLicUsageLocation');
        if (!u || !inp) return;
        const next = String(inp.value || '').trim().toUpperCase();
        if (!/^[A-Z]{2}$/.test(next)) {
            toast('Ländercode: zwei Buchstaben, z. B. AT.');
            return;
        }
        try {
            const token = await getGraphToken();
            await graphJson('PATCH', '/users/' + encodeURIComponent(u.id), token, { usageLocation: next });
            const fresh = await refreshUserFromGraph(token, u.id);
            mergeUserIntoList(fresh);
            appendLog('Nutzungsort gespeichert: ' + next, 'ok');
            toast('Nutzungsort gespeichert.');
            renderLicenseTab();
        } catch (e) {
            const msg = graphErrorFriendly(e);
            appendLog('Nutzungsort: ' + msg, 'err');
            toast(msg);
        }
    }

    function setTab(tab) {
        if (tab === 'gruppen') activeTab = 'gruppen';
        else if (tab === 'lizenzen') activeTab = 'lizenzen';
        else activeTab = 'profil';

        const rows = [
            ['profil', 'pvPanelProfil', 'pvTabProfil'],
            ['lizenzen', 'pvPanelLizenzen', 'pvTabLizenzen'],
            ['gruppen', 'pvPanelGruppen', 'pvTabGruppen']
        ];
        rows.forEach(function (row) {
            const on = activeTab === row[0];
            const p = document.getElementById(row[1]);
            const b = document.getElementById(row[2]);
            if (p) {
                p.classList.toggle('active', on);
                p.setAttribute('aria-hidden', on ? 'false' : 'true');
            }
            if (b) b.setAttribute('aria-selected', on ? 'true' : 'false');
        });

        if (activeTab === 'gruppen' && selectedUserId) {
            if (cachedGroupsForSelection === null) {
                loadGroupsForSelected();
            }
        }
        if (activeTab === 'lizenzen' && selectedUserId) {
            renderLicenseTab();
        }
    }

    function selectUser(userId) {
        selectedUserId = userId || null;
        cachedGroupsForSelection = null;
        activeTab = 'profil';
        profileEditMode = !!selectedUserId;
        const grpSel = document.getElementById('pvGroupSearchResults');
        if (grpSel) {
            grpSel.replaceChildren();
            const hint = document.createElement('span');
            hint.className = 'pv-group-checklist-hint';
            hint.textContent = '(zuerst suchen)';
            grpSel.appendChild(hint);
        }
        const grpQ = document.getElementById('pvGroupSearch');
        if (grpQ) grpQ.value = '';

        const hint = document.getElementById('pvHint');
        const detail = document.getElementById('pvDetail');
        const title = document.getElementById('pvManageTitle');

        if (!selectedUserId) {
            if (hint) hint.style.display = '';
            if (detail) detail.style.display = 'none';
            updateDetailActionButtons();
            renderUserTree();
            return;
        }

        const u = getSelectedUser();

        if (hint) hint.style.display = 'none';
        if (detail) detail.style.display = '';
        if (title) title.textContent = u && u.displayName ? String(u.displayName) : '(ohne Anzeigename)';

        renderProfileTab(u || null, true);
        updateDetailActionButtons();
        setTab(pendingTabAfterSelect || 'profil');
        pendingTabAfterSelect = '';
        renderUserTree();
    }

    async function loadUsers() {
        const btn = document.getElementById('pvBtnLoad');
        const btnCsv = document.getElementById('pvBtnCsv');
        const progress = document.getElementById('pvProgress');
        if (btn) btn.disabled = true;
        if (btnCsv) btnCsv.disabled = true;
        clearLog();
        loadedUsers = [];
        selectedUserId = null;
        cachedGroupsForSelection = null;
        profileEditMode = false;
        const hint = document.getElementById('pvHint');
        const detail = document.getElementById('pvDetail');
        if (hint) hint.style.display = '';
        if (detail) detail.style.display = 'none';
        updateDetailActionButtons();

        try {
            const token = await getGraphToken();
            appendLog('Lade Benutzer aus dem Verzeichnis …', '');

            const initial =
                '/users?$select=' +
                encodeURIComponent(USER_LIST_SELECT) +
                '&$top=999&$orderby=displayName';

            let users;
            try {
                users = await fetchAllPages(token, initial, function (count) {
                    if (progress) {
                        progress.textContent = 'Gelesen: ' + count + ' Person(en) …';
                    }
                });
            } catch (firstErr) {
                appendLog(
                    'Mit Sortierung fehlgeschlagen, lade ohne $orderby … ' +
                        (firstErr && firstErr.message ? firstErr.message : ''),
                    'warn'
                );
                const fallback =
                    '/users?$select=' + encodeURIComponent(USER_LIST_SELECT) + '&$top=999';
                users = await fetchAllPages(token, fallback, function (count) {
                    if (progress) {
                        progress.textContent = 'Gelesen: ' + count + ' Person(en) …';
                    }
                });
            }

            users.sort(function (a, b) {
                return compareStrings(a.displayName, b.displayName);
            });
            loadedUsers = users;
            appendLog('Fertig: ' + users.length + ' Person(en).', 'ok');
            await loadSubscribedSkus(token);
            saveUsersToSession();
            hideCacheBanner();
            refreshDepartmentFilter();
            refreshLicenseFilter();
            updateStatsPanel();
            if (progress) progress.textContent = '';
            updateProgressLine();
        } catch (e) {
            appendLog('Laden: ' + (e && e.message ? e.message : String(e)), 'err');
            toast(String(e && e.message ? e.message : e));
            if (progress) progress.textContent = '';
            updateStatsPanel();
        } finally {
            if (btn) btn.disabled = false;
            if (btnCsv) btnCsv.disabled = !loadedUsers.length;
            renderUserTree();
        }
    }

    function exportCsv() {
        if (!loadedUsers.length) {
            toast('Keine Daten zum Exportieren.');
            return;
        }
        const rows = getVisibleRows();
        const headers = [
            'displayName',
            'userPrincipalName',
            'mail',
            'department',
            'jobTitle',
            'id',
            'accountEnabled',
            'userType',
            'license'
        ];
        const lines = [headers.join(';')];
        for (let i = 0; i < rows.length; i++) {
            const u = rows[i];
            const cells = [];
            for (let h = 0; h < headers.length; h++) {
                const key = headers[h];
                let v;
                if (key === 'license') {
                    const sum = userLicenseSummary(u);
                    v = sum && sum.hasAny ? sum.primaryLabel : '';
                } else {
                    v = u[key];
                }
                if (v === undefined || v === null) v = '';
                v = String(v).replace(/"/g, '""');
                cells.push('"' + v + '"');
            }
            lines.push(cells.join(';'));
        }
        const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'personen-export.csv';
        a.click();
        URL.revokeObjectURL(a.href);
        appendLog('CSV exportiert (' + rows.length + ' Zeilen).', 'ok');
    }

    function bind() {
        const btnLoad = document.getElementById('pvBtnLoad');
        const btnCsv = document.getElementById('pvBtnCsv');
        const filt = document.getElementById('pvFilterText');
        const tree = document.getElementById('pvTree');
        const reRender = function () {
            renderUserTree();
        };

        if (btnLoad) btnLoad.addEventListener('click', () => loadUsers());
        if (btnCsv) {
            btnCsv.disabled = true;
            btnCsv.addEventListener('click', () => exportCsv());
        }
        if (filt) filt.addEventListener('input', reRender);

        const ft = document.getElementById('pvFilterUserType');
        const fa = document.getElementById('pvFilterAccount');
        const fd = document.getElementById('pvFilterDepartment');
        const fl = document.getElementById('pvFilterLicense');
        const fs = document.getElementById('pvSortKey');
        if (ft) ft.addEventListener('change', reRender);
        if (fa) fa.addEventListener('change', reRender);
        if (fd) fd.addEventListener('change', reRender);
        if (fl) fl.addEventListener('change', reRender);
        if (fs) fs.addEventListener('change', reRender);

        if (tree) {
            tree.addEventListener('click', function (ev) {
                const t = ev.target;
                if (!t || !t.closest) return;
                const btn = t.closest('button[data-pv-select-user]');
                if (!btn) return;
                const uid = btn.getAttribute('data-pv-select-user');
                selectUser(uid || null);
            });
        }

        document.querySelectorAll('.detail-tab-btn[data-pv-tab]').forEach(function (b) {
            b.addEventListener('click', function () {
                setTab(b.getAttribute('data-pv-tab'));
            });
        });

        const licPanel = document.getElementById('pvPanelLizenzen');
        if (licPanel) {
            licPanel.addEventListener('click', function (ev) {
                const t = ev.target;
                if (!t || !t.closest) return;
                const rm = t.closest('[data-pv-lic-remove]');
                if (rm) removeLicense(rm.getAttribute('data-pv-lic-remove'));
            });
        }
        document.getElementById('pvLicAssignBtn')?.addEventListener('click', function () {
            assignSelectedLicense();
        });
        document.getElementById('pvLicUsageSave')?.addEventListener('click', function () {
            saveUsageLocation();
        });

        document.getElementById('pvGroupSearchBtn')?.addEventListener('click', function () {
            searchGroupsForAdd();
        });
        document.getElementById('pvGroupSearch')?.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                searchGroupsForAdd();
            }
        });
        document.getElementById('pvGroupAddBtn')?.addEventListener('click', function () {
            addSelectedUserToGroup();
        });
        document.getElementById('pvGroupsReloadBtn')?.addEventListener('click', function () {
            cachedGroupsForSelection = null;
            loadGroupsForSelected();
        });
        const grpPanel = document.getElementById('pvPanelGruppen');
        if (grpPanel) {
            grpPanel.addEventListener('click', function (ev) {
                const t = ev.target;
                if (!t || !t.closest) return;
                const rm = t.closest('[data-pv-group-remove]');
                if (rm) removeUserFromGroup(rm.getAttribute('data-pv-group-remove'));
            });
        }

        const btnNeu = document.getElementById('pvBtnNeu');
        if (btnNeu) btnNeu.addEventListener('click', () => openCreateModal());

        document.getElementById('pvModalCreateClose')?.addEventListener('click', closeCreateModal);
        document.getElementById('pvModalCreateCancel')?.addEventListener('click', closeCreateModal);
        document.getElementById('pvModalCreateSubmit')?.addEventListener('click', () => submitCreateUser());
        document.getElementById('pvModalCreateBackdrop')?.addEventListener('click', function (ev) {
            if (ev.target === ev.currentTarget) closeCreateModal();
        });

        document.getElementById('pvBtnCancelEdit')?.addEventListener('click', function () {
            resetProfileFromGraph();
        });
        document.getElementById('pvBtnSave')?.addEventListener('click', () => saveProfilePatch());
        document.getElementById('pvBtnSaveBottom')?.addEventListener('click', () => saveProfilePatch());
        document.getElementById('pvBtnDelete')?.addEventListener('click', () => openDeleteModal());

        document.getElementById('pvModalDeleteClose')?.addEventListener('click', closeDeleteModal);
        document.getElementById('pvModalDeleteCancel')?.addEventListener('click', closeDeleteModal);
        document.getElementById('pvModalDeleteSubmit')?.addEventListener('click', () => submitDeleteUser());
        document.getElementById('pvDeleteConfirmInput')?.addEventListener('input', syncDeleteConfirmButton);
        document.getElementById('pvDeleteHard')?.addEventListener('change', function () {
            updateDeleteModalUi();
            syncDeleteConfirmButton();
        });
        document.getElementById('pvModalDeleteBackdrop')?.addEventListener('click', function (ev) {
            if (ev.target === ev.currentTarget) closeDeleteModal();
        });

        // Cache-Restore: Daten aus sessionStorage wiederherstellen wenn vorhanden
        const cachedAt = loadUsersFromSession();
        if (cachedAt) {
            refreshDepartmentFilter();
            refreshLicenseFilter();
            updateStatsPanel();
            updateProgressLine();
            showCacheBanner(cachedAt);
            appendLog('Benutzerliste aus Sitzungs-Cache wiederhergestellt (' + loadedUsers.length + ' Person(en)).', 'ok');
        }

        updateDetailActionButtons();
        renderUserTree();

        try {
            const q = new URLSearchParams(window.location.search);
            const tab = String(q.get('tab') || '').toLowerCase();
            if (tab === 'lizenzen' || tab === 'gruppen' || tab === 'profil') pendingTabAfterSelect = tab;
            if (q.get('create') === '1') openCreateModal();
        } catch {
            /* ignore */
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bind);
    } else {
        bind();
    }
})();
