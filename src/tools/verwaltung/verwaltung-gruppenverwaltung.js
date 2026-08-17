(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-verwaltung-gruppe-v1';

    function gug() {
        const G = window.ms365GraphUnifiedGroups;
        if (!G) throw new Error('graph-unified-groups.js muss vor diesem Skript geladen werden.');
        return G;
    }

    function live() {
        const L = window.ms365SlgLiveDetails;
        if (!L) throw new Error('slg-live-details.js muss vor diesem Skript geladen werden.');
        return L;
    }

    async function getGraphToken() {
        return gug().getGraphToken();
    }

    /** @type {string|null} */
    let matchedGroupId = null;

    /** @type {{ members: string[], direktion: string[], rows: { role: string, name: string, email: string }[] }} */
    let listCache = { members: [], direktion: [], rows: [] };

    /** @type {'general'|'owners'|'members'} */
    let activeTab = 'general';

    function toast(msg) {
        const el = document.getElementById('toast');
        if (el) {
            el.textContent = msg;
            el.classList.add('show');
            clearTimeout(toast._t);
            toast._t = setTimeout(function () {
                el.classList.remove('show');
            }, 3800);
        } else if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(msg);
        } else if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(msg);
        } else {
            window.alert(msg);
        }
    }

    function dlgConfirm(message, options) {
        if (typeof window.ms365AppDialogConfirm === 'function') {
            return window.ms365AppDialogConfirm(message, options || {});
        }
        return Promise.resolve(window.confirm(message));
    }

    function normStr(v) {
        return String(v ?? '').trim();
    }
    function normEmail(v) {
        return normStr(v).toLowerCase();
    }

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function isDirektionRole(roleRaw) {
        const r = normStr(roleRaw).toLowerCase();
        if (!r) return false;
        return r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1;
    }

    async function ensureOwners(token, groupId) {
        return gug().ensureOwners(token, groupId, listCache.direktion || []);
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
        if (!el) return;
        el.replaceChildren();
    }

    function loadTenantSettings() {
        if (typeof window.ms365TenantSettingsLoad !== 'function') return null;
        return window.ms365TenantSettingsLoad();
    }

    function readLists() {
        const settings = loadTenantSettings();
        const admin = settings && Array.isArray(settings.admin) ? settings.admin : [];
        const members = [];
        const direktion = [];
        const rows = [];
        const seenM = new Set();
        const seenD = new Set();
        admin.forEach(function (row) {
            const role = normStr(row && (row.role || row.rolle || row.title));
            const name = normStr(row && row.name);
            const email = normEmail(row && row.email);
            rows.push({ role: role, name: name, email: email });
            if (email && email.indexOf('@') !== -1 && !seenM.has(email)) {
                seenM.add(email);
                members.push(email);
            }
            if (!isDirektionRole(role)) return;
            if (!email || email.indexOf('@') === -1 || seenD.has(email)) return;
            seenD.add(email);
            direktion.push(email);
        });
        listCache = { members: members, direktion: direktion, rows: rows };
    }

    function contactLabel(row) {
        const parts = [];
        if (row.role) parts.push(row.role);
        if (row.name) parts.push(row.name);
        if (row.email) parts.push(row.email);
        return parts.join(' · ') || '–';
    }

    function renderContactBox(el, emptyText, limit) {
        if (!el) return;
        el.replaceChildren();
        const rows = listCache.rows || [];
        const first = typeof limit === 'number' ? rows.slice(0, limit) : rows;
        if (!first.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent = emptyText;
            el.appendChild(p);
            return;
        }
        first.forEach(function (row) {
            const d = document.createElement('div');
            d.textContent = contactLabel(row);
            d.style.padding = '4px 0';
            d.style.borderBottom = '1px solid #eef1f4';
            el.appendChild(d);
        });
        if (rows.length > first.length) {
            const more = document.createElement('div');
            more.className = 'muted';
            more.style.paddingTop = '8px';
            more.textContent = '… und ' + String(rows.length - first.length) + ' weitere.';
            el.appendChild(more);
        }
    }

    function updateLeftListUi() {
        const count = document.getElementById('slgVerwaltungCount');
        const line = document.getElementById('slgVerwaltungLine');
        if (count) count.textContent = String(listCache.members.length);
        if (line) line.textContent = matchedGroupId ? 'Gematcht: ' + matchedGroupId : 'Noch kein Match';
        renderContactBox(
            document.getElementById('slgContactPreview'),
            'Keine Einträge in der Verwaltungsliste.',
            40
        );
    }

    function renderOwnerPreview() {
        const el = document.getElementById('slgOwnerPreview');
        if (!el) return;
        el.replaceChildren();
        const list = listCache.direktion || [];
        if (!list.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent = 'Keine Direktion‑Owner in den Schul‑Einstellungen gefunden.';
            el.appendChild(p);
            return;
        }
        list.forEach(function (em) {
            const d = document.createElement('div');
            d.textContent = em;
            d.style.padding = '4px 0';
            d.style.borderBottom = '1px solid #eef1f4';
            el.appendChild(d);
        });
    }

    function renderMemberPreview() {
        renderContactBox(
            document.getElementById('slgMemberPreview'),
            'Keine E‑Mails in der Verwaltungsliste.',
            30
        );
    }

    function setTab(tab) {
        activeTab = tab === 'owners' || tab === 'members' ? tab : 'general';
        document.querySelectorAll('#slgDetailTabs .detail-tab-btn[data-slg-tab]').forEach(function (b) {
            const on = b.getAttribute('data-slg-tab') === activeTab;
            b.setAttribute('aria-selected', on ? 'true' : 'false');
        });
        document.querySelectorAll('[data-slg-tab-content]').forEach(function (p) {
            p.classList.toggle('active', p.getAttribute('data-slg-tab-content') === activeTab);
        });
        if (!matchedGroupId) {
            if (activeTab === 'owners') renderOwnerPreview();
            if (activeTab === 'members') renderMemberPreview();
            return;
        }
        live().onTab(activeTab, matchedGroupId);
    }

    function applyCreateDefaults() {
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const desc = document.getElementById('slgNewDescription');
        if (dn && !dn.value) dn.value = 'Schulverwaltung';
        if (nn && !nn.value) nn.value = 'verwaltung';
        if (desc && !desc.value) desc.value = 'Kontakte der Schulverwaltung (MS365-Schulverwaltung)';
    }

    function getActiveMatchedId() {
        return matchedGroupId;
    }

    function setActiveMatchedId(id) {
        matchedGroupId = id ? String(id) : null;
        live().resetCaches();
    }

    function renderGroupSearchResults(list) {
        const host = document.getElementById('slgGroupSearchResults');
        if (!host) return;
        host.replaceChildren();
        if (!list || !list.length) {
            host.style.display = 'block';
            const p = document.createElement('div');
            p.className = 'muted';
            p.textContent = 'Keine passenden Microsoft 365‑Gruppen (Unified) gefunden.';
            host.appendChild(p);
            return;
        }
        host.style.display = 'block';

        const box = document.createElement('div');
        box.style.border = '1px solid #ced4da';
        box.style.borderRadius = '12px';
        box.style.background = '#fff';
        box.style.overflow = 'hidden';

        list.forEach(function (g, idx) {
            const row = document.createElement('div');
            row.style.display = 'grid';
            row.style.gridTemplateColumns = '1fr auto';
            row.style.gap = '10px';
            row.style.padding = '10px 12px';
            row.style.borderTop = idx === 0 ? '0' : '1px solid #eef1f4';
            row.style.alignItems = 'center';

            const left = document.createElement('div');
            const dn = normStr(g && g.displayName) || '(ohne Namen)';
            const mail = normStr(g && g.mail) || '–';
            const nick = normStr(g && g.mailNickname) || '–';
            left.innerHTML =
                '<div style="font-weight:700;line-height:1.25;">' +
                escapeHtml(dn) +
                '</div>' +
                '<div class="muted" style="margin-top:2px;">Mail‑Nickname: <code>' +
                escapeHtml(nick) +
                '</code> · SMTP: ' +
                escapeHtml(mail) +
                '</div>' +
                '<div class="muted" style="margin-top:2px;">Gruppen‑ID: <code>' +
                escapeHtml(g && g.id ? g.id : '') +
                '</code></div>';

            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-success';
            btn.textContent = 'Matchen';
            btn.addEventListener('click', function () {
                if (!g || !g.id) return;
                setActiveMatchedId(String(g.id));
                saveState();
                live().fillForm(g);
                live().setMatchedMode(true);
                updateLeftListUi();
                live().loadGroup({ silent: true });
                toast('Gruppe gematcht.');
            });

            row.appendChild(left);
            row.appendChild(btn);
            box.appendChild(row);
        });
        host.appendChild(box);
    }

    async function runSearchGroups() {
        const inp = document.getElementById('slgGroupSearch');
        const q = inp && inp.value ? inp.value.trim() : '';
        if (!q) {
            toast('Bitte einen Suchbegriff eingeben.');
            return;
        }
        try {
            const token = await getGraphToken();
            const list = await gug().searchUnifiedGroups(token, q);
            renderGroupSearchResults(list);
            if (!list.length) toast('Keine passenden Gruppen gefunden.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function runCreateAndMatch() {
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const dd = document.getElementById('slgNewDescription');
        const ct = document.getElementById('slgNewCreateTeam');
        const displayName = dn ? dn.value : '';
        const mailNick = nn ? nn.value : '';
        const desc = dd ? dd.value : '';
        if (!normStr(displayName) || !normStr(mailNick)) {
            toast('Bitte Anzeigename und Alias/Mail‑Nickname ausfüllen.');
            return;
        }
        try {
            const token = await getGraphToken();
            const g = await gug().createUnifiedGroup(token, displayName, mailNick, desc);
            await ensureOwners(token, g.id);
            if (ct && ct.checked) {
                toast('Gruppe angelegt – Team wird bereitgestellt …');
                await gug().provisionTeamForGroup(token, g.id);
            }
            setActiveMatchedId(String(g.id));
            saveState();
            live().fillForm(g);
            live().setMatchedMode(true);
            updateLeftListUi();
            await live().loadGroup({ silent: true });
            toast('Gruppe angelegt und gematcht.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    function runUnmatch() {
        if (!getActiveMatchedId()) return;
        setActiveMatchedId(null);
        saveState();
        live().loadGroup({ silent: true });
        toast('Match gelöst.');
    }

    function openEntraForMatched() {
        const gid = getActiveMatchedId();
        if (!gid) return;
        const url =
            'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Members/groupId/' +
            encodeURIComponent(gid);
        window.open(url, '_blank', 'noopener');
    }

    async function runSyncMembers() {
        const gid = getActiveMatchedId();
        if (!gid) {
            toast('Zuerst eine Gruppe matchen oder anlegen.');
            return;
        }
        const emails = listCache.members || [];
        if (!emails.length) {
            toast('Keine E‑Mails in der Verwaltungsliste.');
            return;
        }
        clearSyncLog();
        appendSyncLog('Start: Verwaltung (' + emails.length + ' Adressen) …', '');
        try {
            const token = await getGraphToken();
            const r = await gug().syncEmailsToGroup(token, gid, emails, 'Verwaltung', appendSyncLog);
            appendSyncLog('Fertig: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            await ensureOwners(token, gid);
            live().invalidateMembership();
            await live().loadMembers();
            toast('Synchronisation abgeschlossen.');
        } catch (e) {
            appendSyncLog('Abbruch: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        }
    }

    function buildStateObject() {
        return {
            kind: STORAGE_KEY,
            savedAt: new Date().toISOString(),
            matched: {
                verwaltungGroupId: matchedGroupId
            },
            vwNewDisplayName: document.getElementById('slgNewDisplayName')
                ? document.getElementById('slgNewDisplayName').value
                : '',
            vwNewMailNick: document.getElementById('slgNewMailNick')
                ? document.getElementById('slgNewMailNick').value
                : '',
            vwNewDescription: document.getElementById('slgNewDescription')
                ? document.getElementById('slgNewDescription').value
                : '',
            vwNewCreateTeam: document.getElementById('slgNewCreateTeam')
                ? !!document.getElementById('slgNewCreateTeam').checked
                : false
        };
    }

    function applyStateObject(o) {
        if (!o || typeof o !== 'object') return;
        if (o.matched && typeof o.matched === 'object') {
            matchedGroupId = o.matched.verwaltungGroupId ? String(o.matched.verwaltungGroupId) : null;
        } else if (o.verwaltungGroupId) {
            matchedGroupId = String(o.verwaltungGroupId);
        }
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const dd = document.getElementById('slgNewDescription');
        const ct = document.getElementById('slgNewCreateTeam');
        if (dn && o.vwNewDisplayName !== undefined) dn.value = String(o.vwNewDisplayName || '');
        if (nn && o.vwNewMailNick !== undefined) nn.value = String(o.vwNewMailNick || '');
        if (dd && o.vwNewDescription !== undefined) dd.value = String(o.vwNewDescription || '');
        if (ct && o.vwNewCreateTeam !== undefined) ct.checked = !!o.vwNewCreateTeam;
        live().resetCaches();
        const host = document.getElementById('slgGroupSearchResults');
        if (host) {
            host.replaceChildren();
            host.style.display = 'none';
        }
        live().setMatchedMode(!!matchedGroupId);
        live().fillForm(matchedGroupId ? { id: matchedGroupId } : null);
        updateLeftListUi();
        setTab('general');
    }

    function saveState() {
        try {
            const obj = buildStateObject();
            localStorage.setItem(STORAGE_KEY, JSON.stringify(obj));
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: { verwaltungGroupId: obj.matched.verwaltungGroupId },
                    verwaltungDraft: {
                        vwNewDisplayName: obj.vwNewDisplayName,
                        vwNewMailNick: obj.vwNewMailNick,
                        vwNewDescription: obj.vwNewDescription,
                        vwNewCreateTeam: obj.vwNewCreateTeam
                    }
                });
            }
        } catch {
            // ignore
        }
    }

    function loadState() {
        let rawLocal = null;
        try {
            rawLocal = localStorage.getItem(STORAGE_KEY);
        } catch {
            rawLocal = null;
        }
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const su = window.ms365AppDataV2.getSetup();
                const hasId = su && su.matched && !!su.matched.verwaltungGroupId;
                if (hasId || !rawLocal) {
                    const d = su.verwaltungDraft || {};
                    applyStateObject({
                        matched: su.matched,
                        vwNewDisplayName: d.vwNewDisplayName,
                        vwNewMailNick: d.vwNewMailNick,
                        vwNewDescription: d.vwNewDescription,
                        vwNewCreateTeam: d.vwNewCreateTeam
                    });
                    return;
                }
            }
        } catch {
            // ignore
        }
        try {
            if (!rawLocal) return;
            applyStateObject(JSON.parse(rawLocal));
        } catch {
            // ignore
        }
    }

    function clearStorage() {
        try {
            localStorage.removeItem(STORAGE_KEY);
            matchedGroupId = null;
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: { verwaltungGroupId: null }
                });
            }
            saveState();
            live().loadGroup({ silent: true });
            updateLeftListUi();
            toast('Zurückgesetzt.');
        } catch (e) {
            toast('Löschen fehlgeschlagen: ' + (e.message || e));
        }
    }

    async function onLogin() {
        const btn = document.getElementById('slgBtnLogin');
        if (btn) btn.disabled = true;
        try {
            await getGraphToken();
            toast('Angemeldet.');
            if (getActiveMatchedId()) await live().loadGroup({ silent: true });
        } catch (e) {
            toast('Anmeldung: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    function onClick(id, fn) {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', fn);
    }

    function wire() {
        live().bind({
            toast: toast,
            dlgConfirm: dlgConfirm,
            getGroupId: getActiveMatchedId,
            getActiveTab: function () {
                return activeTab;
            },
            ensureDirektionOwners: function (token, gid) {
                if (!(listCache.direktion && listCache.direktion.length)) {
                    throw new Error('Keine Direktion‑Adressen in den Schul‑Einstellungen.');
                }
                return ensureOwners(token, gid);
            },
            onUnmatched: function () {
                renderOwnerPreview();
                renderMemberPreview();
                updateLeftListUi();
            },
            onAfterLoad: function () {
                updateLeftListUi();
            }
        });
        live().wire();

        document.querySelectorAll('#slgDetailTabs .detail-tab-btn[data-slg-tab]').forEach(function (b) {
            b.addEventListener('click', function () {
                setTab(b.getAttribute('data-slg-tab') || 'general');
            });
        });

        onClick('slgBtnLogin', function () {
            onLogin();
        });
        onClick('slgBtnReloadLists', function () {
            readLists();
            updateLeftListUi();
            renderOwnerPreview();
            renderMemberPreview();
            toast('Listen neu eingelesen.');
        });
        onClick('slgBtnSearch', function () {
            runSearchGroups();
        });
        onClick('slgBtnCreate', function () {
            runCreateAndMatch();
        });
        onClick('slgBtnUnmatch', function () {
            runUnmatch();
        });
        onClick('slgBtnOpenEntra', function () {
            openEntraForMatched();
        });
        onClick('slgBtnSync', function () {
            runSyncMembers();
        });

        const groupSearch = document.getElementById('slgGroupSearch');
        if (groupSearch) {
            groupSearch.addEventListener('keydown', function (ev) {
                if (ev.key === 'Enter') {
                    ev.preventDefault();
                    runSearchGroups();
                }
            });
        }

        onClick('slgBtnSaveState', function () {
            saveState();
            toast('Gespeichert.');
        });
        onClick('slgBtnLoadState', function () {
            loadState();
            toast('Geladen.');
            if (getActiveMatchedId()) live().loadGroup({ silent: true });
        });
        onClick('slgBtnClearStorage', function () {
            clearStorage();
        });
    }

    function init() {
        readLists();
        loadState();
        updateLeftListUi();
        renderOwnerPreview();
        renderMemberPreview();
        wire();
        if (!getActiveMatchedId()) {
            live().setMatchedMode(false);
            applyCreateDefaults();
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
