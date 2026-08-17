(function () {
    'use strict';

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

    function dataV2() {
        return window.ms365AppDataV2 || null;
    }

    /** @type {'subject' | 'arge'} */
    let activeKind = 'subject';
    let activeCode = '';
    let listFilter = '';
    /** @type {{ code: string, name: string, subjects?: string[] }[]} */
    let catalog = { subject: [], arge: [] };
    /** @type {string[]} */
    let direktion = [];
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
    function normCode(v) {
        return normStr(v).toUpperCase();
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

    function getCatalogLink(kind, code) {
        const api = dataV2();
        if (api && typeof api.getCatalogLink === 'function') return api.getCatalogLink(kind, code);
        return null;
    }

    function upsertCatalogLink(entry) {
        const api = dataV2();
        if (api && typeof api.upsertCatalogLink === 'function') return api.upsertCatalogLink(entry);
        return null;
    }

    function rowsForKind(kind) {
        return kind === 'arge' ? catalog.arge : catalog.subject;
    }

    function getActiveRow() {
        const list = rowsForKind(activeKind);
        const code = normCode(activeCode);
        for (let i = 0; i < list.length; i++) {
            if (normCode(list[i].code) === code) return list[i];
        }
        return null;
    }

    function getActiveGroupId() {
        const link = getCatalogLink(activeKind, activeCode);
        const id = link && link.graphGroupId ? String(link.graphGroupId).trim() : '';
        return id || null;
    }

    function mailPrefix(kind) {
        const api = dataV2();
        const su = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
        const raw = kind === 'arge' ? (su && su.argeGroupMailPrefix) || 'ag' : (su && su.subjectGroupMailPrefix) || 'fach';
        if (api && typeof api.mailNicknamePrefixSanitize === 'function') {
            return api.mailNicknamePrefixSanitize(raw, 24) || (kind === 'arge' ? 'ag' : 'fach');
        }
        return kind === 'arge' ? 'ag' : 'fach';
    }

    function deriveNick(kind, code) {
        const pre = mailPrefix(kind);
        const tail = gug().sanitizeMailNickname(String(code || 'x')).slice(0, 40);
        return gug().sanitizeUnifiedGroupMailNickname(String(pre + tail).toLowerCase()).slice(0, 60);
    }

    function isDirektionRole(roleRaw) {
        const r = normStr(roleRaw).toLowerCase();
        return !!r && (r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1);
    }

    function readLists() {
        const settings = typeof window.ms365TenantSettingsLoad === 'function' ? window.ms365TenantSettingsLoad() : null;
        catalog.subject = Array.isArray(settings && settings.subjects) ? settings.subjects.slice() : [];
        catalog.arge = Array.isArray(settings && settings.arges) ? settings.arges.slice() : [];
        const out = [];
        const seen = new Set();
        const admin = settings && Array.isArray(settings.admin) ? settings.admin : [];
        admin.forEach(function (row) {
            if (!isDirektionRole(row && row.role)) return;
            const em = normEmail(row && row.email);
            if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        direktion = out;
    }

    function renderOwnerPreview() {
        const el = document.getElementById('slgOwnerPreview');
        if (!el) return;
        el.replaceChildren();
        if (!direktion.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent = 'Keine Direktion‑Owner in den Schul‑Einstellungen gefunden.';
            el.appendChild(p);
            return;
        }
        direktion.forEach(function (em) {
            const d = document.createElement('div');
            d.textContent = em;
            d.style.padding = '4px 0';
            d.style.borderBottom = '1px solid #eef1f4';
            el.appendChild(d);
        });
    }

    function renderMemberPreview() {
        const el = document.getElementById('slgMemberPreview');
        if (!el) return;
        el.replaceChildren();
        const row = getActiveRow();
        const p = document.createElement('p');
        p.style.margin = '0';
        p.style.color = '#6c757d';
        if (activeKind === 'arge' && row && Array.isArray(row.subjects) && row.subjects.length) {
            p.textContent = 'Zugeordnete Fächer: ' + row.subjects.join(', ') + '. Mitglieder nach dem Match in Graph pflegen.';
        } else {
            p.textContent = 'Keine Mitgliederliste in den Stammdaten. Nach dem Match Personen über Suche hinzufügen.';
        }
        el.appendChild(p);
    }

    function applyCreateDefaults() {
        const row = getActiveRow();
        const code = row ? row.code : activeCode;
        const name = row && row.name ? row.name : code;
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const desc = document.getElementById('slgNewDescription');
        const label = activeKind === 'arge' ? 'Arbeitsgruppe ' : 'Fach ';
        if (dn) dn.value = label + (name || code || '');
        if (nn) nn.value = deriveNick(activeKind, code);
        if (desc) {
            desc.value =
                label +
                (name || code || '') +
                (activeKind === 'arge' ? ' (MS365-Schulverwaltung / ARGE)' : ' (MS365-Schulverwaltung / Fachgruppe)');
        }
        const search = document.getElementById('slgGroupSearch');
        if (search && !normStr(search.value)) search.value = name || code || '';
    }

    function refreshMatchUi() {
        const gid = getActiveGroupId();
        const row = getActiveRow();
        const title = document.getElementById('slgDetailTitle');
        if (title) {
            const prefix = activeKind === 'arge' ? 'ARGE' : 'Fach';
            title.textContent = row ? prefix + ' ' + (row.name || row.code) : prefix;
        }
        live().resetCaches();
        live().setMatchedMode(!!gid);
        live().fillForm(gid ? { id: gid } : null);
        renderOwnerPreview();
        renderMemberPreview();
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
        const gid = getActiveGroupId();
        if (!gid) {
            if (activeTab === 'owners') renderOwnerPreview();
            if (activeTab === 'members') renderMemberPreview();
            return;
        }
        live().onTab(activeTab, gid);
    }

    function renderLeftList() {
        const host = document.getElementById('afgListItems');
        const summary = document.getElementById('afgListSummary');
        const empty = document.getElementById('afgEmptyHint');
        const wrap = document.getElementById('afgDetailWrap');
        if (!host) return;
        host.replaceChildren();
        const q = listFilter.toLowerCase();
        const all = rowsForKind(activeKind);
        const list = all.filter(function (row) {
            if (!q) return true;
            const hay = (row.code + ' ' + (row.name || '')).toLowerCase();
            return hay.indexOf(q) !== -1;
        });
        let matchedN = 0;
        all.forEach(function (row) {
            const link = getCatalogLink(activeKind, row.code);
            if (link && link.graphGroupId) matchedN++;
        });
        if (summary) {
            summary.textContent =
                String(all.length) +
                (activeKind === 'arge' ? ' ARGE' : ' Fächer') +
                ' · ' +
                String(matchedN) +
                ' gematcht' +
                (q ? ' · Filter: ' + String(list.length) : '');
        }
        const hasRows = all.length > 0;
        if (empty) empty.style.display = hasRows ? 'none' : '';
        if (wrap) wrap.style.display = hasRows ? '' : 'none';

        if (!list.length) {
            const li = document.createElement('li');
            const p = document.createElement('p');
            p.className = 'muted';
            p.style.margin = '10px 12px';
            p.textContent = hasRows ? 'Keine Treffer im Filter.' : 'Liste ist leer.';
            li.appendChild(p);
            host.appendChild(li);
            return;
        }

        list.forEach(function (row) {
            const link = getCatalogLink(activeKind, row.code);
            const gid = link && link.graphGroupId ? String(link.graphGroupId) : '';
            const li = document.createElement('li');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.setAttribute('data-afg-code', row.code);
            if (normCode(row.code) === normCode(activeCode)) btn.setAttribute('aria-current', 'true');
            const main = document.createElement('span');
            main.className = 'slg-side-main';
            const t = document.createElement('span');
            t.className = 'slg-side-title';
            t.textContent = (row.name || row.code) + (row.name && row.code ? ' (' + row.code + ')' : '');
            const meta = document.createElement('span');
            meta.className = 'muted slg-side-meta';
            meta.textContent = gid ? 'Gematcht: ' + gid : 'Noch kein Match';
            main.appendChild(t);
            main.appendChild(meta);
            btn.appendChild(main);
            li.appendChild(btn);
            host.appendChild(li);
        });
    }

    function ensureActiveCode() {
        const list = rowsForKind(activeKind);
        if (!list.length) {
            activeCode = '';
            return;
        }
        const has = list.some(function (r) {
            return normCode(r.code) === normCode(activeCode);
        });
        if (!has) activeCode = list[0].code;
    }

    function setActiveKind(kind) {
        activeKind = kind === 'arge' ? 'arge' : 'subject';
        document.querySelectorAll('[data-afg-kind]').forEach(function (b) {
            b.setAttribute('aria-pressed', b.getAttribute('data-afg-kind') === activeKind ? 'true' : 'false');
        });
        ensureActiveCode();
        const search = document.getElementById('slgGroupSearch');
        if (search) search.value = '';
        const host = document.getElementById('slgGroupSearchResults');
        if (host) {
            host.replaceChildren();
            host.style.display = 'none';
        }
        renderLeftList();
        applyCreateDefaults();
        setTab('general');
        refreshMatchUi();
    }

    function setActiveCode(code) {
        activeCode = normCode(code);
        const search = document.getElementById('slgGroupSearch');
        if (search) search.value = '';
        const host = document.getElementById('slgGroupSearchResults');
        if (host) {
            host.replaceChildren();
            host.style.display = 'none';
        }
        renderLeftList();
        applyCreateDefaults();
        setTab('general');
        refreshMatchUi();
        if (getActiveGroupId()) live().loadGroup({ silent: true });
    }

    function persistMatch(g, mode) {
        const row = getActiveRow();
        upsertCatalogLink({
            kind: activeKind,
            code: activeCode,
            graphGroupId: g && g.id ? String(g.id) : '',
            displayName: (g && g.displayName) || (row && row.name) || '',
            mailNickname: (g && g.mailNickname) || '',
            mode: mode
        });
        renderLeftList();
    }

    function renderGroupSearchResults(list) {
        const host = document.getElementById('slgGroupSearchResults');
        if (!host) return;
        host.replaceChildren();
        host.style.display = 'block';
        if (!list || !list.length) {
            const p = document.createElement('div');
            p.className = 'muted';
            p.textContent = 'Keine passenden Microsoft 365‑Gruppen (Unified) gefunden.';
            host.appendChild(p);
            return;
        }
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
                if (!g || !g.id || !activeCode) return;
                persistMatch(g, 'matched');
                live().fillForm(g);
                live().setMatchedMode(true);
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
        if (!activeCode) {
            toast('Bitte zuerst einen Katalogeintrag wählen.');
            return;
        }
        const inp = document.getElementById('slgGroupSearch');
        const q = inp && inp.value ? inp.value.trim() : '';
        if (!q) {
            toast('Bitte einen Suchbegriff eingeben.');
            return;
        }
        try {
            const token = await gug().getGraphToken();
            const list = await gug().searchUnifiedGroups(token, q);
            renderGroupSearchResults(list);
            if (!list.length) toast('Keine passenden Gruppen gefunden.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function runCreateAndMatch() {
        if (!activeCode) {
            toast('Bitte zuerst einen Katalogeintrag wählen.');
            return;
        }
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
            const token = await gug().getGraphToken();
            const g = await gug().createUnifiedGroup(token, displayName, mailNick, desc);
            await gug().ensureOwners(token, g.id, direktion || []);
            if (ct && ct.checked) {
                toast('Gruppe angelegt – Team wird bereitgestellt …');
                await gug().provisionTeamForGroup(token, g.id);
            }
            persistMatch(g, 'created');
            live().fillForm(g);
            live().setMatchedMode(true);
            await live().loadGroup({ silent: true });
            toast('Gruppe angelegt und gematcht.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    function runUnmatch() {
        if (!getActiveGroupId()) return;
        const api = dataV2();
        if (api && typeof api.clearCatalogLinkGroup === 'function') {
            api.clearCatalogLinkGroup(activeKind, activeCode);
        } else {
            persistMatch({ id: '', displayName: '', mailNickname: '' }, '');
        }
        renderLeftList();
        live().loadGroup({ silent: true });
        toast('Match gelöst.');
    }

    function openEntraForMatched() {
        const gid = getActiveGroupId();
        if (!gid) return;
        window.open(
            'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Members/groupId/' +
                encodeURIComponent(gid),
            '_blank',
            'noopener'
        );
    }

    async function onLogin() {
        const btn = document.getElementById('slgBtnLogin');
        if (btn) btn.disabled = true;
        try {
            await gug().getGraphToken();
            toast('Angemeldet.');
            if (getActiveGroupId()) await live().loadGroup({ silent: true });
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
            getGroupId: getActiveGroupId,
            getActiveTab: function () {
                return activeTab;
            },
            ensureDirektionOwners: function (token, gid) {
                if (!direktion.length) throw new Error('Keine Direktion‑Adressen in den Schul‑Einstellungen.');
                return gug().ensureOwners(token, gid, direktion);
            },
            onUnmatched: function () {
                renderOwnerPreview();
                renderMemberPreview();
                renderLeftList();
            },
            onAfterLoad: function () {
                renderLeftList();
            }
        });
        live().wire();

        document.querySelectorAll('[data-afg-kind]').forEach(function (b) {
            b.addEventListener('click', function () {
                setActiveKind(b.getAttribute('data-afg-kind') === 'arge' ? 'arge' : 'subject');
                if (getActiveGroupId()) live().loadGroup({ silent: true });
            });
        });
        const listHost = document.getElementById('afgListItems');
        if (listHost) {
            listHost.addEventListener('click', function (ev) {
                const t = ev.target;
                const item = t && t.closest ? t.closest('button[data-afg-code]') : null;
                if (!item) return;
                setActiveCode(item.getAttribute('data-afg-code') || '');
            });
        }
        const filter = document.getElementById('afgListFilter');
        if (filter) {
            filter.addEventListener('input', function () {
                listFilter = String(filter.value || '').trim();
                renderLeftList();
            });
        }
        document.querySelectorAll('#slgDetailTabs .detail-tab-btn[data-slg-tab]').forEach(function (b) {
            b.addEventListener('click', function () {
                setTab(b.getAttribute('data-slg-tab') || 'general');
            });
        });
        onClick('slgBtnLogin', onLogin);
        onClick('slgBtnReloadLists', function () {
            readLists();
            ensureActiveCode();
            renderLeftList();
            applyCreateDefaults();
            refreshMatchUi();
            toast('Listen neu eingelesen.');
        });
        onClick('slgBtnSearch', runSearchGroups);
        onClick('slgBtnCreate', runCreateAndMatch);
        onClick('slgBtnUnmatch', runUnmatch);
        onClick('slgBtnOpenEntra', openEntraForMatched);
        const groupSearch = document.getElementById('slgGroupSearch');
        if (groupSearch) {
            groupSearch.addEventListener('keydown', function (ev) {
                if (ev.key === 'Enter') {
                    ev.preventDefault();
                    runSearchGroups();
                }
            });
        }
    }

    function init() {
        readLists();
        ensureActiveCode();
        wire();
        setActiveKind('subject');
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
