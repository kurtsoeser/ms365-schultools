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

    let activeKey = '';
    let listFilter = '';
    /** @type {{ code: string, name: string, year: string, headName?: string, headEmail?: string, stableMailNickname?: string }[]} */
    let classes = [];
    /** @type {{ klasse: string, name: string, email: string }[]} */
    let students = [];
    /** @type {string[]} */
    let direktion = [];
    let schoolYearLabel = '';
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

    function sanitizeNick(raw) {
        return String(raw || '')
            .trim()
            .replace(/[^a-zA-Z0-9]/g, '')
            .toLowerCase()
            .slice(0, 60);
    }

    function rowKey(row) {
        return normCode(row && row.code) || normStr(row && row.name).toUpperCase();
    }

    function getActiveRow() {
        const key = normCode(activeKey) || normStr(activeKey).toUpperCase();
        for (let i = 0; i < classes.length; i++) {
            if (rowKey(classes[i]) === key) return classes[i];
        }
        return null;
    }

    function listClassTeams() {
        const api = dataV2();
        if (!api || typeof api.getContainer !== 'function') return [];
        const c = api.getContainer();
        const raw = c && c.core && Array.isArray(c.core.classTeams) ? c.core.classTeams : [];
        if (typeof api.normalizeCoreClassTeams === 'function') return api.normalizeCoreClassTeams(raw);
        return raw;
    }

    function deriveNick(row) {
        if (!row) return '';
        const fromRow = sanitizeNick(row.stableMailNickname);
        if (fromRow) return fromRow;
        if (typeof window.ms365DeriveClassStableMailNickname === 'function') {
            const d = sanitizeNick(window.ms365DeriveClassStableMailNickname(row.year || '', row.code || ''));
            if (d) return d;
        }
        const y = normStr(row.year);
        const yy = /^\d{4}$/.test(y) ? y : '';
        const tail = String(normCode(row.code) || '')
            .replace(/[^0-9A-Za-z]/g, '')
            .toLowerCase()
            .slice(0, 24);
        if (yy && tail) return ('jg' + yy + tail).toLowerCase().slice(0, 60);
        if (tail) return ('jg' + tail).toLowerCase().slice(0, 60);
        return '';
    }

    function findClassTeam(row) {
        if (!row) return null;
        const teams = listClassTeams();
        const nick = deriveNick(row);
        if (nick) {
            for (let i = 0; i < teams.length; i++) {
                if (sanitizeNick(teams[i].stableMailNickname) === nick) return teams[i];
            }
        }
        const code = normCode(row.code);
        if (code) {
            for (let i = 0; i < teams.length; i++) {
                if (normCode(teams[i].classCode) === code) return teams[i];
            }
        }
        return null;
    }

    function persistNickForRow(row) {
        const existing = findClassTeam(row);
        if (existing && existing.stableMailNickname) return sanitizeNick(existing.stableMailNickname);
        return deriveNick(row);
    }

    function getActiveGroupId() {
        const row = getActiveRow();
        const team = findClassTeam(row);
        const id = team && team.graphGroupId ? String(team.graphGroupId).trim() : '';
        return id || null;
    }

    function isDirektionRole(roleRaw) {
        const r = normStr(roleRaw).toLowerCase();
        return !!r && (r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1);
    }

    function currentYearFromV2() {
        const api = dataV2();
        if (!api || typeof api.getContainer !== 'function') return '';
        const c = api.getContainer();
        return c && c.years ? String(c.years.current || '').trim() : '';
    }

    function readLists() {
        const settings = typeof window.ms365TenantSettingsLoad === 'function' ? window.ms365TenantSettingsLoad() : null;
        classes = Array.isArray(settings && settings.classes) ? settings.classes.slice() : [];
        students = Array.isArray(settings && settings.students) ? settings.students.slice() : [];
        schoolYearLabel = currentYearFromV2();
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
        const yearEl = document.getElementById('jgYearLabel');
        if (yearEl) {
            yearEl.textContent = schoolYearLabel
                ? 'Schuljahr ' + schoolYearLabel + ' – links Klasse wählen, rechts matchen oder anlegen.'
                : 'Aktuelles Schuljahr – links Klasse wählen, rechts matchen oder anlegen.';
        }
    }

    function studentsForClass(row) {
        if (!row) return [];
        const code = normCode(row.code);
        const name = normStr(row.name).toLowerCase();
        return students.filter(function (s) {
            const k = normStr(s && s.klasse);
            if (!k) return false;
            if (code && normCode(k) === code) return true;
            if (name && k.toLowerCase() === name) return true;
            return false;
        });
    }

    function emailsForClass(row) {
        const seen = new Set();
        const out = [];
        studentsForClass(row).forEach(function (s) {
            const em = normEmail(s && s.email);
            if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
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
        const list = studentsForClass(row);
        const emails = emailsForClass(row);
        if (!list.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent =
                'Keine Schüler:innen dieser Klasse in den Stammdaten. Nach dem Match Personen über Suche hinzufügen.';
            el.appendChild(p);
            return;
        }
        const first = list.slice(0, 30);
        first.forEach(function (s) {
            const d = document.createElement('div');
            const parts = [];
            if (s.name) parts.push(s.name);
            if (s.email) parts.push(s.email);
            if (!parts.length) parts.push(s.klasse || '–');
            d.textContent = parts.join(' · ');
            d.style.padding = '4px 0';
            d.style.borderBottom = '1px solid #eef1f4';
            el.appendChild(d);
        });
        const more = document.createElement('div');
        more.className = 'muted';
        more.style.paddingTop = '8px';
        more.textContent =
            String(list.length) +
            ' Einträge · ' +
            String(emails.length) +
            ' mit E‑Mail' +
            (list.length > first.length ? ' · Anzeige der ersten 30' : '') +
            '.';
        el.appendChild(more);
    }

    function applyCreateDefaults() {
        const row = getActiveRow();
        const code = row ? row.code : activeKey;
        const name = row && row.name ? row.name : code;
        const nick = persistNickForRow(row);
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const desc = document.getElementById('slgNewDescription');
        if (dn) dn.value = name || ('Klasse ' + (code || ''));
        if (nn) nn.value = nick;
        if (desc) {
            desc.value =
                'Jahrgangsgruppe ' +
                (name || code || '') +
                (row && row.year ? ' / Abschluss ' + row.year : '') +
                ' (MS365-Schulverwaltung)';
        }
        const search = document.getElementById('slgGroupSearch');
        if (search) search.value = nick || name || code || '';
    }

    function refreshMatchUi() {
        const gid = getActiveGroupId();
        const row = getActiveRow();
        const title = document.getElementById('slgDetailTitle');
        if (title) {
            if (!row) title.textContent = 'Jahrgangsgruppe';
            else {
                const bits = [];
                bits.push(row.name || row.code || 'Klasse');
                if (row.code && row.name && row.name !== row.code) bits[0] = row.name + ' (' + row.code + ')';
                title.textContent = bits[0];
            }
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
        const host = document.getElementById('jgListItems');
        const summary = document.getElementById('jgListSummary');
        const empty = document.getElementById('jgEmptyHint');
        const wrap = document.getElementById('jgDetailWrap');
        if (!host) return;
        host.replaceChildren();
        const q = listFilter.toLowerCase();
        const all = classes;
        const list = all.filter(function (row) {
            if (!q) return true;
            const nick = persistNickForRow(row);
            const hay = (row.code + ' ' + (row.name || '') + ' ' + (row.year || '') + ' ' + nick).toLowerCase();
            return hay.indexOf(q) !== -1;
        });
        let matchedN = 0;
        all.forEach(function (row) {
            const team = findClassTeam(row);
            if (team && team.graphGroupId) matchedN++;
        });
        if (summary) {
            summary.textContent =
                String(all.length) +
                ' Klassen' +
                (schoolYearLabel ? ' (' + schoolYearLabel + ')' : '') +
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
            const team = findClassTeam(row);
            const gid = team && team.graphGroupId ? String(team.graphGroupId) : '';
            const nick = persistNickForRow(row);
            const li = document.createElement('li');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.setAttribute('data-jg-code', rowKey(row));
            if (rowKey(row) === (normCode(activeKey) || normStr(activeKey).toUpperCase())) {
                btn.setAttribute('aria-current', 'true');
            }
            const main = document.createElement('span');
            main.className = 'slg-side-main';
            const t = document.createElement('span');
            t.className = 'slg-side-title';
            t.textContent = (row.name || row.code) + (row.name && row.code && row.name !== row.code ? ' (' + row.code + ')' : '');
            const meta = document.createElement('span');
            meta.className = 'muted slg-side-meta';
            const bits = [];
            if (row.year) bits.push('Abschluss ' + row.year);
            if (nick) bits.push(nick);
            bits.push(gid ? 'Gematcht' : 'Noch kein Match');
            meta.textContent = bits.join(' · ');
            main.appendChild(t);
            main.appendChild(meta);
            btn.appendChild(main);
            li.appendChild(btn);
            host.appendChild(li);
        });
    }

    function ensureActiveKey() {
        if (!classes.length) {
            activeKey = '';
            return;
        }
        const has = classes.some(function (r) {
            return rowKey(r) === (normCode(activeKey) || normStr(activeKey).toUpperCase());
        });
        if (!has) activeKey = rowKey(classes[0]);
    }

    function setActiveKey(code) {
        activeKey = normCode(code) || normStr(code).toUpperCase();
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
        const api = dataV2();
        if (!api || typeof api.upsertClassTeam !== 'function') {
            toast('classTeams-Speicher (app-data-v2) nicht verfügbar.');
            return;
        }
        const nick = persistNickForRow(row);
        if (!nick) {
            toast('Kein gültiger Mail‑Nickname für diese Klasse (Kürzel und Abschlussjahr prüfen).');
            return;
        }
        api.upsertClassTeam({
            stableMailNickname: nick,
            graphGroupId: g && g.id ? String(g.id) : '',
            classCode: row && row.code ? row.code : activeKey,
            displayName: (g && g.displayName) || (row && row.name) || '',
            abschlussJahr: row && row.year ? row.year : '',
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
                if (!g || !g.id || !activeKey) return;
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
        if (!activeKey) {
            toast('Bitte zuerst eine Klasse wählen.');
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
        if (!activeKey) {
            toast('Bitte zuerst eine Klasse wählen.');
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
            const emails = emailsForClass(getActiveRow());
            if (emails.length) {
                await gug().syncEmailsToGroup(token, g.id, emails, 'Klasse', function () {});
            }
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
        persistMatch({ id: '', displayName: '', mailNickname: '' }, '');
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

    async function runSyncMembers() {
        const gid = getActiveGroupId();
        if (!gid) {
            toast('Zuerst eine Gruppe matchen oder anlegen.');
            return;
        }
        const emails = emailsForClass(getActiveRow());
        if (!emails.length) {
            toast('Keine Schüler‑E‑Mails für diese Klasse in den Schul‑Einstellungen.');
            return;
        }
        clearSyncLog();
        appendSyncLog('Start: Klasse (' + emails.length + ' Adressen) …', '');
        try {
            const token = await gug().getGraphToken();
            const r = await gug().syncEmailsToGroup(token, gid, emails, 'Klasse', appendSyncLog);
            appendSyncLog('Fertig: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            if (direktion.length) await gug().ensureOwners(token, gid, direktion);
            live().invalidateMembership();
            await live().loadMembers();
            toast('Synchronisation abgeschlossen.');
        } catch (e) {
            appendSyncLog('Abbruch: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        }
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

        const listHost = document.getElementById('jgListItems');
        if (listHost) {
            listHost.addEventListener('click', function (ev) {
                const t = ev.target;
                const item = t && t.closest ? t.closest('button[data-jg-code]') : null;
                if (!item) return;
                setActiveKey(item.getAttribute('data-jg-code') || '');
            });
        }
        const filter = document.getElementById('jgListFilter');
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
            ensureActiveKey();
            renderLeftList();
            applyCreateDefaults();
            refreshMatchUi();
            toast('Listen neu eingelesen.');
        });
        onClick('slgBtnSearch', runSearchGroups);
        onClick('slgBtnCreate', runCreateAndMatch);
        onClick('slgBtnUnmatch', runUnmatch);
        onClick('slgBtnOpenEntra', openEntraForMatched);
        onClick('slgBtnSync', runSyncMembers);
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
        ensureActiveKey();
        wire();
        renderLeftList();
        applyCreateDefaults();
        setTab('general');
        refreshMatchUi();
        if (getActiveGroupId()) live().loadGroup({ silent: true });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
