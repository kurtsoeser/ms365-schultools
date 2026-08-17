(function () {
    'use strict';

    /** @type {null|{ toast: Function, dlgConfirm: Function, getGroupId: Function, getActiveTab: Function, ensureDirektionOwners: Function, onUnmatched: Function, onAfterLoad: Function }} */
    let ctx = null;
    /** @type {object|null} */
    let liveGroup = null;
    /** @type {object[]} */
    let ownersCache = [];
    let membersCountCache = -1;
    let ownersLoadedForId = '';
    let membersLoadedForId = '';

    function gug() {
        const G = window.ms365GraphUnifiedGroups;
        if (!G) throw new Error('graph-unified-groups.js muss vor diesem Skript geladen werden.');
        return G;
    }

    function toast(msg) {
        if (ctx && ctx.toast) ctx.toast(msg);
    }

    function dlgConfirm(message, options) {
        if (ctx && ctx.dlgConfirm) return ctx.dlgConfirm(message, options);
        return Promise.resolve(window.confirm(message));
    }

    function getGroupId() {
        return ctx && ctx.getGroupId ? ctx.getGroupId() : null;
    }

    function showEl(id, on) {
        const el = document.getElementById(id);
        if (el) el.style.display = on ? '' : 'none';
    }

    function formatDateTimeAT(iso) {
        const s = String(iso || '').trim();
        if (!s) return '';
        const d = new Date(s);
        if (isNaN(d.getTime())) return s;
        try {
            return d.toLocaleString('de-AT', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit'
            });
        } catch {
            return s;
        }
    }

    function fillForm(group) {
        liveGroup = group || null;
        const name = document.getElementById('slgLiveName');
        const art = document.getElementById('slgLiveArt');
        const mail = document.getElementById('slgLiveMail');
        const vis = document.getElementById('slgLiveVisibility');
        const exp = document.getElementById('slgLiveExpires');
        const alias = document.getElementById('slgLiveAlias');
        const desc = document.getElementById('slgLiveDescription');
        const idEl = document.getElementById('slgLiveId');
        if (!group) {
            if (name) name.value = '';
            if (art) art.value = '';
            if (mail) mail.value = '';
            if (vis) vis.value = 'Private';
            if (exp) exp.value = '';
            if (alias) alias.value = '';
            if (desc) desc.value = '';
            if (idEl) idEl.value = '';
            return;
        }
        if (name) name.value = String(group.displayName || '');
        if (art) art.value = gug().groupArtLabel(group);
        if (mail) mail.value = String(group.mail || '');
        if (vis) {
            const v = String(group.visibility || '').trim();
            vis.value = v === 'Public' ? 'Public' : 'Private';
        }
        if (exp) exp.value = formatDateTimeAT(group.expirationDateTime);
        if (alias) alias.value = String(group.mailNickname || '');
        if (desc) desc.value = String(group.description || '');
        if (idEl) idEl.value = String(group.id || '');
    }

    function setMatchedMode(has) {
        showEl('slgUnmatchedPanel', !has);
        showEl('slgMatchedPanel', has);
        showEl('slgOwnersUnmatched', !has);
        showEl('slgOwnersMatched', has);
        showEl('slgMembersUnmatched', !has);
        showEl('slgMembersMatched', has);
        const btnOpen = document.getElementById('slgBtnOpenEntra');
        const btnUn = document.getElementById('slgBtnUnmatch');
        if (btnOpen) btnOpen.disabled = !has;
        if (btnUn) btnUn.disabled = !has;
        const sub = document.getElementById('slgDetailSubtitle');
        if (!sub) return;
        if (!has) {
            sub.textContent = 'Gruppe matchen oder anlegen';
        } else if (liveGroup && liveGroup.displayName) {
            sub.textContent = liveGroup.displayName;
        } else {
            sub.textContent = 'Gematcht – anmelden oder „Neu laden“, um Details zu holen';
        }
    }

    function invalidateMembership() {
        ownersCache = [];
        membersCountCache = -1;
        ownersLoadedForId = '';
        membersLoadedForId = '';
    }

    function resetCaches() {
        invalidateMembership();
        liveGroup = null;
    }

    function fillUserSearchSelect(selId, users) {
        const sel = document.getElementById(selId);
        if (!sel) return;
        sel.replaceChildren();
        if (!users || !users.length) {
            const opt = document.createElement('option');
            opt.value = '';
            opt.textContent = '(keine Treffer)';
            sel.appendChild(opt);
            return;
        }
        users.forEach(function (u) {
            const opt = document.createElement('option');
            opt.value = u.id || '';
            opt.textContent = gug().personLabel(u) || (u.id ? String(u.id) : '');
            sel.appendChild(opt);
        });
    }

    function renderPersonRows(wrap, list, emptyText, removeAttr) {
        if (!wrap) return;
        wrap.replaceChildren();
        const items = Array.isArray(list) ? list : [];
        if (!items.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent = emptyText;
            wrap.appendChild(p);
            return;
        }
        items.forEach(function (person) {
            const row = document.createElement('div');
            row.style.display = 'flex';
            row.style.justifyContent = 'space-between';
            row.style.alignItems = 'flex-start';
            row.style.gap = '10px';
            row.style.padding = '8px 0';
            row.style.borderBottom = '1px solid #e9ecef';
            const txt = document.createElement('div');
            txt.style.whiteSpace = 'pre-wrap';
            txt.style.lineHeight = '1.35';
            txt.style.fontSize = '0.92em';
            txt.textContent = gug().personLabel(person) || '–';
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn';
            btn.style.padding = '6px 10px';
            btn.style.fontSize = '0.85em';
            btn.textContent = 'Entfernen';
            btn.setAttribute(removeAttr, person.id || '');
            row.appendChild(txt);
            row.appendChild(btn);
            wrap.appendChild(row);
        });
    }

    function renderOwnersList(owners) {
        renderPersonRows(
            document.getElementById('slgOwnersList'),
            owners,
            'Keine Besitzer gefunden (Achtung: das ist meist ein Problem).',
            'data-slg-remove-owner'
        );
    }

    function renderMembersList(result, totalCount) {
        const wrap = document.getElementById('slgMembersList');
        if (!wrap) return;
        wrap.replaceChildren();
        const list = result && Array.isArray(result.items) ? result.items : [];
        const truncated = !!(result && result.truncated);
        const head = document.createElement('div');
        head.style.display = 'flex';
        head.style.justifyContent = 'space-between';
        head.style.alignItems = 'baseline';
        head.style.gap = '10px';
        head.style.marginBottom = '8px';
        const left = document.createElement('div');
        left.style.fontWeight = '900';
        left.style.color = '#32325d';
        left.textContent =
            'Mitglieder: ' + (totalCount >= 0 ? String(totalCount) : String(list.length)) + (truncated ? ' (Anzeige gekürzt)' : '');
        const right = document.createElement('div');
        right.className = 'muted';
        right.style.fontWeight = '700';
        right.textContent = truncated ? 'gekürzt' : 'vollständig';
        head.appendChild(left);
        head.appendChild(right);
        wrap.appendChild(head);
        if (!list.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent = 'Keine Mitglieder.';
            wrap.appendChild(p);
            return;
        }
        list.forEach(function (m) {
            const row = document.createElement('div');
            row.style.display = 'flex';
            row.style.justifyContent = 'space-between';
            row.style.alignItems = 'flex-start';
            row.style.gap = '10px';
            row.style.padding = '8px 0';
            row.style.borderBottom = '1px solid #e9ecef';
            const txt = document.createElement('div');
            txt.style.whiteSpace = 'pre-wrap';
            txt.style.lineHeight = '1.35';
            txt.style.fontSize = '0.92em';
            txt.textContent = gug().personLabel(m) || '–';
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn';
            btn.style.padding = '6px 10px';
            btn.style.fontSize = '0.85em';
            btn.textContent = 'Entfernen';
            btn.setAttribute('data-slg-remove-member', m.id || '');
            row.appendChild(txt);
            row.appendChild(btn);
            wrap.appendChild(row);
        });
    }

    function setBusy(ids, on) {
        ids.forEach(function (id) {
            const el = document.getElementById(id);
            if (el) el.disabled = !!on;
        });
    }

    async function loadOwnersNow() {
        const gid = getGroupId();
        if (!gid) return;
        setBusy(['slgOwnersReloadBtn', 'slgOwnerAddBtn', 'slgOwnerSearchBtn', 'slgOwnersEnsureDirektionBtn'], true);
        try {
            const token = await gug().getGraphToken();
            ownersCache = await gug().fetchGroupOwners(token, gid);
            ownersLoadedForId = gid;
            renderOwnersList(ownersCache);
        } catch (e) {
            toast('Owner laden: ' + (e.message || e));
        } finally {
            setBusy(['slgOwnersReloadBtn', 'slgOwnerAddBtn', 'slgOwnerSearchBtn', 'slgOwnersEnsureDirektionBtn'], false);
        }
    }

    async function loadMembersNow() {
        const gid = getGroupId();
        if (!gid) return;
        setBusy(['slgMembersReloadBtn', 'slgMemberAddBtn', 'slgMemberSearchBtn', 'slgBtnSync'], true);
        try {
            const token = await gug().getGraphToken();
            try {
                membersCountCache = await gug().fetchGroupMemberCount(token, gid);
            } catch {
                membersCountCache = -1;
            }
            const result = await gug().fetchGroupMembers(token, gid);
            membersLoadedForId = gid;
            renderMembersList(result, membersCountCache);
        } catch (e) {
            toast('Mitglieder laden: ' + (e.message || e));
        } finally {
            setBusy(['slgMembersReloadBtn', 'slgMemberAddBtn', 'slgMemberSearchBtn', 'slgBtnSync'], false);
        }
    }

    async function loadGroup(opts) {
        const silent = !!(opts && opts.silent);
        const gid = getGroupId();
        if (!gid) {
            resetCaches();
            fillForm(null);
            setMatchedMode(false);
            if (ctx && ctx.onUnmatched) ctx.onUnmatched();
            return;
        }
        setMatchedMode(true);
        try {
            const token = await gug().getGraphToken();
            const g = await gug().fetchGroup(token, gid);
            fillForm(g);
            setMatchedMode(true);
            if (ctx && ctx.onAfterLoad) ctx.onAfterLoad();
            const tab = ctx && ctx.getActiveTab ? ctx.getActiveTab() : 'general';
            if (tab === 'owners') await loadOwnersNow();
            else if (tab === 'members') await loadMembersNow();
            if (!silent) toast('Gruppe geladen.');
        } catch (e) {
            fillForm({ id: gid, displayName: '(nicht geladen)' });
            toast('Gruppe laden: ' + (e.message || e));
        }
    }

    function onTab(tab, gid) {
        if (!gid) return;
        if (tab === 'owners' && ownersLoadedForId !== gid) loadOwnersNow();
        if (tab === 'members' && membersLoadedForId !== gid) loadMembersNow();
    }

    function bindEnter(id, fn) {
        const el = document.getElementById(id);
        if (!el) return;
        el.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                fn();
            }
        });
    }

    function onClick(id, fn) {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', fn);
    }

    async function runUpdate() {
        const gid = getGroupId();
        if (!gid) return;
        const nameEl = document.getElementById('slgLiveName');
        const descEl = document.getElementById('slgLiveDescription');
        const visEl = document.getElementById('slgLiveVisibility');
        const displayName = nameEl ? String(nameEl.value || '').trim() : '';
        if (!displayName) {
            toast('Bitte einen Anzeigenamen eingeben.');
            return;
        }
        try {
            const token = await gug().getGraphToken();
            await gug().patchGroup(token, gid, {
                displayName: displayName,
                description: descEl ? descEl.value : '',
                visibility: visEl ? visEl.value : 'Private'
            });
            await loadGroup({ silent: true });
            toast('Gruppe aktualisiert.');
        } catch (e) {
            toast('Update: ' + (e.message || e));
        }
    }

    async function runSearch(kind) {
        const inp = document.getElementById(kind === 'owner' ? 'slgOwnerSearch' : 'slgMemberSearch');
        const q = inp ? String(inp.value || '').trim() : '';
        if (!q) {
            toast('Bitte einen Suchbegriff eingeben.');
            return;
        }
        const btnId = kind === 'owner' ? 'slgOwnerSearchBtn' : 'slgMemberSearchBtn';
        const selId = kind === 'owner' ? 'slgOwnerSearchResults' : 'slgMemberSearchResults';
        const btn = document.getElementById(btnId);
        if (btn) btn.disabled = true;
        try {
            const token = await gug().getGraphToken();
            const users = await gug().searchUsers(token, q);
            fillUserSearchSelect(selId, users);
            toast('Suche: ' + users.length + ' Treffer.');
        } catch (e) {
            toast('Suche: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    async function runAddOwner() {
        const gid = getGroupId();
        if (!gid) return;
        const sel = document.getElementById('slgOwnerSearchResults');
        const userId = sel && sel.value ? String(sel.value).trim() : '';
        if (!userId) {
            toast('Bitte zuerst einen Benutzer aus den Treffern auswählen.');
            return;
        }
        const btn = document.getElementById('slgOwnerAddBtn');
        if (btn) btn.disabled = true;
        try {
            const token = await gug().getGraphToken();
            await gug().addOwnerWithMemberFallback(token, gid, userId);
            await loadOwnersNow();
            toast('Owner hinzugefügt.');
        } catch (e) {
            toast('Owner hinzufügen: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    async function runEnsureDirektion() {
        const gid = getGroupId();
        if (!gid || !ctx || typeof ctx.ensureDirektionOwners !== 'function') return;
        try {
            const token = await gug().getGraphToken();
            await ctx.ensureDirektionOwners(token, gid);
            await loadOwnersNow();
            toast('Direktion als Owner gesetzt.');
        } catch (e) {
            toast('Direktion setzen: ' + (e.message || e));
        }
    }

    async function runRemoveOwner(ownerId) {
        const gid = getGroupId();
        if (!gid || !ownerId) return;
        if (ownersCache.length <= 1) {
            toast('Der letzte Besitzer kann nicht entfernt werden.');
            return;
        }
        if (!(await dlgConfirm('Diesen Owner wirklich entfernen?', { title: 'Owner', okText: 'Entfernen', danger: true }))) {
            return;
        }
        try {
            const token = await gug().getGraphToken();
            await gug().removeGroupOwner(token, gid, ownerId);
            await loadOwnersNow();
            toast('Owner entfernt.');
        } catch (e) {
            toast('Owner entfernen: ' + (e.message || e));
        }
    }

    async function runAddMember() {
        const gid = getGroupId();
        if (!gid) return;
        const sel = document.getElementById('slgMemberSearchResults');
        const userId = sel && sel.value ? String(sel.value).trim() : '';
        if (!userId) {
            toast('Bitte zuerst einen Benutzer aus den Treffern auswählen.');
            return;
        }
        const btn = document.getElementById('slgMemberAddBtn');
        if (btn) btn.disabled = true;
        try {
            const token = await gug().getGraphToken();
            try {
                await gug().graphAddMember(token, gid, userId);
            } catch (e) {
                if (!gug().isDuplicateMemberError(e)) throw e;
            }
            membersLoadedForId = '';
            await loadMembersNow();
            toast('Mitglied hinzugefügt.');
        } catch (e) {
            toast('Mitglied hinzufügen: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    async function runRemoveMember(memberId) {
        const gid = getGroupId();
        if (!gid || !memberId) return;
        if (!(await dlgConfirm('Dieses Mitglied wirklich entfernen?', { title: 'Mitglied', okText: 'Entfernen', danger: true }))) {
            return;
        }
        try {
            const token = await gug().getGraphToken();
            await gug().removeGroupMember(token, gid, memberId);
            membersLoadedForId = '';
            await loadMembersNow();
            toast('Mitglied entfernt.');
        } catch (e) {
            toast('Mitglied entfernen: ' + (e.message || e));
        }
    }

    function wire() {
        onClick('slgBtnUpdateGroup', function () {
            runUpdate();
        });
        onClick('slgBtnRefreshGroup', function () {
            loadGroup({ silent: false });
        });
        onClick('slgOwnerSearchBtn', function () {
            runSearch('owner');
        });
        onClick('slgOwnerAddBtn', function () {
            runAddOwner();
        });
        onClick('slgOwnersReloadBtn', function () {
            loadOwnersNow();
        });
        onClick('slgOwnersEnsureDirektionBtn', function () {
            runEnsureDirektion();
        });
        onClick('slgMemberSearchBtn', function () {
            runSearch('member');
        });
        onClick('slgMemberAddBtn', function () {
            runAddMember();
        });
        onClick('slgMembersReloadBtn', function () {
            loadMembersNow();
        });
        bindEnter('slgOwnerSearch', function () {
            runSearch('owner');
        });
        bindEnter('slgMemberSearch', function () {
            runSearch('member');
        });
        const ownersList = document.getElementById('slgOwnersList');
        if (ownersList) {
            ownersList.addEventListener('click', function (ev) {
                const t = ev.target;
                const btn = t && t.closest ? t.closest('button[data-slg-remove-owner]') : null;
                if (!btn) return;
                runRemoveOwner(btn.getAttribute('data-slg-remove-owner') || '');
            });
        }
        const membersList = document.getElementById('slgMembersList');
        if (membersList) {
            membersList.addEventListener('click', function (ev) {
                const t = ev.target;
                const btn = t && t.closest ? t.closest('button[data-slg-remove-member]') : null;
                if (!btn) return;
                runRemoveMember(btn.getAttribute('data-slg-remove-member') || '');
            });
        }
    }

    function bind(nextCtx) {
        ctx = nextCtx || null;
    }

    window.ms365SlgLiveDetails = {
        bind: bind,
        wire: wire,
        fillForm: fillForm,
        setMatchedMode: setMatchedMode,
        resetCaches: resetCaches,
        invalidateMembership: invalidateMembership,
        loadGroup: loadGroup,
        loadOwners: loadOwnersNow,
        loadMembers: loadMembersNow,
        onTab: onTab
    };
})();
