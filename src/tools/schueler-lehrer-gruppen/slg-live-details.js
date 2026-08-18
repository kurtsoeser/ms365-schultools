(function () {
    'use strict';

    /** @type {null|{ toast: Function, dlgConfirm: Function, getGroupId: Function, getActiveTab: Function, ensureDirektionOwners: Function, onUnmatched: Function, onAfterLoad: Function, onMembersCount: Function, getGraphToken: Function, confirmUpdate: Function, alwaysMatched: boolean }} */
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

    async function graphToken() {
        if (ctx && typeof ctx.getGraphToken === 'function') return ctx.getGraphToken();
        return gug().getGraphToken();
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

    function isoToDateInput(iso) {
        const s = String(iso || '').trim();
        if (!s) return '';
        const d = new Date(s);
        if (isNaN(d.getTime())) return '';
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + day;
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
        const created = document.getElementById('slgLiveCreated');
        const teamEl = document.getElementById('slgLiveTeam');
        if (!group) {
            if (name) name.value = '';
            if (art) art.value = '';
            if (mail) mail.value = '';
            if (mail) mail.removeAttribute('data-graph-mail');
            if (vis) {
                vis.disabled = false;
                vis.value = 'Private';
            }
            if (exp) exp.value = '';
            if (alias) alias.value = '';
            if (desc) desc.value = '';
            if (idEl) idEl.value = '';
            if (created) created.value = '';
            setTeamStatusEl(teamEl, '', false);
            syncTeamActions(false, '');
            return;
        }
        if (name) name.value = String(group.displayName || '');
        if (art) art.value = gug().groupArtLabel(group);
        if (mail) {
            mail.value = String(group.mail || '');
            mail.setAttribute('data-graph-mail', String(group.mail || ''));
        }
        if (vis) {
            const v = String(group.visibility || '').trim();
            const hasEmpty = !!vis.querySelector('option[value=""]');
            const unified = !(gug().isUnifiedGroup) || gug().isUnifiedGroup(group);
            if (hasEmpty && !unified) {
                vis.value = '';
                vis.disabled = true;
            } else {
                vis.disabled = false;
                vis.value = v === 'Public' ? 'Public' : 'Private';
            }
        }
        if (exp) {
            if (exp.type === 'date') exp.value = isoToDateInput(group.expirationDateTime);
            else if (exp.type === 'datetime-local') {
                const d = group.expirationDateTime ? new Date(group.expirationDateTime) : null;
                exp.value =
                    d && !isNaN(d.getTime()) ? d.toISOString().slice(0, 16) : '';
            } else exp.value = formatDateTimeAT(group.expirationDateTime);
        }
        const expHint = document.getElementById('slgLiveExpiresHint');
        if (expHint) {
            expHint.textContent = group.expirationDateTime
                ? 'Kommt aus der Gruppen-Lebenszyklusrichtlinie des Tenants. „Ablauf verlängern“ schiebt das Datum um die Richtlinien-Dauer weiter.'
                : 'Kein Ablaufdatum in Microsoft. Graph kann kein frei gewähltes Datum setzen – nur verlängern, wenn im Tenant eine Lebenszyklusrichtlinie gilt.';
        }
        if (alias) alias.value = String(group.mailNickname || '');
        if (desc) desc.value = String(group.description || '');
        if (idEl) idEl.value = String(group.id || '');
        if (created) created.value = formatDateTimeAT(group.createdDateTime);
        const unified = !(gug().isUnifiedGroup) || gug().isUnifiedGroup(group);
        const hasTeam = !!(gug().groupHasTeam && gug().groupHasTeam(group));
        if (!unified) {
            setTeamStatusEl(teamEl, 'Kein Microsoft 365‑Team', false);
            syncTeamActions(false, '');
            const btn = document.getElementById('slgBtnProvisionTeam');
            if (btn) btn.hidden = true;
        } else {
            setTeamStatusEl(teamEl, hasTeam ? 'Team vorhanden' : 'Kein Team', hasTeam);
            syncTeamActions(hasTeam, hasTeam ? group.teamWebUrl || fallbackTeamUrl(group.id) : '');
        }
    }

    function setTeamStatusEl(el, text, hasTeam) {
        if (!el) return;
        const t = String(text || '');
        if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
            el.value = t;
            return;
        }
        el.replaceChildren();
        if (t) {
            const ico = document.createElement('i');
            ico.className = hasTeam ? 'bi bi-microsoft-teams' : 'bi bi-dash-circle-fill';
            ico.setAttribute('aria-hidden', 'true');
            el.appendChild(ico);
            el.appendChild(document.createTextNode(t));
        } else {
            el.textContent = '–';
        }
        el.classList.toggle('is-ok', !!hasTeam);
        el.classList.toggle('is-warn', !hasTeam && !!t);
    }

    function fallbackTeamUrl(groupId) {
        const gid = String(groupId || '').trim();
        if (!gid) return '';
        try {
            if (typeof gug().teamDeepLink === 'function') return gug().teamDeepLink(gid, '') || '';
        } catch {
            /* ignore */
        }
        return (
            'https://teams.microsoft.com/l/team/' +
            encodeURIComponent(gid) +
            '/conversations?groupId=' +
            encodeURIComponent(gid)
        );
    }

    function syncTeamActions(hasTeam, url) {
        const btn = document.getElementById('slgBtnProvisionTeam');
        const a = document.getElementById('slgLiveTeamLink');
        const hint = document.getElementById('slgLiveTeamHint');
        const gid = getGroupId();
        if (btn) {
            const showCreate = !hasTeam;
            btn.hidden = !showCreate;
            btn.disabled = !showCreate || !gid;
            btn.title = showCreate ? 'Team für diese Microsoft 365‑Gruppe anlegen' : '';
        }
        if (a) {
            const href = String(url || '').trim();
            a.hidden = !hasTeam;
            a.href = hasTeam && href ? href : '#';
        }
        if (hint) hint.hidden = !!hasTeam;
    }

    async function resolveTeamLink(token, group) {
        if (!group || !group.id) {
            syncTeamActions(false, '');
            return;
        }
        const hasTeam = !!(gug().groupHasTeam && gug().groupHasTeam(group));
        if (!hasTeam) {
            syncTeamActions(false, '');
            return;
        }
        let url = fallbackTeamUrl(group.id);
        try {
            if (typeof gug().fetchTeamWebUrl === 'function') {
                const fetched = await gug().fetchTeamWebUrl(token, group.id);
                if (fetched) url = fetched;
            }
        } catch {
            /* Fallback bleibt */
        }
        if (liveGroup && liveGroup.id === group.id) liveGroup.teamWebUrl = url;
        syncTeamActions(true, url);
    }

    function setMatchedMode(has) {
        const always = !!(ctx && ctx.alwaysMatched);
        const matched = always || !!has;
        showEl('slgUnmatchedPanel', !matched);
        showEl('slgMatchedPanel', matched);
        showEl('slgOwnersUnmatched', !matched);
        showEl('slgOwnersMatched', matched);
        showEl('slgMembersUnmatched', !matched);
        showEl('slgMembersMatched', matched);
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
            const token = await graphToken();
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
            const token = await graphToken();
            try {
                membersCountCache = await gug().fetchGroupMemberCount(token, gid);
            } catch {
                membersCountCache = -1;
            }
            const result = await gug().fetchGroupMembers(token, gid);
            membersLoadedForId = gid;
            renderMembersList(result, membersCountCache);
            if (ctx && typeof ctx.onMembersCount === 'function') {
                ctx.onMembersCount(gid, membersCountCache);
            }
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
            const token = await graphToken();
            const g = await gug().fetchGroup(token, gid);
            fillForm(g);
            setMatchedMode(true);
            await resolveTeamLink(token, g);
            if (ctx && ctx.onAfterLoad) await ctx.onAfterLoad(g);
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
        const aliasEl = document.getElementById('slgLiveAlias');
        const displayName = nameEl ? String(nameEl.value || '').trim() : '';
        if (!displayName) {
            toast('Bitte einen Anzeigenamen eingeben.');
            return;
        }
        if (ctx && typeof ctx.confirmUpdate === 'function') {
            if (!(await ctx.confirmUpdate(liveGroup))) return;
        }
        const patch = {
            displayName: displayName,
            description: descEl ? descEl.value : '',
            visibility: visEl ? visEl.value : 'Private'
        };
        if (aliasEl && !aliasEl.readOnly) {
            const raw = String(aliasEl.value || '').trim();
            if (!raw) {
                toast('Bitte einen Alias (Mail‑Nickname) eingeben.');
                return;
            }
            const nick = gug().sanitizeUnifiedGroupMailNickname(raw);
            aliasEl.value = nick;
            patch.mailNickname = nick;
        }
        try {
            const token = await graphToken();
            await gug().patchGroup(token, gid, patch);
            await loadGroup({ silent: true });
            let extra = '';
            if (ctx && typeof ctx.onAfterUpdate === 'function') {
                extra = (await ctx.onAfterUpdate(liveGroup)) || '';
            }
            toast(('Gruppe aktualisiert.' + extra).trim());
        } catch (e) {
            toast('Update: ' + (e.message || e));
        }
    }

    async function runProvisionTeam() {
        const gid = getGroupId();
        if (!gid) return;
        if (liveGroup && gug().groupHasTeam && gug().groupHasTeam(liveGroup)) {
            toast('Diese Gruppe hat bereits ein Microsoft Team.');
            return;
        }
        if (
            !(await dlgConfirm(
                'Für diese Microsoft 365‑Gruppe ein Microsoft Team anlegen?\n\nOwner und Mitglieder der Gruppe werden zum Team.',
                { title: 'Team anlegen', okText: 'Team anlegen' }
            ))
        ) {
            return;
        }
        setBusy(['slgBtnProvisionTeam'], true);
        try {
            const token = await graphToken();
            await gug().provisionTeamForGroup(token, gid);
            await loadGroup({ silent: true });
            toast('Team angelegt.');
        } catch (e) {
            const msg = String((e && e.message) || e || '');
            if (/\b409\b/.test(msg) || /Conflict|already exists|already provisioned/i.test(msg)) {
                await loadGroup({ silent: true });
                toast('Team war bereits vorhanden.');
            } else {
                toast('Team anlegen: ' + (e.message || e));
            }
        } finally {
            const hasTeam = !!(liveGroup && gug().groupHasTeam && gug().groupHasTeam(liveGroup));
            syncTeamActions(hasTeam, hasTeam ? (liveGroup && liveGroup.teamWebUrl) || fallbackTeamUrl(gid) : '');
        }
    }

    async function runRenewExpiration() {
        const gid = getGroupId();
        if (!gid) return;
        if (
            !(await dlgConfirm(
                'Ablaufdatum dieser Gruppe verlängern?\n\nMicrosoft schiebt das Datum um die Dauer der Tenant-Lebenszyklusrichtlinie weiter. Ein frei gewähltes Datum ist per Graph nicht möglich.',
                { title: 'Ablauf verlängern', okText: 'Verlängern' }
            ))
        ) {
            return;
        }
        try {
            const token = await graphToken();
            if (typeof gug().renewGroup !== 'function') {
                throw new Error('renewGroup fehlt.');
            }
            await gug().renewGroup(token, gid);
            await loadGroup({ silent: true });
            const shown = liveGroup && liveGroup.expirationDateTime ? formatDateTimeAT(liveGroup.expirationDateTime) : '';
            toast(
                shown
                    ? 'Ablauf verlängert. Neues Datum in Microsoft: ' + shown + '.'
                    : 'Verlängern gemeldet. Wenn weiter kein Datum erscheint, gilt im Tenant keine Gruppen-Lebenszyklusrichtlinie.'
            );
        } catch (e) {
            const msg = String((e && e.message) || e || '');
            if (/lifecycle|policy|not enabled|does not have/i.test(msg)) {
                toast(
                    'Ablauf verlängern nicht möglich: Im Tenant fehlt eine Gruppen-Lebenszyklusrichtlinie (Entra → Gruppen → Ablauf). Ein einzelnes Datum lässt sich per Graph nicht setzen.'
                );
            } else {
                toast('Ablauf verlängern: ' + (e.message || e));
            }
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
            const token = await graphToken();
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
            const token = await graphToken();
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
            const token = await graphToken();
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
            const token = await graphToken();
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
            const token = await graphToken();
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
            const token = await graphToken();
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
        onClick('slgBtnRenewExpires', function () {
            runRenewExpiration();
        });
        onClick('slgBtnProvisionTeam', function () {
            runProvisionTeam();
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
