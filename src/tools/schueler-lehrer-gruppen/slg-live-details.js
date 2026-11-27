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
    let photoObjectUrl = '';
    let loadGroupGen = 0;
    /** @type {{ name: string, description: string, visibility: string, alias: string }|null} */
    let dirtySnapshot = null;
    let dirtyWired = false;

    function setFieldText(el, text, emptyFallback) {
        if (!el) return;
        const v = String(text == null ? '' : text);
        const shown = v || (emptyFallback != null ? String(emptyFallback) : '–');
        if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT') {
            el.value = v;
        } else {
            el.textContent = shown;
        }
    }

    function getFieldText(el) {
        if (!el) return '';
        if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT') {
            return String(el.value || '');
        }
        const t = String(el.textContent || '').trim();
        return t === '–' ? '' : t;
    }

    function readEditableSnapshot() {
        const name = document.getElementById('slgLiveName');
        const desc = document.getElementById('slgLiveDescription');
        const vis = document.getElementById('slgLiveVisibility');
        const alias = document.getElementById('slgLiveAlias');
        return {
            name: name ? String(name.value || '') : '',
            description: desc ? String(desc.value || '') : '',
            visibility: vis ? String(vis.value || '') : '',
            alias: alias && alias.tagName === 'INPUT' && !alias.readOnly ? String(alias.value || '') : ''
        };
    }

    function markClean() {
        dirtySnapshot = readEditableSnapshot();
        syncSaveButton();
        hideSaveFeedback();
    }

    function isDirty() {
        if (!dirtySnapshot) return false;
        const cur = readEditableSnapshot();
        return (
            cur.name !== dirtySnapshot.name ||
            cur.description !== dirtySnapshot.description ||
            cur.visibility !== dirtySnapshot.visibility ||
            cur.alias !== dirtySnapshot.alias
        );
    }

    function syncSaveButton() {
        const btn = document.getElementById('slgBtnUpdateGroup');
        if (!btn) return;
        const gid = getGroupId();
        btn.disabled = !gid || !isDirty();
    }

    function hideSaveFeedback() {
        const el = document.getElementById('slgSaveFeedback');
        if (el) el.classList.remove('is-visible');
    }

    function showSaveFeedback() {
        const el = document.getElementById('slgSaveFeedback');
        if (!el) return;
        el.innerHTML = '<i class="bi bi-check-circle-fill" aria-hidden="true"></i>Gespeichert';
        el.classList.add('is-visible');
        window.setTimeout(function () {
            el.classList.remove('is-visible');
        }, 2500);
    }

    async function confirmIfDirty(label) {
        if (!isDirty()) return true;
        const name = label || currentGroupLabel();
        const ok = await dlgConfirm(
            'Die Änderungen an „' + name + '“ wurden noch nicht gespeichert.',
            {
                title: 'Nicht gespeicherte Änderungen',
                okText: 'Verwerfen',
                cancelText: 'Abbrechen',
                danger: true
            }
        );
        return !!ok;
    }

    function wireDirtyTracking() {
        if (dirtyWired) return;
        dirtyWired = true;
        ['slgLiveName', 'slgLiveDescription', 'slgLiveVisibility', 'slgLiveAlias'].forEach(function (id) {
            const el = document.getElementById(id);
            if (!el) return;
            el.addEventListener('input', syncSaveButton);
            el.addEventListener('change', syncSaveButton);
        });
    }

    function revokePhotoObjectUrl() {
        if (!photoObjectUrl) return;
        try {
            URL.revokeObjectURL(photoObjectUrl);
        } catch {
            /* ignore */
        }
        photoObjectUrl = '';
    }

    function groupPhotoInitials(displayName) {
        if (typeof gug().groupPhotoInitials === 'function') {
            return gug().groupPhotoInitials(displayName);
        }
        const s = String(displayName || '').trim();
        if (!s) return '?';
        const parts = s.split(/\s+/).filter(Boolean);
        if (parts.length >= 2) {
            return (parts[0].charAt(0) + parts[1].charAt(0)).toUpperCase();
        }
        return s.slice(0, 2).toUpperCase();
    }

    function setGroupPhotoUi(opts) {
        const wrap = document.getElementById('slgGroupPhotoWrap');
        const img = document.getElementById('slgGroupPhotoImg');
        const initials = document.getElementById('slgGroupPhotoInitials');
        const removeBtn = document.getElementById('slgBtnRemoveGroupPhoto');
        const fileInput = document.getElementById('slgGroupPhotoFile');
        if (!wrap) return;

        const matched = !!(opts && opts.matched);
        if (!matched) {
            revokePhotoObjectUrl();
            wrap.hidden = true;
            wrap.style.display = 'none';
            if (img) {
                img.hidden = true;
                img.removeAttribute('src');
            }
            if (initials) {
                initials.hidden = false;
                initials.textContent = '–';
            }
            if (removeBtn) removeBtn.hidden = true;
            if (fileInput) fileInput.value = '';
            return;
        }

        wrap.hidden = false;
        wrap.style.display = '';
        const name = String((opts && opts.displayName) || '').trim();
        if (initials) {
            initials.textContent = groupPhotoInitials(name);
        }
        if (opts && opts.hasPhoto && opts.url) {
            if (img) {
                img.src = opts.url;
                img.alt = name ? 'Gruppenbild: ' + name : 'Gruppenbild';
                img.hidden = false;
            }
            if (initials) initials.hidden = true;
            if (removeBtn) removeBtn.hidden = false;
        } else {
            revokePhotoObjectUrl();
            if (img) {
                img.hidden = true;
                img.removeAttribute('src');
            }
            if (initials) initials.hidden = false;
            if (removeBtn) removeBtn.hidden = true;
        }
    }

    async function loadGroupPhoto(token, groupId, displayName) {
        if (!document.getElementById('slgGroupPhotoWrap')) return;
        revokePhotoObjectUrl();
        const name = String(displayName || '').trim();
        try {
            const blob = await gug().fetchGroupPhotoBlob(token, groupId);
            if (blob && blob.size) {
                photoObjectUrl = URL.createObjectURL(blob);
                setGroupPhotoUi({ matched: true, displayName: name, hasPhoto: true, url: photoObjectUrl });
                return;
            }
        } catch {
            /* Fallback: Initialen anzeigen */
        }
        setGroupPhotoUi({ matched: true, displayName: name, hasPhoto: false });
    }

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

    function pinAuthReturnHere() {
        try {
            if (typeof window.ms365AuthPinReturnUrl === 'function') {
                window.ms365AuthPinReturnUrl();
                return;
            }
        } catch {
            /* ignore */
        }
        try {
            sessionStorage.setItem('ms365-post-login-url', window.location.href);
        } catch {
            /* ignore */
        }
    }

    async function graphToken() {
        // Schreibrechte laufen über die IT-App und können die Redirect-Seite öffnen.
        // Vorher die aktuelle Werkzeugseite merken, nicht das zuletzt geöffnete.
        pinAuthReturnHere();
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

    /** Nur ID aus lokalem Match – noch kein Graph-Objekt. */
    function needsGraphGroupHydrate(group) {
        if (!group || !group.id) return false;
        if (String(group.displayName || '') === '(nicht geladen)') return false;
        if (group.createdDateTime || group.mailNickname || group.mail) return false;
        if (group.visibility === 'Public' || group.visibility === 'Private') return false;
        if (Array.isArray(group.groupTypes) && group.groupTypes.length) return false;
        if (group['@odata.type']) return false;
        return true;
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
        const statusRow = document.getElementById('slgDetailStatusRow');
        if (!group) {
            if (name) name.value = '';
            setFieldText(art, '', '–');
            setFieldText(mail, '', '–');
            if (mail) mail.removeAttribute('data-graph-mail');
            if (vis) {
                vis.disabled = false;
                vis.value = 'Private';
            }
            setFieldText(exp, '', 'Kein Ablaufdatum');
            setFieldText(alias, '', '–');
            if (desc) desc.value = '';
            setFieldText(idEl, '', '–');
            setFieldText(created, '', '–');
            setTeamStatusEl(teamEl, '', false);
            syncTeamActions(false, '');
            setGroupPhotoUi({ matched: false });
            syncMembershipTarget();
            if (statusRow) statusRow.hidden = true;
            markClean();
            return;
        }
        if (name) name.value = String(group.displayName || '');
        setFieldText(art, gug().groupArtLabel(group), '–');
        if (mail) {
            setFieldText(mail, String(group.mail || ''), '–');
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
            if (exp.tagName === 'INPUT' && exp.type === 'date') exp.value = isoToDateInput(group.expirationDateTime);
            else if (exp.tagName === 'INPUT' && exp.type === 'datetime-local') {
                const d = group.expirationDateTime ? new Date(group.expirationDateTime) : null;
                exp.value =
                    d && !isNaN(d.getTime()) ? d.toISOString().slice(0, 16) : '';
            } else {
                setFieldText(
                    exp,
                    group.expirationDateTime ? formatDateTimeAT(group.expirationDateTime) : '',
                    'Kein Ablaufdatum'
                );
            }
        }
        const expHint = document.getElementById('slgLiveExpiresHint');
        if (expHint) {
            expHint.textContent = group.expirationDateTime
                ? 'Kommt aus der Gruppen-Lebenszyklusrichtlinie des Tenants. „Ablauf verlängern“ schiebt das Datum um die Richtlinien-Dauer weiter.'
                : 'Kein Ablaufdatum in Microsoft. Graph kann kein frei gewähltes Datum setzen – nur verlängern, wenn im Tenant eine Lebenszyklusrichtlinie gilt.';
        }
        setFieldText(alias, String(group.mailNickname || ''), '–');
        if (desc) desc.value = String(group.description || '');
        setFieldText(idEl, String(group.id || ''), '–');
        setFieldText(created, formatDateTimeAT(group.createdDateTime), '–');
        if (statusRow) statusRow.hidden = false;
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
        setGroupPhotoUi({
            matched: true,
            displayName: group.displayName || '',
            hasPhoto: false
        });
        syncMembershipTarget();
        markClean();
        if (needsGraphGroupHydrate(group)) {
            const sub = document.getElementById('slgDetailSubtitle');
            if (sub) sub.textContent = 'Verknüpft – Details werden geladen …';
            queueMicrotask(function () {
                const gid = getGroupId();
                if (gid && String(gid) === String(group.id)) {
                    void loadGroup({ silent: true });
                }
            });
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
        const statusRow = document.getElementById('slgDetailStatusRow');
        const badge = document.getElementById('slgDetailSyncBadge');
        if (sub) {
            if (!has) {
                sub.textContent = 'Noch nicht mit Microsoft 365 verknüpft';
            } else if (liveGroup && liveGroup.displayName) {
                sub.textContent = 'Microsoft 365-Gruppe';
            } else {
                sub.textContent = 'Microsoft 365-Gruppe';
            }
        }
        if (statusRow) statusRow.hidden = false;
        if (badge) {
            badge.className =
                'gd-sync-badge ' + (has ? 'gd-sync-badge--ok' : 'gd-sync-badge--warn');
            badge.innerHTML = has
                ? '<i class="bi bi-check-circle-fill" aria-hidden="true"></i>Verknüpft'
                : '<i class="bi bi-exclamation-circle" aria-hidden="true"></i>Nicht verknüpft';
        }
    }

    function invalidateMembership() {
        ownersCache = [];
        membersCountCache = -1;
        ownersLoadedForId = '';
        membersLoadedForId = '';
    }

    function clearPersonPick(kind) {
        const inp = document.getElementById(kind === 'owner' ? 'slgOwnerSearch' : 'slgMemberSearch');
        const host = document.getElementById(kind === 'owner' ? 'slgOwnerSearchResults' : 'slgMemberSearchResults');
        if (inp) inp.value = '';
        if (host) {
            host.replaceChildren();
            if (host.dataset) delete host.dataset.groupId;
        }
    }

    function resetCaches() {
        loadGroupGen += 1;
        invalidateMembership();
        liveGroup = null;
        revokePhotoObjectUrl();
        clearPersonPick('owner');
        clearPersonPick('member');
    }

    function currentGroupLabel() {
        const title = document.getElementById('slgDetailTitle');
        const fromTitle = title ? String(title.textContent || '').trim() : '';
        if (fromTitle) return fromTitle;
        if (liveGroup && liveGroup.displayName) return String(liveGroup.displayName);
        return 'diese Gruppe';
    }

    function syncMembershipTarget() {
        if (window.ms365GroupDetail && typeof window.ms365GroupDetail.syncMembershipTarget === 'function') {
            window.ms365GroupDetail.syncMembershipTarget();
        }
    }

    function searchTargetsCurrentGroup(host) {
        const expected = host && host.dataset ? String(host.dataset.groupId || '') : '';
        const current = String(getGroupId() || '');
        return !!current && expected === current;
    }

    function readCheckedUsers(host) {
        const checked = host
            ? Array.from(host.querySelectorAll('.slg-user-checklist__item input[type="checkbox"]:checked'))
            : [];
        return checked
            .map(function (cb) {
                const row = cb.closest('.slg-user-checklist__item');
                const labelEl = row ? row.querySelector('span') : null;
                const label = labelEl ? String(labelEl.textContent || '').trim() : String(cb.value || '');
                return { id: String(cb.value || '').trim(), label: label };
            })
            .filter(function (x) {
                return !!x.id;
            });
    }

    function fillUserSearchMultiChecklist(selId, users) {
        const host = document.getElementById(selId);
        if (!host) return;
        host.dataset.groupId = String(getGroupId() || '');
        host.replaceChildren();
        const list = Array.isArray(users) ? users : [];
        if (!list.length) {
            const p = document.createElement('div');
            p.className = 'muted';
            p.textContent = '(keine Treffer)';
            host.appendChild(p);
            return;
        }

        const selAllRow = document.createElement('label');
        selAllRow.className = 'slg-user-checklist__selectall';
        const selAllCb = document.createElement('input');
        selAllCb.type = 'checkbox';
        selAllCb.setAttribute('aria-label', 'Alle auswählen');
        selAllRow.appendChild(selAllCb);
        const selAllTxt = document.createElement('span');
        selAllTxt.textContent = 'Alle auswählen (' + list.length + ')';
        selAllRow.appendChild(selAllTxt);
        host.appendChild(selAllRow);

        const items = [];
        list.forEach(function (u) {
            const label = document.createElement('label');
            label.className = 'slg-user-checklist__item';
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = u.id || '';
            cb.setAttribute('aria-label', gug().personLabel(u) || cb.value || 'Benutzer');
            const txt = document.createElement('span');
            txt.textContent = gug().personLabel(u) || (u.id ? String(u.id) : '');
            label.appendChild(cb);
            label.appendChild(txt);
            cb.addEventListener('change', function () {
                const checked = !!cb.checked;
                label.classList.toggle('is-checked', checked);
                const all = host.querySelectorAll('.slg-user-checklist__item input[type="checkbox"]');
                const checkedCount = host.querySelectorAll(
                    '.slg-user-checklist__item input[type="checkbox"]:checked'
                ).length;
                selAllCb.indeterminate = checkedCount > 0 && checkedCount < all.length;
                selAllCb.checked = checkedCount === all.length;
            });
            host.appendChild(label);
            items.push({ cb: cb, label: label });
        });

        function syncSelectAll() {
            const all = host.querySelectorAll('.slg-user-checklist__item input[type="checkbox"]');
            const checkedCount = host.querySelectorAll(
                '.slg-user-checklist__item input[type="checkbox"]:checked'
            ).length;
            selAllCb.indeterminate = checkedCount > 0 && checkedCount < all.length;
            selAllCb.checked = checkedCount === all.length;
        }

        selAllCb.addEventListener('change', function () {
            items.forEach(function (it) {
                it.cb.checked = selAllCb.checked;
                it.label.classList.toggle('is-checked', selAllCb.checked);
            });
            syncSelectAll();
        });

        // Init
        syncSelectAll();
    }

    function personParts(person) {
        const dn = person && person.displayName ? String(person.displayName).trim() : '';
        const mail = person
            ? String(person.mail || person.userPrincipalName || '').trim()
            : '';
        const label = gug().personLabel(person) || '–';
        return {
            name: dn || mail || label,
            mail: dn && mail && dn !== mail ? mail : dn ? '' : mail,
            initials: groupPhotoInitials(dn || mail || label)
        };
    }

    function setPeopleCount(elId, n) {
        const el = document.getElementById(elId);
        if (!el) return;
        if (n == null || n < 0) {
            el.textContent = '–';
            return;
        }
        el.textContent = String(n);
    }

    function setTabCount(elId, n) {
        const el = document.getElementById(elId);
        if (!el) return;
        if (n == null || n < 0) {
            el.hidden = true;
            el.textContent = '';
            return;
        }
        el.hidden = false;
        el.textContent = String(n);
    }

    function updateObjectHeaderSummary() {
        const counts = document.getElementById('slgDetailCounts');
        if (!counts) return;
        const parts = [];
        if (membersCountCache >= 0) {
            parts.push(membersCountCache === 1 ? '1 Mitglied' : String(membersCountCache) + ' Mitglieder');
        }
        if (ownersCache && ownersCache.length) {
            parts.push(ownersCache.length === 1 ? '1 Besitzer' : String(ownersCache.length) + ' Besitzer');
        }
        const teamEl = document.getElementById('slgLiveTeam');
        const teamTxt = teamEl ? String(teamEl.textContent || '').trim() : '';
        if (/Team vorhanden/i.test(teamTxt)) parts.push('Team vorhanden');
        if (!parts.length) {
            counts.hidden = true;
            counts.textContent = '';
            return;
        }
        counts.hidden = false;
        counts.textContent = parts.join(' · ');
    }

    function renderPersonRows(wrap, list, emptyText, removeAttr, opts) {
        if (!wrap) return;
        wrap.replaceChildren();
        const items = Array.isArray(list) ? list : [];
        const lastOwnerLock = !!(opts && opts.lastOwnerLock);
        if (!items.length) {
            const p = document.createElement('p');
            p.className = 'muted';
            p.style.margin = '8px 0';
            p.textContent = emptyText;
            wrap.appendChild(p);
            return;
        }
        items.forEach(function (person) {
            const parts = personParts(person);
            const row = document.createElement('div');
            row.className = 'gd-person-row';
            const avatar = document.createElement('span');
            avatar.className = 'gd-person-row__avatar';
            avatar.setAttribute('aria-hidden', 'true');
            avatar.textContent = parts.initials;
            const main = document.createElement('div');
            main.className = 'gd-person-row__main';
            const name = document.createElement('span');
            name.className = 'gd-person-row__name';
            name.textContent = parts.name;
            main.appendChild(name);
            if (parts.mail) {
                const mail = document.createElement('span');
                mail.className = 'gd-person-row__mail';
                mail.textContent = parts.mail;
                main.appendChild(mail);
            }
            if (opts && opts.metaText) {
                const meta = document.createElement('span');
                meta.className = 'gd-person-row__meta';
                meta.textContent = opts.metaText;
                main.appendChild(meta);
            }
            const ov = document.createElement('details');
            ov.className = 'gd-overflow';
            const sum = document.createElement('summary');
            sum.setAttribute('aria-label', 'Weitere Aktionen');
            sum.textContent = '···';
            const menu = document.createElement('div');
            menu.className = 'gd-overflow__menu';
            menu.setAttribute('role', 'menu');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'gd-overflow__danger';
            btn.setAttribute('role', 'menuitem');
            btn.textContent =
                removeAttr === 'data-slg-remove-owner' ? 'Als Besitzer entfernen' : 'Aus Gruppe entfernen';
            btn.setAttribute(removeAttr, person.id || '');
            if (lastOwnerLock && items.length <= 1) {
                btn.disabled = true;
                btn.title = 'Mindestens ein Besitzer ist erforderlich.';
            }
            menu.appendChild(btn);
            ov.appendChild(sum);
            ov.appendChild(menu);
            row.appendChild(avatar);
            row.appendChild(main);
            row.appendChild(ov);
            wrap.appendChild(row);
        });
    }

    function renderOwnersList(owners) {
        const list = Array.isArray(owners) ? owners : [];
        setPeopleCount('slgOwnersHeadCount', list.length);
        setTabCount('slgTabOwnersCount', list.length);
        renderPersonRows(
            document.getElementById('slgOwnersList'),
            list,
            'Keine Besitzer gefunden.',
            'data-slg-remove-owner',
            { lastOwnerLock: true }
        );
        updateObjectHeaderSummary();
    }

    function renderMembersList(result, totalCount) {
        const wrap = document.getElementById('slgMembersList');
        if (!wrap) return;
        const list = result && Array.isArray(result.items) ? result.items : [];
        const truncated = !!(result && result.truncated);
        const count = totalCount >= 0 ? totalCount : list.length;
        setPeopleCount('slgMembersHeadCount', count);
        setTabCount('slgTabMembersCount', count);
        renderPersonRows(wrap, list, truncated ? 'Keine Mitglieder in der gekürzten Anzeige.' : 'Keine Mitglieder.', 'data-slg-remove-member', {
            metaText: document.getElementById('slgBtnSync') ? '✓ Stammdaten + Microsoft 365' : ''
        });
        if (truncated) {
            const note = document.createElement('p');
            note.className = 'muted';
            note.style.margin = '8px 0 0';
            note.textContent = 'Anzeige gekürzt – nicht alle Mitglieder geladen.';
            wrap.appendChild(note);
        }
        updateObjectHeaderSummary();
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
            toast('Besitzer laden: ' + (e.message || e));
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
        const myGen = ++loadGroupGen;
        setMatchedMode(true);
        try {
            const token = await graphToken();
            if (myGen !== loadGroupGen || String(getGroupId() || '') !== String(gid)) return;
            const g = await gug().fetchGroup(token, gid);
            if (myGen !== loadGroupGen || String(getGroupId() || '') !== String(gid)) return;
            fillForm(g);
            setMatchedMode(true);
            await resolveTeamLink(token, g);
            if (myGen !== loadGroupGen || String(getGroupId() || '') !== String(gid)) return;
            await loadGroupPhoto(token, gid, g.displayName);
            if (myGen !== loadGroupGen || String(getGroupId() || '') !== String(gid)) return;
            if (ctx && ctx.onAfterLoad) await ctx.onAfterLoad(g);
            const tab = ctx && ctx.getActiveTab ? ctx.getActiveTab() : 'general';
            if (tab === 'owners') await loadOwnersNow();
            else if (tab === 'members') await loadMembersNow();
            if (!silent) toast('Gruppe geladen.');
        } catch (e) {
            if (myGen !== loadGroupGen || String(getGroupId() || '') !== String(gid)) return;
            const msg = String((e && e.message) || e || '');
            const redirectPending = /Weiterleitung zur Anmeldung/i.test(msg);
            if (
                !liveGroup ||
                String(liveGroup.id) !== String(gid) ||
                liveGroup.displayName === '(nicht geladen)'
            ) {
                fillForm({ id: gid });
            }
            setMatchedMode(true);
            if (!redirectPending && (!silent || !/popup_window_error|error opening popup/i.test(msg))) {
                toast('Gruppe laden: ' + msg);
            }
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
        if (aliasEl && aliasEl.tagName === 'INPUT' && !aliasEl.readOnly) {
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
            showSaveFeedback();
            toast(('Änderungen gespeichert.' + extra).trim());
        } catch (e) {
            toast('Speichern: ' + (e.message || e));
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
                'Für diese Microsoft 365‑Gruppe ein Microsoft Team anlegen?\n\nBesitzer und Mitglieder der Gruppe werden zum Team.',
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
        const gidAtStart = String(getGroupId() || '');
        if (btn) btn.disabled = true;
        try {
            const token = await graphToken();
            const users = await gug().searchUsers(token, q);
            if (String(getGroupId() || '') !== gidAtStart) {
                toast(
                    'Die Auswahl ist jetzt „' +
                        currentGroupLabel() +
                        '“. Bitte die Person erneut suchen, damit sie dieser Gruppe zugeordnet wird.'
                );
                return;
            }
            fillUserSearchMultiChecklist(selId, users);
            toast('Suche: ' + users.length + ' Treffer.');
        } catch (e) {
            toast('Suche: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    async function runAddOwner() {
        const host = document.getElementById('slgOwnerSearchResults');
        const selected = readCheckedUsers(host);
        if (!selected.length) {
            toast('Bitte zuerst mindestens einen Benutzer aus den Treffern auswählen.');
            return;
        }
        const gid = getGroupId();
        if (!gid) {
            toast('Bitte zuerst eine gematchte Gruppe wählen.');
            return;
        }
        if (!searchTargetsCurrentGroup(host)) {
            if (host) host.replaceChildren();
            toast(
                'Die Auswahl ist jetzt „' +
                    currentGroupLabel() +
                    '“. Bitte die Person erneut suchen, damit sie dieser Gruppe zugeordnet wird.'
            );
            return;
        }
        const names = selected.map(function (x) {
            return x.label;
        });
        if (
            !(await dlgConfirm(
                'Diese Person(en) als Besitzer zu „' +
                    currentGroupLabel() +
                    '“ hinzufügen?\n\n' +
                    names.join('\n') +
                    '\n\nBestehende Besitzer bleiben erhalten.',
                { title: 'Besitzer hinzufügen', okText: 'Hinzufügen' }
            ))
        ) {
            return;
        }
        const btn = document.getElementById('slgOwnerAddBtn');
        if (btn) btn.disabled = true;
        try {
            const token = await graphToken();
            let ok = 0;
            let fail = 0;
            for (const s of selected) {
                try {
                    await gug().addOwnerWithMemberFallback(token, gid, s.id);
                    ok++;
                } catch (e) {
                    fail++;
                }
            }
            await loadOwnersNow();
            toast(ok + ' Besitzer hinzugefügt' + (fail ? ', ' + fail + ' Fehler.' : '.'));
            if (host) host.replaceChildren();
        } catch (e) {
            toast('Besitzer hinzufügen: ' + (e.message || e));
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
            toast('Direktion als Besitzer hinzugefügt.');
        } catch (e) {
            toast('Direktion als Besitzer: ' + (e.message || e));
        }
    }

    async function runRemoveOwner(ownerId) {
        const gid = getGroupId();
        if (!gid || !ownerId) return;
        if (ownersCache.length <= 1) {
            toast('Der letzte Besitzer kann nicht entfernt werden.');
            return;
        }
        if (!(await dlgConfirm('Diesen Besitzer wirklich entfernen?', { title: 'Besitzer', okText: 'Entfernen', danger: true }))) {
            return;
        }
        try {
            const token = await graphToken();
            await gug().removeGroupOwner(token, gid, ownerId);
            await loadOwnersNow();
            toast('Besitzer entfernt.');
        } catch (e) {
            toast('Besitzer entfernen: ' + (e.message || e));
        }
    }

    async function runAddMember() {
        const host = document.getElementById('slgMemberSearchResults');
        const selected = readCheckedUsers(host);
        if (!selected.length) {
            toast('Bitte zuerst mindestens einen Benutzer aus den Treffern auswählen.');
            return;
        }
        const gid = getGroupId();
        if (!gid) {
            toast('Bitte zuerst eine gematchte Gruppe wählen.');
            return;
        }
        if (!searchTargetsCurrentGroup(host)) {
            if (host) host.replaceChildren();
            toast(
                'Die Auswahl ist jetzt „' +
                    currentGroupLabel() +
                    '“. Bitte die Person erneut suchen, damit sie dieser Gruppe zugeordnet wird.'
            );
            return;
        }
        const btn = document.getElementById('slgMemberAddBtn');
        if (btn) btn.disabled = true;
        try {
            const token = await graphToken();
            const groupNames = selected.map(function (x) {
                return x.label;
            });
            if (
                !(
                    await dlgConfirm(
                        'Diese Person(en) als Mitglied zu „' +
                            currentGroupLabel() +
                            '“ hinzufügen?\n\n' +
                            groupNames.join('\n'),
                        { title: 'Mitglieder hinzufügen', okText: 'Hinzufügen' }
                    )
                )
            ) {
                return;
            }

            let ok = 0;
            let skip = 0;
            let fail = 0;
            for (const s of selected) {
                try {
                    await gug().graphAddMember(token, gid, s.id);
                    ok++;
                } catch (e) {
                    if (gug().isDuplicateMemberError(e)) {
                        skip++;
                    } else {
                        fail++;
                        // Fehler pro Person werden in der Gesamt-Tost-Zeile zusammengefasst.
                    }
                }
            }

            membersLoadedForId = '';
            await loadMembersNow();
            toast(
                ok + ' hinzugefügt' +
                    (skip ? ', ' + skip + ' übersprungen' : '') +
                    (fail ? ', ' + fail + ' Fehler.' : '.')
            );
            // Search results leeren
            if (host) {
                host.replaceChildren();
            }
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

    async function runUploadGroupPhoto(file) {
        const gid = getGroupId();
        if (!gid || !file) return;
        const max =
            typeof gug().GROUP_PHOTO_MAX_BYTES === 'number' ? gug().GROUP_PHOTO_MAX_BYTES : 4 * 1024 * 1024;
        if (file.size > max) {
            toast('Bild ist zu groß (max. 4 MB).');
            return;
        }
        const okTypes = ['image/jpeg', 'image/png', 'image/webp'];
        const ct = String(file.type || '').trim() || 'image/jpeg';
        if (okTypes.indexOf(ct) === -1) {
            toast('Bitte JPEG, PNG oder WebP verwenden.');
            return;
        }
        setBusy(['slgBtnRemoveGroupPhoto'], true);
        try {
            const token = await graphToken();
            const hasTeam = !!(liveGroup && gug().groupHasTeam && gug().groupHasTeam(liveGroup));
            await gug().setGroupPhoto(token, gid, file, ct);
            const teamSync = await gug().syncTeamPhotoForGroup(token, gid, 'set', file, ct, { hasTeam: hasTeam });
            if (window.ms365GroupPhotoThumb && typeof window.ms365GroupPhotoThumb.invalidate === 'function') {
                window.ms365GroupPhotoThumb.invalidate(gid);
            }
            await loadGroupPhoto(token, gid, liveGroup && liveGroup.displayName);
            const extra =
                typeof gug().teamPhotoSyncHint === 'function' ? gug().teamPhotoSyncHint(teamSync) : '';
            toast(('Gruppenbild hochgeladen.' + extra).trim());
        } catch (e) {
            toast('Gruppenbild: ' + (e.message || e));
        } finally {
            setBusy(['slgBtnRemoveGroupPhoto'], false);
            const fi = document.getElementById('slgGroupPhotoFile');
            if (fi) fi.value = '';
        }
    }

    async function runRemoveGroupPhoto() {
        const gid = getGroupId();
        if (!gid) return;
        if (
            !(await dlgConfirm('Gruppenbild wirklich entfernen?', {
                title: 'Gruppenbild',
                okText: 'Entfernen',
                danger: true
            }))
        ) {
            return;
        }
        setBusy(['slgBtnRemoveGroupPhoto'], true);
        try {
            const token = await graphToken();
            const hasTeam = !!(liveGroup && gug().groupHasTeam && gug().groupHasTeam(liveGroup));
            await gug().deleteGroupPhoto(token, gid);
            const teamSync = await gug().syncTeamPhotoForGroup(token, gid, 'delete', undefined, undefined, {
                hasTeam: hasTeam
            });
            if (window.ms365GroupPhotoThumb && typeof window.ms365GroupPhotoThumb.invalidate === 'function') {
                window.ms365GroupPhotoThumb.invalidate(gid);
            }
            await loadGroupPhoto(token, gid, liveGroup && liveGroup.displayName);
            const extra =
                typeof gug().teamPhotoSyncHint === 'function' ? gug().teamPhotoSyncHint(teamSync) : '';
            toast(('Gruppenbild entfernt.' + extra).trim());
        } catch (e) {
            toast('Gruppenbild entfernen: ' + (e.message || e));
        } finally {
            setBusy(['slgBtnRemoveGroupPhoto'], false);
        }
    }

    function wire() {
        wireDirtyTracking();
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
            if (isDirty()) {
                confirmIfDirty().then(function (ok) {
                    if (ok) loadGroup({ silent: false });
                });
                return;
            }
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
        let ownerSearchTimer = 0;
        let memberSearchTimer = 0;
        const ownerSearch = document.getElementById('slgOwnerSearch');
        if (ownerSearch) {
            ownerSearch.addEventListener('input', function () {
                window.clearTimeout(ownerSearchTimer);
                const q = String(ownerSearch.value || '').trim();
                if (q.length < 2) {
                    const host = document.getElementById('slgOwnerSearchResults');
                    if (host) host.replaceChildren();
                    return;
                }
                ownerSearchTimer = window.setTimeout(function () {
                    runSearch('owner');
                }, 280);
            });
        }
        const memberSearch = document.getElementById('slgMemberSearch');
        if (memberSearch) {
            memberSearch.addEventListener('input', function () {
                window.clearTimeout(memberSearchTimer);
                const q = String(memberSearch.value || '').trim();
                if (q.length < 2) {
                    const host = document.getElementById('slgMemberSearchResults');
                    if (host) host.replaceChildren();
                    return;
                }
                memberSearchTimer = window.setTimeout(function () {
                    runSearch('member');
                }, 280);
            });
        }
        const photoFile = document.getElementById('slgGroupPhotoFile');
        if (photoFile) {
            photoFile.addEventListener('change', function () {
                const f = photoFile.files && photoFile.files[0];
                if (f) runUploadGroupPhoto(f);
            });
        }
        onClick('slgBtnRemoveGroupPhoto', function () {
            const ov = document.getElementById('slgPhotoOverflow');
            if (ov) ov.open = false;
            runRemoveGroupPhoto();
        });
        const ownersList = document.getElementById('slgOwnersList');
        if (ownersList) {
            ownersList.addEventListener('click', function (ev) {
                const t = ev.target;
                const btn = t && t.closest ? t.closest('button[data-slg-remove-owner]') : null;
                if (!btn || btn.disabled) return;
                const details = btn.closest('details');
                if (details) details.open = false;
                runRemoveOwner(btn.getAttribute('data-slg-remove-owner') || '');
            });
        }
        const membersList = document.getElementById('slgMembersList');
        if (membersList) {
            membersList.addEventListener('click', function (ev) {
                const t = ev.target;
                const btn = t && t.closest ? t.closest('button[data-slg-remove-member]') : null;
                if (!btn) return;
                const details = btn.closest('details');
                if (details) details.open = false;
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
        onTab: onTab,
        isDirty: isDirty,
        confirmIfDirty: confirmIfDirty,
        markClean: markClean,
        getFieldText: getFieldText,
        setFieldText: setFieldText
    };
})();
