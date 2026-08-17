(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-schueler-lehrer-gruppen-v2';

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

    function gd() {
        const G = window.ms365GroupDetail;
        if (!G) throw new Error('group-detail.js muss vor diesem Skript geladen werden.');
        return G;
    }

    async function getGraphToken() {
        return gug().getGraphToken();
    }

    /** @type {'schueler' | 'lehrer'} */
    let activeKind = 'schueler';

    /** @type {{ schuelerGroupId: string|null, lehrerGroupId: string|null }} */
    let matched = { schuelerGroupId: null, lehrerGroupId: null };

    /** @type {{ students: string[], teachers: string[], direktion: string[] }} */
    let listCache = { students: [], teachers: [], direktion: [] };

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
        if (el) el.replaceChildren();
    }

    function loadTenantSettings() {
        if (typeof window.ms365TenantSettingsLoad !== 'function') return null;
        return window.ms365TenantSettingsLoad();
    }

    function isDirektionRole(roleRaw) {
        const r = normStr(roleRaw).toLowerCase();
        if (!r) return false;
        return r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1;
    }

    function collectDirektionOwnerEmails(settings) {
        const out = [];
        const seen = new Set();
        const admin = settings && Array.isArray(settings.admin) ? settings.admin : [];
        admin.forEach(function (row) {
            if (!isDirektionRole(row && row.role)) return;
            const em = normEmail(row && row.email);
            if (!em || em.indexOf('@') === -1) return;
            if (seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    function collectEmails(arr) {
        const out = [];
        const seen = new Set();
        (Array.isArray(arr) ? arr : []).forEach(function (row) {
            const em = normEmail(row && row.email);
            if (!em || em.indexOf('@') === -1) return;
            if (seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    function readLists() {
        const settings = loadTenantSettings();
        listCache.students = collectEmails(settings && settings.students);
        listCache.teachers = collectEmails(settings && settings.teachers);
        listCache.direktion = collectDirektionOwnerEmails(settings);
    }

    function updateLeftListUi() {
        const sCount = document.getElementById('slgSchuelerCount');
        const tCount = document.getElementById('slgLehrerCount');
        if (sCount) sCount.textContent = String(listCache.students.length);
        if (tCount) tCount.textContent = String(listCache.teachers.length);

        const sLine = document.getElementById('slgSchuelerLine');
        const tLine = document.getElementById('slgLehrerLine');
        if (sLine) sLine.textContent = matched.schuelerGroupId ? 'Gematcht: ' + matched.schuelerGroupId : 'Noch kein Match';
        if (tLine) tLine.textContent = matched.lehrerGroupId ? 'Gematcht: ' + matched.lehrerGroupId : 'Noch kein Match';
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
        const el = document.getElementById('slgMemberPreview');
        if (!el) return;
        el.replaceChildren();
        const list = activeKind === 'schueler' ? listCache.students : listCache.teachers;
        const first = list.slice(0, 30);
        if (!first.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent = 'Keine E‑Mails in dieser Liste.';
            el.appendChild(p);
            return;
        }
        first.forEach(function (em) {
            const d = document.createElement('div');
            d.textContent = em;
            d.style.padding = '4px 0';
            d.style.borderBottom = '1px solid #eef1f4';
            el.appendChild(d);
        });
        if (list.length > first.length) {
            const more = document.createElement('div');
            more.className = 'muted';
            more.style.paddingTop = '8px';
            more.textContent = '… und ' + String(list.length - first.length) + ' weitere.';
            el.appendChild(more);
        }
    }

    function applyCreateDefaults() {
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const desc = document.getElementById('slgNewDescription');
        if (activeKind === 'schueler') {
            if (dn && !dn.value) dn.value = 'Schüler:innen';
            if (nn && !nn.value) nn.value = 'schueler';
            if (desc && !desc.value) desc.value = 'Alle Schüler:innen (MS365-Schulverwaltung / Schul‑Liste)';
        } else {
            if (dn && !dn.value) dn.value = 'Lehrer:innen';
            if (nn && !nn.value) nn.value = 'lehrer';
            if (desc && !desc.value) desc.value = 'Alle Lehrer:innen (MS365-Schulverwaltung / Schul‑Liste)';
        }
    }

    function getActiveMatchedId() {
        return activeKind === 'schueler' ? matched.schuelerGroupId : matched.lehrerGroupId;
    }

    function setActiveMatchedId(id) {
        if (activeKind === 'schueler') matched.schuelerGroupId = id;
        else matched.lehrerGroupId = id;
        live().resetCaches();
    }

    function setActiveKind(kind) {
        activeKind = kind === 'lehrer' ? 'lehrer' : 'schueler';
        const title = document.getElementById('slgDetailTitle');
        if (title) title.textContent = activeKind === 'schueler' ? 'Schüler:innen' : 'Lehrer:innen';

        document.querySelectorAll('button[data-slg-kind]').forEach(function (btn) {
            const on = btn.getAttribute('data-slg-kind') === activeKind;
            btn.setAttribute('aria-current', on ? 'true' : 'false');
        });

        applyCreateDefaults();
        live().resetCaches();
        gd().clearSearchResults();
        const gid = getActiveMatchedId();
        live().setMatchedMode(!!gid);
        live().fillForm(gid ? { id: gid } : null);
        updateLeftListUi();
        gd().setTab('general');
    }

    async function runSyncMembers() {
        const gid = getActiveMatchedId();
        if (!gid) {
            toast('Zuerst eine Gruppe matchen oder anlegen.');
            return;
        }
        const emails = activeKind === 'schueler' ? listCache.students : listCache.teachers;
        if (!emails.length) {
            toast('Keine E‑Mails in dieser Liste.');
            return;
        }
        clearSyncLog();
        appendSyncLog(
            'Start: ' + (activeKind === 'schueler' ? 'Schüler:innen' : 'Lehrer:innen') + ' (' + emails.length + ' Adressen) …',
            ''
        );
        try {
            const token = await getGraphToken();
            const label = activeKind === 'schueler' ? 'Schüler' : 'Lehrer';
            const lc = window.ms365StudentClassLifecycle;
            let joinEmails = emails;
            let leaveEmails = [];
            if (lc && typeof lc.reconcileSammelgruppe === 'function' && typeof gug().fetchGroupMembers === 'function') {
                const mem = await gug().fetchGroupMembers(token, gid);
                const current = (mem.items || [])
                    .map(function (m) {
                        return String((m && (m.mail || m.userPrincipalName)) || '')
                            .trim()
                            .toLowerCase();
                    })
                    .filter(function (em) {
                        return em.indexOf('@') !== -1;
                    });
                const rec = lc.reconcileSammelgruppe(emails, current);
                joinEmails = rec.join;
                leaveEmails = rec.leave;
                appendSyncLog('Abgleich mit Stammliste: +' + joinEmails.length + ' / −' + leaveEmails.length + '.', '');
            }
            if (joinEmails.length) {
                const r = await gug().syncEmailsToGroup(token, gid, joinEmails, label, appendSyncLog);
                appendSyncLog('Aufnehmen: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            }
            if (leaveEmails.length && typeof gug().removeEmailsFromGroup === 'function') {
                const r = await gug().removeEmailsFromGroup(token, gid, leaveEmails, label, appendSyncLog);
                appendSyncLog('Entfernen: ' + r.ok + ' OK, übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            }
            if (!joinEmails.length && !leaveEmails.length) {
                appendSyncLog('Keine Änderungen gegenüber der Stammliste.', 'ok');
            }
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
            activeKind: activeKind,
            matched: {
                schuelerGroupId: matched.schuelerGroupId,
                lehrerGroupId: matched.lehrerGroupId
            },
            slgNewDisplayName: document.getElementById('slgNewDisplayName')
                ? document.getElementById('slgNewDisplayName').value
                : '',
            slgNewMailNick: document.getElementById('slgNewMailNick')
                ? document.getElementById('slgNewMailNick').value
                : '',
            slgNewDescription: document.getElementById('slgNewDescription')
                ? document.getElementById('slgNewDescription').value
                : '',
            slgNewCreateTeam: document.getElementById('slgNewCreateTeam')
                ? !!document.getElementById('slgNewCreateTeam').checked
                : false
        };
    }

    function applyStateObject(o) {
        if (!o || typeof o !== 'object') return;
        if (o.matched && typeof o.matched === 'object') {
            matched.schuelerGroupId = o.matched.schuelerGroupId ? String(o.matched.schuelerGroupId) : null;
            matched.lehrerGroupId = o.matched.lehrerGroupId ? String(o.matched.lehrerGroupId) : null;
        }
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const dd = document.getElementById('slgNewDescription');
        const ct = document.getElementById('slgNewCreateTeam');
        if (dn && o.slgNewDisplayName !== undefined) dn.value = String(o.slgNewDisplayName || '');
        if (nn && o.slgNewMailNick !== undefined) nn.value = String(o.slgNewMailNick || '');
        if (dd && o.slgNewDescription !== undefined) dd.value = String(o.slgNewDescription || '');
        if (ct && o.slgNewCreateTeam !== undefined) ct.checked = !!o.slgNewCreateTeam;
        setActiveKind(o.activeKind === 'lehrer' ? 'lehrer' : 'schueler');
    }

    function saveState() {
        try {
            const obj = buildStateObject();
            localStorage.setItem(STORAGE_KEY, JSON.stringify(obj));
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: obj.matched,
                    slgDraft: {
                        activeKind: obj.activeKind,
                        slgNewDisplayName: obj.slgNewDisplayName,
                        slgNewMailNick: obj.slgNewMailNick,
                        slgNewDescription: obj.slgNewDescription,
                        slgNewCreateTeam: obj.slgNewCreateTeam
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
                const hasIds = su && su.matched && !!(su.matched.schuelerGroupId || su.matched.lehrerGroupId);
                if (hasIds || !rawLocal) {
                    const d = su.slgDraft || {};
                    applyStateObject({
                        matched: su.matched,
                        activeKind: d.activeKind === 'lehrer' ? 'lehrer' : 'schueler',
                        slgNewDisplayName: d.slgNewDisplayName,
                        slgNewMailNick: d.slgNewMailNick,
                        slgNewDescription: d.slgNewDescription,
                        slgNewCreateTeam: d.slgNewCreateTeam
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
            matched = { schuelerGroupId: null, lehrerGroupId: null };
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: { schuelerGroupId: null, lehrerGroupId: null }
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

    function onClick(id, fn) {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', fn);
    }

    function mountDetail() {
        gd().mount('#groupDetailHost', {
            title: 'Schüler:innen',
            searchPlaceholder: 'z. B. lehrer oder @schule.at',
            unmatchedCreateHint:
                'Legt eine Microsoft 365‑Gruppe (Unified) an. Optional auch als Team bereitstellen.',
            membersUnmatchedHint:
                'Mitglieder kommen aus der Schul‑Liste. Nach dem Match können Sie live verwalten und die Liste synchronisieren.',
            membersUnmatchedTitle: 'Vorschau Schul‑Liste (erste 30)',
            membersMatchedHint:
                'Live aus Microsoft Graph. „Mitglieder synchronisieren“ fügt fehlende Adressen aus der Schul‑Liste hinzu (entfernt niemanden).',
            features: { syncMembers: true },
            live: {
                toast: toast,
                dlgConfirm: dlgConfirm,
                getGroupId: getActiveMatchedId,
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
            },
            match: {
                persistMatch: function (g) {
                    setActiveMatchedId(String(g.id));
                    saveState();
                },
                persistUnmatch: function () {
                    setActiveMatchedId(null);
                    saveState();
                },
                ensureOwners: function (token, gid) {
                    return ensureOwners(token, gid);
                },
                afterMatch: function () {
                    updateLeftListUi();
                }
            },
            onTabUnmatched: function (tab) {
                if (tab === 'owners') renderOwnerPreview();
                if (tab === 'members') renderMemberPreview();
            }
        });
    }

    function wire() {
        const listHost = document.getElementById('slgListItems');
        if (listHost) {
            listHost.addEventListener('click', function (ev) {
                const t = ev.target;
                if (!t || !t.closest) return;
                const item = t.closest('button[data-slg-kind]');
                if (!item) return;
                const kind = item.getAttribute('data-slg-kind');
                setActiveKind(kind === 'lehrer' ? 'lehrer' : 'schueler');
                saveState();
                if (getActiveMatchedId()) live().loadGroup({ silent: true });
            });
        }

        onClick('slgBtnReloadLists', function () {
            readLists();
            updateLeftListUi();
            renderOwnerPreview();
            renderMemberPreview();
            toast('Listen neu eingelesen.');
        });
        onClick('slgBtnSync', function () {
            runSyncMembers();
        });
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
        mountDetail();
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
