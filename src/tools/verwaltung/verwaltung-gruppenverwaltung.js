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

    function gd() {
        const G = window.ms365GroupDetail;
        if (!G) throw new Error('group-detail.js muss vor diesem Skript geladen werden.');
        return G;
    }

    async function getGraphToken() {
        return gug().getGraphToken();
    }

    /** @type {string|null} */
    let matchedGroupId = null;

    /** @type {{ members: string[], direktion: string[], rows: { role: string, name: string, email: string, defaultKey?: string }[], roles: { code: string, name: string }[] }} */
    let listCache = { members: [], direktion: [], rows: [], roles: [] };

    /** @type {'group'|'role'} */
    let activeView = 'group';
    /** @type {string} */
    let activeRoleCode = '';
    let listFilter = '';
    /** @type {Record<string, number>} */
    let graphMemberCounts = {};
    let countsFetchGen = 0;
    /** @type {ReturnType<import('../../shared/membership-review-ui.js').createMembershipReview>|null} */
    let membershipReview = null;

    function graphCountFor(groupId) {
        const id = String(groupId || '').trim();
        if (!id) return null;
        const n = graphMemberCounts[id];
        return typeof n === 'number' && n >= 0 ? n : null;
    }

    async function refreshGraphMemberCounts() {
        const gid = matchedGroupId;
        if (!gid) {
            updateMismatchUi();
            return;
        }
        const gen = ++countsFetchGen;
        try {
            const token = await getGraphToken();
            if (gen !== countsFetchGen) return;
            const n = await gug().fetchGroupMemberCount(token, gid);
            if (typeof n === 'number' && n >= 0) graphMemberCounts[gid] = n;
            if (gen !== countsFetchGen) return;
            updateMismatchUi();
        } catch {
            updateMismatchUi();
        }
    }

    function paintVerwaltungCounts() {
        const listN = (listCache.members || []).length;
        const gid = matchedGroupId;
        const groupN = graphCountFor(gid);
        const wrap = document.getElementById('slgVerwaltungCounts');
        const listEl = document.getElementById('slgVerwaltungCount');
        const groupEl = document.getElementById('slgVerwaltungGroupCount');
        if (listEl) listEl.textContent = String(listN);
        if (groupEl) groupEl.textContent = gid ? (groupN === null ? '–' : String(groupN)) : '–';
        if (!wrap) return;
        wrap.classList.remove('is-match', 'is-mismatch');
        const known = gid && groupN !== null;
        if (known) {
            const same = listN === groupN;
            wrap.classList.add(same ? 'is-match' : 'is-mismatch');
            wrap.title = same
                ? 'Verwaltungsliste und Gruppe: je ' + listN + ' – Anzahl stimmt überein.'
                : 'Verwaltungsliste: ' + listN + ' · Gruppe: ' + groupN + ' Mitglieder.';
        } else {
            wrap.title = gid
                ? 'Verwaltungsliste: ' + listN + ' E-Mails. Mitgliederzahl wird aus Microsoft Graph geladen.'
                : 'Verwaltungsliste: ' + listN + ' E-Mails. Noch keine Gruppe gematcht.';
        }
    }

    function updateMismatchUi() {
        paintVerwaltungCounts();
        if (!membershipReview) return;
        const gid = matchedGroupId;
        const listN = (listCache.members || []).length;
        const groupN = graphCountFor(gid);
        if (gid && groupN !== null && listN !== groupN) {
            membershipReview.updateMismatchBar([
                {
                    key: 'verwaltung',
                    label: 'Verwaltung',
                    listN: listN,
                    groupN: groupN,
                    gid: gid
                }
            ]);
        } else {
            membershipReview.updateMismatchBar([]);
        }
    }

    function initMembershipReview() {
        const R = window.ms365MembershipReviewUi;
        if (!R || typeof R.createMembershipReview !== 'function') return;
        membershipReview = R.createMembershipReview({
            mode: 'sammelgruppe',
            tool: 'verwaltung',
            syncLabel: 'Verwaltung',
            getGraphToken: getGraphToken,
            getGroupId: getActiveMatchedId,
            getLocalEmails: function () {
                return listCache.members || [];
            },
            getActiveReviewKey: function () {
                return 'verwaltung';
            },
            getReviewTitle: function () {
                return 'Mitglieder-Abgleich: Verwaltung';
            },
            toast: toast,
            dlgConfirm: dlgConfirm,
            appendSyncLog: appendSyncLog,
            live: {
                invalidateMembership: function () {
                    live().invalidateMembership();
                },
                loadMembers: function () {
                    return live().loadMembers();
                }
            },
            refreshCounts: refreshGraphMemberCounts,
            onAfterChange: async function () {
                readLists();
                updateLeftListUi();
                renderMemberPreview();
            },
            openImport: async function (emails) {
                const Ui = window.ms365MembershipImportUi;
                if (!Ui || typeof Ui.openMembershipImportDialog !== 'function') {
                    throw new Error('membership-import-ui.js fehlt.');
                }
                return Ui.openMembershipImportDialog({
                    kind: 'verwaltung',
                    emails: emails,
                    getGraphToken: getGraphToken,
                    loadSettings: function () {
                        return loadTenantSettings();
                    },
                    saveSettings: function (settings) {
                        if (typeof window.ms365TenantSettingsSave === 'function') {
                            window.ms365TenantSettingsSave(settings);
                        }
                        readLists();
                        persistLists();
                    },
                    toast: toast,
                    dlgConfirm: dlgConfirm,
                    logAction: function (entry) {
                        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                            window.ms365ActionLog.append(
                                Object.assign({ tool: 'verwaltung' }, entry || {})
                            );
                        }
                    },
                    onApplied: async function () {
                        updateLeftListUi();
                        renderMemberPreview();
                        if (membershipReview && membershipReview.getState()) {
                            await membershipReview.loadReview('verwaltung');
                        }
                    }
                });
            },
            logAction: function (action, target, summary) {
                if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                    window.ms365ActionLog.append({
                        tool: 'verwaltung',
                        action: action,
                        target: target,
                        summary: summary,
                        result: 'ok'
                    });
                }
            },
            labels: {
                onlyLocalTitle: 'Nur in der Verwaltungsliste',
                onlyLocalHint: 'In den Stammdaten, aber nicht in der Microsoft-365-Gruppe.',
                onlyGraphTitle: 'Nur in der Microsoft-365-Gruppe',
                onlyGraphHint:
                    'In der Gruppe online, aber nicht in der Verwaltungsliste – Rolle zuweisen und in Stammdaten übernehmen.'
            }
        });
        membershipReview.wire();
    }

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

    function personMatchesRole(row, role) {
        if (typeof window.ms365TenantSettingsPersonMatchesAdminRole === 'function') {
            return window.ms365TenantSettingsPersonMatchesAdminRole(row, role);
        }
        if (!row || !role) return false;
        const r = normStr(row.role).toLowerCase();
        const n = normStr(role.name).toLowerCase();
        const c = normStr(role.code).toLowerCase();
        return !!(r && (r === n || r === c));
    }

    function roleCodeFromName(name) {
        if (typeof window.ms365TenantSettingsAdminRoleCodeFromName === 'function') {
            return window.ms365TenantSettingsAdminRoleCodeFromName(name);
        }
        return normStr(name)
            .toUpperCase()
            .replace(/\s+/g, '')
            .replace(/[^A-Z0-9ÄÖÜß-]/g, '')
            .slice(0, 24);
    }

    function uniqueRoleCode(desired, usedCodes) {
        let code = roleCodeFromName(desired) || 'ROLLE';
        const used = new Set(
            (usedCodes || []).map(function (c) {
                return String(c || '').toLowerCase();
            })
        );
        if (!used.has(code.toLowerCase())) return code;
        let i = 2;
        while (used.has((code + String(i)).toLowerCase())) i += 1;
        return (code + String(i)).slice(0, 24);
    }

    function getActiveRole() {
        const code = normStr(activeRoleCode).toUpperCase();
        if (!code) return null;
        for (let i = 0; i < listCache.roles.length; i++) {
            if (normStr(listCache.roles[i].code).toUpperCase() === code) return listCache.roles[i];
        }
        return null;
    }

    function peopleForRole(role) {
        if (!role) return [];
        return (listCache.rows || []).filter(function (row) {
            return personMatchesRole(row, role);
        });
    }

    function administrationGroupsFromLists(roles, rows) {
        if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
            return window.ms365TenantSettingsAdminRolesAndAdminToGroups(roles, rows);
        }
        return (Array.isArray(roles) ? roles : []).map(function (r) {
            return {
                code: r.code || '',
                name: r.name || '',
                people: (Array.isArray(rows) ? rows : [])
                    .filter(function (row) {
                        return personMatchesRole(row, r);
                    })
                    .map(function (row) {
                        const person = { name: row.name || '', email: row.email || '' };
                        if (row.defaultKey) person.defaultKey = row.defaultKey;
                        return person;
                    })
            };
        });
    }

    function persistLists() {
        const settings = loadTenantSettings() || {};
        settings.admin = (listCache.rows || []).map(function (r) {
            const row = { role: r.role || '', name: r.name || '', email: r.email || '' };
            if (r.defaultKey) row.defaultKey = r.defaultKey;
            return row;
        });
        settings.adminRoles = (listCache.roles || []).map(function (r) {
            return { code: r.code || '', name: r.name || '' };
        });
        settings.administration = administrationGroupsFromLists(listCache.roles || [], listCache.rows || []);
        if (typeof window.ms365TenantSettingsSave === 'function') {
            window.ms365TenantSettingsSave(settings);
        }
        readLists();
    }

    function readLists() {
        const settings = loadTenantSettings();
        const admin = settings && Array.isArray(settings.admin) ? settings.admin : [];
        const rolesIn = settings && Array.isArray(settings.adminRoles) ? settings.adminRoles : [];
        const members = [];
        const direktion = [];
        const rows = [];
        const seenM = new Set();
        const seenD = new Set();
        admin.forEach(function (row) {
            const role = normStr(row && (row.role || row.rolle || row.title));
            const name = normStr(row && row.name);
            const email = normEmail(row && row.email);
            const defaultKey = normStr(row && row.defaultKey);
            const rec = { role: role, name: name, email: email };
            if (defaultKey) rec.defaultKey = defaultKey;
            rows.push(rec);
            if (email && email.indexOf('@') !== -1 && !seenM.has(email)) {
                seenM.add(email);
                members.push(email);
            }
            if (!isDirektionRole(role) && !isDirektionRole(defaultKey)) return;
            if (!email || email.indexOf('@') === -1 || seenD.has(email)) return;
            seenD.add(email);
            direktion.push(email);
        });
        let roles = rolesIn.map(function (r) {
            return { code: normStr(r && r.code).toUpperCase(), name: normStr(r && r.name) };
        }).filter(function (r) {
            return r.code || r.name;
        });
        if (typeof window.ms365TenantSettingsNormalizeAdminRoles === 'function') {
            roles = window.ms365TenantSettingsNormalizeAdminRoles(roles, rows);
        }
        listCache = { members: members, direktion: direktion, rows: rows, roles: roles };
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
        const line = document.getElementById('slgVerwaltungLine');
        if (line) line.textContent = matchedGroupId ? 'Gematcht: ' + matchedGroupId : 'Noch kein Match';
        const groupBtn = document.querySelector('#slgListItems [data-vw-kind="group"]');
        if (groupBtn) groupBtn.setAttribute('aria-current', activeView === 'group' ? 'true' : 'false');
        renderRoleList();
        renderContactBox(
            document.getElementById('slgContactPreview'),
            'Keine Einträge in der Verwaltungsliste.',
            40
        );
        updateMismatchUi();
    }

    function startCellEdit(td, initialValue, onCommit) {
        const prevText = String(initialValue ?? '');
        const input = document.createElement('input');
        input.type = 'text';
        input.value = prevText;
        input.style.width = '100%';
        input.style.font = 'inherit';
        input.style.boxSizing = 'border-box';
        td.replaceChildren(input);
        input.focus();
        input.select();
        const commit = function () {
            onCommit(normStr(input.value));
        };
        const cancel = function () {
            onCommit(prevText, { cancelled: true });
        };
        input.addEventListener('keydown', function (e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                commit();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                cancel();
            }
        });
        input.addEventListener('blur', commit);
    }

    function rolePassesFilter(role) {
        const q = normStr(listFilter).toLowerCase();
        if (!q) return true;
        const people = peopleForRole(role);
        const blob = [role.name, role.code]
            .concat(
                people.map(function (p) {
                    return (p.name || '') + ' ' + (p.email || '');
                })
            )
            .join(' ')
            .toLowerCase();
        return blob.indexOf(q) !== -1;
    }

    function renderRoleList() {
        const host = document.getElementById('vwRoleList');
        const summary = document.getElementById('vwRoleListSummary');
        if (!host) return;
        host.replaceChildren();
        const roles = (listCache.roles || []).filter(rolePassesFilter);
        if (summary) {
            summary.textContent =
                roles.length === (listCache.roles || []).length
                    ? roles.length + ' Rolle(n)'
                    : roles.length + ' von ' + String((listCache.roles || []).length) + ' Rollen';
        }
        if (!roles.length) {
            const li = document.createElement('li');
            const p = document.createElement('p');
            p.className = 'muted';
            p.style.margin = '10px 12px';
            p.textContent = 'Keine Rollen – „+ Rolle“ oder Standardrollen.';
            li.appendChild(p);
            host.appendChild(li);
            return;
        }
        roles.forEach(function (role) {
            const people = peopleForRole(role);
            const li = document.createElement('li');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'slg-side-btn';
            btn.setAttribute('data-vw-kind', 'role');
            btn.setAttribute('data-vw-role-code', role.code || '');
            const on = activeView === 'role' && normStr(activeRoleCode).toUpperCase() === normStr(role.code).toUpperCase();
            btn.setAttribute('aria-current', on ? 'true' : 'false');
            const nPeople = people.length;
            btn.innerHTML =
                '<span class="slg-side-main">' +
                '<span class="slg-side-title">' +
                escapeHtml(role.name || role.code || 'Rolle') +
                '</span>' +
                '<span class="muted slg-side-meta"><code>' +
                escapeHtml(role.code || '') +
                '</code></span></span>' +
                '<span class="slg-side-count"><span class="n">' +
                String(nPeople) +
                '</span><span class="l">' +
                (nPeople === 1 ? 'Person' : 'Personen') +
                '</span></span>';
            btn.addEventListener('click', function () {
                setActiveView('role', role.code);
            });
            li.appendChild(btn);
            host.appendChild(li);
        });
    }

    function setActiveView(view, roleCode) {
        activeView = view === 'role' ? 'role' : 'group';
        activeRoleCode = activeView === 'role' ? String(roleCode || '') : '';
        const groupPanel = document.getElementById('vwGroupPanel');
        const rolePanel = document.getElementById('vwRolePanel');
        const headActions = document.getElementById('vwGroupHeadActions');
        const title = document.getElementById('slgDetailTitle');
        const sub = document.getElementById('slgDetailSubtitle');
        if (groupPanel) groupPanel.style.display = activeView === 'group' ? '' : 'none';
        if (rolePanel) rolePanel.style.display = activeView === 'role' ? '' : 'none';
        if (headActions) headActions.style.display = activeView === 'group' ? '' : 'none';
        if (activeView === 'group') {
            if (title) title.textContent = 'Sammelgruppe Verwaltung';
            if (sub) sub.textContent = matchedGroupId ? 'Gematchte Microsoft‑365‑Gruppe' : 'Gruppe matchen oder anlegen';
        } else {
            const role = getActiveRole();
            if (title) title.textContent = role && role.name ? role.name : 'Rolle';
            if (sub) sub.textContent = 'Personen dieser Rolle pflegen';
            fillRoleForm();
            renderRolePeopleTable();
        }
        updateLeftListUi();
    }

    function fillRoleForm() {
        const role = getActiveRole();
        const inpN = document.getElementById('vwRoleName');
        const inpC = document.getElementById('vwRoleCode');
        if (inpN) inpN.value = role ? role.name || '' : '';
        if (inpC) inpC.value = role ? role.code || '' : '';
    }

    function renderRolePeopleTable() {
        const tbody = document.getElementById('vwRolePeopleBody');
        if (!tbody) return;
        tbody.replaceChildren();
        const role = getActiveRole();
        const people = peopleForRole(role);
        if (!people.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 3;
            td.style.color = '#6c757d';
            td.textContent = 'Keine Personen – „+ Person“.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            return;
        }
        people.forEach(function (person) {
            const globalIdx = listCache.rows.indexOf(person);
            const tr = document.createElement('tr');
            const tdName = document.createElement('td');
            tdName.textContent = person.name || '';
            tdName.title = 'Doppelklick zum Bearbeiten';
            tdName.addEventListener('dblclick', function () {
                startCellEdit(tdName, person.name, function (next, meta) {
                    if (globalIdx < 0 || !listCache.rows[globalIdx]) return renderRolePeopleTable();
                    if (!(meta && meta.cancelled)) listCache.rows[globalIdx].name = next;
                    persistLists();
                    updateLeftListUi();
                    renderRolePeopleTable();
                });
            });
            const tdEmail = document.createElement('td');
            tdEmail.textContent = person.email || '';
            tdEmail.title = 'Doppelklick zum Bearbeiten';
            tdEmail.addEventListener('dblclick', function () {
                startCellEdit(tdEmail, person.email, function (next, meta) {
                    if (globalIdx < 0 || !listCache.rows[globalIdx]) return renderRolePeopleTable();
                    if (!(meta && meta.cancelled)) listCache.rows[globalIdx].email = next.toLowerCase();
                    persistLists();
                    updateLeftListUi();
                    renderRolePeopleTable();
                });
            });
            const tdAction = document.createElement('td');
            tdAction.className = 'action-cell';
            const btnDel = document.createElement('button');
            btnDel.type = 'button';
            btnDel.className = 'mini-btn';
            btnDel.textContent = '✕';
            btnDel.title = 'Person löschen';
            btnDel.addEventListener('click', function () {
                if (globalIdx < 0) return;
                listCache.rows.splice(globalIdx, 1);
                persistLists();
                updateLeftListUi();
                renderRolePeopleTable();
                toast('Person entfernt.');
            });
            tdAction.appendChild(btnDel);
            tr.append(tdName, tdEmail, tdAction);
            tbody.appendChild(tr);
        });
    }

    async function addRole() {
        const name = await (typeof window.ms365AppDialogPrompt === 'function'
            ? window.ms365AppDialogPrompt('Bezeichnung der neuen Rolle', '', {
                  title: 'Rolle anlegen',
                  inputLabel: 'Rolle'
              })
            : Promise.resolve(window.prompt('Bezeichnung der neuen Rolle')));
        const label = normStr(name);
        if (!label) return;
        const exists = (listCache.roles || []).some(function (r) {
            return normStr(r.name).toLowerCase() === label.toLowerCase();
        });
        if (exists) {
            toast('Diese Rolle gibt es bereits.');
            return;
        }
        const code = uniqueRoleCode(
            label,
            (listCache.roles || []).map(function (r) {
                return r.code;
            })
        );
        listCache.roles.push({ code: code, name: label });
        persistLists();
        setActiveView('role', code);
        toast('Rolle angelegt.');
    }

    function addDefaultRoles() {
        const defaults =
            typeof window.ms365TenantSettingsDefaultAdminRoles === 'function'
                ? window.ms365TenantSettingsDefaultAdminRoles()
                : [];
        const seen = new Set(
            (listCache.roles || []).map(function (r) {
                return String(r.code || '').toLowerCase();
            })
        );
        let added = 0;
        defaults.forEach(function (d) {
            const k = String(d.code || '').toLowerCase();
            if (k && seen.has(k)) return;
            if (k) seen.add(k);
            listCache.roles.push({ code: d.code, name: d.name });
            added += 1;
        });
        persistLists();
        updateLeftListUi();
        toast(added ? added + ' Standardrolle(n) ergänzt.' : 'Standardrollen sind bereits vorhanden.');
    }

    function saveActiveRole() {
        const role = getActiveRole();
        if (!role) {
            toast('Keine Rolle ausgewählt.');
            return;
        }
        const inpN = document.getElementById('vwRoleName');
        const inpC = document.getElementById('vwRoleCode');
        const nextName = inpN ? normStr(inpN.value) : role.name;
        const nextCode = inpC ? normStr(inpC.value).toUpperCase() : role.code;
        if (!nextName) {
            toast('Bitte eine Bezeichnung eingeben.');
            return;
        }
        const oldName = role.name;
        if (nextName !== oldName && typeof window.ms365TenantSettingsRenameAdminRole === 'function') {
            const renamed = window.ms365TenantSettingsRenameAdminRole(listCache.roles, listCache.rows, oldName, nextName);
            listCache.roles = renamed.roles;
            listCache.rows = renamed.admin;
        } else {
            role.name = nextName;
        }
        const found =
            (listCache.roles || []).find(function (r) {
                return (
                    normStr(r.name).toLowerCase() === nextName.toLowerCase() ||
                    normStr(r.code).toUpperCase() === normStr(activeRoleCode).toUpperCase()
                );
            }) || role;
        const clash = (listCache.roles || []).some(function (r) {
            return r !== found && normStr(r.code).toUpperCase() === nextCode;
        });
        if (nextCode && !clash) found.code = nextCode;
        else if (nextCode && clash) toast('Kürzel bereits vergeben – Bezeichnung gespeichert, Kürzel unverändert.');
        persistLists();
        const still = (listCache.roles || []).find(function (r) {
            return normStr(r.name).toLowerCase() === nextName.toLowerCase();
        });
        setActiveView('role', still ? still.code : found.code);
        toast('Rolle gespeichert.');
    }

    async function deleteActiveRole() {
        const role = getActiveRole();
        if (!role) return;
        const people = peopleForRole(role);
        if (people.length) {
            toast('Rolle ist noch ' + people.length + ' Person(en) zugeordnet. Zuerst Personen entfernen oder umbenennen.');
            return;
        }
        const ok = await dlgConfirm('Rolle „' + (role.name || role.code) + '“ wirklich löschen?', {
            title: 'Rolle löschen',
            okText: 'Löschen',
            cancelText: 'Abbrechen'
        });
        if (!ok) return;
        listCache.roles = (listCache.roles || []).filter(function (r) {
            return normStr(r.code).toUpperCase() !== normStr(role.code).toUpperCase();
        });
        persistLists();
        setActiveView('group');
        toast('Rolle gelöscht.');
    }

    function addPersonToActiveRole() {
        const role = getActiveRole();
        if (!role) {
            toast('Bitte zuerst eine Rolle wählen.');
            return;
        }
        listCache.rows.push({
            role: role.name || role.code,
            name: '',
            email: '',
            defaultKey: role.name || ''
        });
        persistLists();
        renderRolePeopleTable();
        updateLeftListUi();
        toast('Personenzeile hinzugefügt.');
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
            p.textContent = 'Keine Direktion‑Besitzer in den Stammdaten gefunden.';
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
            let joinEmails = emails;
            let leaveEmails = [];
            if (typeof gug().fetchGroupMembers === 'function') {
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
                const M = window.ms365MembershipReconcile;
                if (M && typeof M.diffMemberships === 'function') {
                    const diff = M.diffMemberships(emails, current);
                    joinEmails = diff.onlyLocal;
                    leaveEmails = diff.onlyGraph;
                } else {
                    const lc = window.ms365StudentClassLifecycle;
                    if (lc && typeof lc.reconcileSammelgruppe === 'function') {
                        const rec = lc.reconcileSammelgruppe(emails, current);
                        joinEmails = rec.join;
                        leaveEmails = rec.leave;
                    }
                }
                appendSyncLog(
                    'Abgleich mit Verwaltungsliste: +' + joinEmails.length + ' / −' + leaveEmails.length + '.',
                    ''
                );
            }
            if (joinEmails.length) {
                const r = await gug().syncEmailsToGroup(token, gid, joinEmails, 'Verwaltung', appendSyncLog);
                appendSyncLog('Aufnehmen: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            }
            if (leaveEmails.length && typeof gug().removeEmailsFromGroup === 'function') {
                const r = await gug().removeEmailsFromGroup(token, gid, leaveEmails, 'Verwaltung', appendSyncLog);
                appendSyncLog('Entfernen: ' + r.ok + ' OK, übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            }
            if (!joinEmails.length && !leaveEmails.length) {
                appendSyncLog('Keine Änderungen gegenüber der Verwaltungsliste.', 'ok');
            }
            await ensureOwners(token, gid);
            live().invalidateMembership();
            await live().loadMembers();
            await refreshGraphMemberCounts();
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
        gd().clearSearchResults();
        live().setMatchedMode(!!matchedGroupId);
        live().fillForm(matchedGroupId ? { id: matchedGroupId } : null);
        updateLeftListUi();
        gd().setTab('general');
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

    function onClick(id, fn) {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', fn);
    }

    function mountDetail() {
        gd().mount('#groupDetailHost', {
            title: 'Sammelgruppe Verwaltung',
            searchPlaceholder: 'z. B. verwaltung oder @schule.at',
            unmatchedCreateHint:
                'Legt eine Microsoft 365‑Gruppe (Unified) an. Optional auch als Team bereitstellen.',
            membersUnmatchedHint:
                'Mitglieder kommen aus der Verwaltungsliste. Nach dem Match können Sie live verwalten und die Liste synchronisieren.',
            membersUnmatchedTitle: 'Vorschau Verwaltungsliste (erste 30)',
            membersMatchedHint:
                'Live aus Microsoft Graph. „Mitglieder synchronisieren“ gleicht die Gruppe mit der Verwaltungsliste ab (fehlende hinzufügen, nicht gelistete entfernen).',
            features: { syncMembers: true, membershipReview: true },
            ids: {
                wrap: 'vwGroupPanel',
                headActions: 'vwGroupHeadActions',
                afterWrap: 'vwRolePanel'
            },
            live: {
                toast: toast,
                dlgConfirm: dlgConfirm,
                getGroupId: getActiveMatchedId,
                ensureDirektionOwners: function (token, gid) {
                    if (!(listCache.direktion && listCache.direktion.length)) {
                        throw new Error('Keine Direktion‑Adressen in den Stammdaten.');
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
                    void refreshGraphMemberCounts();
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
        onClick('slgBtnReloadLists', function () {
            readLists();
            updateLeftListUi();
            renderOwnerPreview();
            renderMemberPreview();
            if (activeView === 'role') {
                fillRoleForm();
                renderRolePeopleTable();
            }
            toast('Listen neu eingelesen.');
        });
        onClick('vwBtnAddRole', function () {
            addRole();
        });
        onClick('vwBtnDefaultRoles', function () {
            addDefaultRoles();
        });
        onClick('vwBtnSaveRole', function () {
            saveActiveRole();
        });
        onClick('vwBtnDeleteRole', function () {
            deleteActiveRole();
        });
        onClick('vwBtnAddPerson', function () {
            addPersonToActiveRole();
        });
        const groupBtn = document.querySelector('#slgListItems [data-vw-kind="group"]');
        if (groupBtn) {
            groupBtn.addEventListener('click', function () {
                setActiveView('group');
            });
        }
        const filter = document.getElementById('vwListFilter');
        if (filter) {
            filter.addEventListener('input', function () {
                listFilter = filter.value || '';
                renderRoleList();
            });
        }
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
        initMembershipReview();
        readLists();
        loadState();
        updateLeftListUi();
        renderOwnerPreview();
        renderMemberPreview();
        wire();
        setActiveView('group');
        if (!getActiveMatchedId()) {
            live().setMatchedMode(false);
            applyCreateDefaults();
        } else {
            void refreshGraphMemberCounts();
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
