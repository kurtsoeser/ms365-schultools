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

    /** @type {{ schulleitung: string|null, verwaltung: string|null, custom: Record<string, string|null> }} */
    let matchedGroupIds = { schulleitung: null, verwaltung: null, custom: {} };

    /** @type {{ members: string[], direktion: string[], rows: object[], roles: object[], audienceGroups: object[], memberships: object[] }} */
    let listCache = { members: [], direktion: [], rows: [], roles: [], audienceGroups: [], memberships: [] };

    /** @type {string} schulleitung | verwaltung | audience:va-* | role */
    let activeView = 'verwaltung';
    /** @type {string} */
    let activeRoleCode = '';
    let listFilter = '';
    /** @type {Record<string, number>} */
    let graphMemberCounts = {};
    let countsFetchGen = 0;
    /** @type {ReturnType<import('../../shared/membership-review-ui.js').createMembershipReview>|null} */
    let membershipReview = null;
    /** @type {{ refresh: function, resetPreview?: function }|null} */
    let splitMigrationUi = null;

    function audienceApi() {
        return window.ms365AdministrationAudience || {};
    }

    function rolePolicyApi() {
        return audienceApi();
    }

    function applyRolePolicyDefaults(role) {
        const api = rolePolicyApi();
        if (!role || typeof api.adminRolePolicyFromRecord !== 'function') return role || {};
        return Object.assign({}, role, api.adminRolePolicyFromRecord(role));
    }

    function adminGroupMetaForRole(roleName, roleCode, adminGroups) {
        const rl = normStr(roleName).toLowerCase();
        const rc = normStr(roleCode).toLowerCase();
        for (let i = 0; i < (adminGroups || []).length; i++) {
            const g = adminGroups[i];
            if (!g) continue;
            const gn = normStr(g.name).toLowerCase();
            const gc = normStr(g.code).toLowerCase();
            if ((rl && gn === rl) || (rc && gc === rc)) return g;
        }
        return null;
    }

    function mergeRoleFromAdminGroup(role, adminGroups) {
        const g = adminGroupMetaForRole(role.name, role.code, adminGroups);
        if (!g) return applyRolePolicyDefaults(role);
        return applyRolePolicyDefaults(
            Object.assign({}, role, {
                tier: role.tier || g.tier || '',
                seatMode: g.seatMode || role.seatMode,
                m365ResourceType: g.m365ResourceType || role.m365ResourceType,
                m365ResourceId: g.m365ResourceId || role.m365ResourceId,
                m365ResourceEmail: g.m365ResourceEmail || role.m365ResourceEmail,
                m365ResourceLabel: g.m365ResourceLabel || role.m365ResourceLabel
            })
        );
    }

    function serializeRoleForSettings(r) {
        const row = { code: r.code || '', name: r.name || '' };
        if (r.tier) row.tier = r.tier;
        if (r.seatMode) row.seatMode = r.seatMode;
        if (r.m365ResourceType && r.m365ResourceType !== 'none') row.m365ResourceType = r.m365ResourceType;
        if (r.m365ResourceId) row.m365ResourceId = r.m365ResourceId;
        if (r.m365ResourceEmail) row.m365ResourceEmail = r.m365ResourceEmail;
        if (r.m365ResourceLabel) row.m365ResourceLabel = r.m365ResourceLabel;
        return row;
    }

    function roleSeatLabel(role) {
        const api = rolePolicyApi();
        if (typeof api.adminRoleSeatModeLabel === 'function') {
            return api.adminRoleSeatModeLabel(role && role.seatMode);
        }
        return '';
    }

    function roleM365Short(role) {
        const api = rolePolicyApi();
        if (typeof api.adminRoleM365KindShortLabel === 'function') {
            return api.adminRoleM365KindShortLabel(role && role.m365ResourceType);
        }
        return '';
    }

    function updateRolePeopleHint() {
        const el = document.getElementById('vwRolePeopleHint');
        const role = getActiveRole();
        if (!el || !role) return;
        const api = rolePolicyApi();
        const single =
            typeof api.normalizeAdminRoleSeatMode === 'function' &&
            api.normalizeAdminRoleSeatMode(role.seatMode, role) === 'single';
        el.textContent = single
            ? 'Genau eine Person für diese Rolle. Name und E‑Mail per Doppelklick bearbeiten.'
            : 'Mehrere Personen möglich (z. B. zwei Sekretariate). Name und E‑Mail per Doppelklick bearbeiten.';
    }

    function updateRoleM365Ui() {
        const kindEl = document.getElementById('vwRoleM365Kind');
        const row = document.getElementById('vwRoleM365PickRow');
        const summary = document.getElementById('vwRoleM365Summary');
        const role = getActiveRole();
        const api = rolePolicyApi();
        const kind =
            kindEl && typeof api.normalizeAdminRoleM365Kind === 'function'
                ? api.normalizeAdminRoleM365Kind(kindEl.value)
                : 'none';
        if (row) row.hidden = kind === 'none';
        if (!summary) return;
        if (!role || kind === 'none') {
            summary.textContent = '–';
            return;
        }
        const label = normStr(role.m365ResourceLabel);
        const email = normStr(role.m365ResourceEmail);
        const id = normStr(role.m365ResourceId);
        if (label || email) {
            summary.textContent = [label, email].filter(Boolean).join(' · ');
        } else if (id) {
            summary.textContent = 'Objekt-ID: ' + id;
        } else {
            summary.textContent =
                kind === 'sharedMailbox'
                    ? 'Noch kein Postfach gewählt – „In Entra wählen“ oder nach dem Speichern manuell per E‑Mail ergänzen.'
                    : 'Noch keine Gruppe gewählt – „In Entra wählen“.';
        }
    }

    function readRolePolicyFromForm(role) {
        const seatEl = document.getElementById('vwRoleSeatMode');
        const kindEl = document.getElementById('vwRoleM365Kind');
        const api = rolePolicyApi();
        if (seatEl && typeof api.normalizeAdminRoleSeatMode === 'function') {
            role.seatMode = api.normalizeAdminRoleSeatMode(seatEl.value, role);
        }
        const prevKind = role.m365ResourceType;
        const nextKind =
            kindEl && typeof api.normalizeAdminRoleM365Kind === 'function'
                ? api.normalizeAdminRoleM365Kind(kindEl.value)
                : 'none';
        role.m365ResourceType = nextKind;
        if (nextKind === 'none') {
            role.m365ResourceId = '';
            role.m365ResourceEmail = '';
            role.m365ResourceLabel = '';
        } else if (prevKind !== nextKind) {
            role.m365ResourceId = '';
            role.m365ResourceEmail = '';
            role.m365ResourceLabel = '';
        }
    }

    async function pickRoleM365Resource() {
        const role = getActiveRole();
        if (!role) return;
        const kindEl = document.getElementById('vwRoleM365Kind');
        const api = rolePolicyApi();
        const kind =
            kindEl && typeof api.normalizeAdminRoleM365Kind === 'function'
                ? api.normalizeAdminRoleM365Kind(kindEl.value)
                : 'none';
        if (kind === 'none') {
            toast('Bitte zuerst M365‑Gruppe oder Postfach wählen.');
            return;
        }
        try {
            if (kind === 'group') {
                const mod = await import('../../shared/entra-group-picker.js');
                const picked = await mod.pickEntraGroup({
                    title: 'M365‑Gruppe für Rolle „' + (role.name || role.code) + '“'
                });
                if (!picked) return;
                role.m365ResourceType = 'group';
                role.m365ResourceId = picked.id || '';
                role.m365ResourceEmail = normEmail(picked.mail || picked.email || '');
                role.m365ResourceLabel = normStr(picked.displayName || picked.name || picked.mail);
            } else {
                const mod = await import('../../shared/entra-user-picker.js');
                const picked = await mod.pickEntraUser({
                    title: 'Postfach / Konto für Rolle „' + (role.name || role.code) + '“',
                    hint: 'Freigegebene Postfächer erscheinen oft als Benutzer mit der Postfach‑Adresse.'
                });
                if (!picked) return;
                role.m365ResourceType = 'sharedMailbox';
                role.m365ResourceId = picked.id || '';
                role.m365ResourceEmail = normEmail(picked.mail || picked.userPrincipalName || '');
                role.m365ResourceLabel = normStr(picked.displayName || picked.mail);
            }
            persistLists();
            fillRoleForm();
            updateRoleM365Ui();
            toast('Microsoft‑365‑Bezug gespeichert.');
        } catch (e) {
            toast('Auswahl fehlgeschlagen: ' + (e && e.message ? e.message : e));
        }
    }

    function clearRoleM365Resource() {
        const role = getActiveRole();
        if (!role) return;
        role.m365ResourceId = '';
        role.m365ResourceEmail = '';
        role.m365ResourceLabel = '';
        const kindEl = document.getElementById('vwRoleM365Kind');
        if (kindEl) kindEl.value = 'none';
        role.m365ResourceType = 'none';
        persistLists();
        fillRoleForm();
        updateRoleM365Ui();
        toast('M365‑Bezug entfernt.');
    }

    function isAudienceGroupView(view) {
        const v = String(view || activeView || '');
        return v === 'schulleitung' || v === 'verwaltung' || v.indexOf('audience:') === 0;
    }

    function isSammelView() {
        return isAudienceGroupView(activeView);
    }

    function audienceGroupIdFromView(view) {
        const v = String(view || '').trim();
        if (v.indexOf('audience:') === 0) return v.slice('audience:'.length);
        return normalizeSammelKind(v);
    }

    function normalizeSammelKind(view) {
        const v = String(view || '').trim().toLowerCase();
        if (v === 'schulleitung') return 'schulleitung';
        if (v === 'verwaltung' || v === 'group') return 'verwaltung';
        return '';
    }

    function sammelMeta(kind) {
        const id = audienceGroupIdFromView(kind);
        if (id === 'schulleitung') {
            return {
                kind: 'schulleitung',
                label: 'Schulleitung',
                setupField: 'schulleitungGroupId',
                reviewKey: 'schulleitung',
                audienceGroupId: 'schulleitung'
            };
        }
        if (id === 'verwaltung') {
            return {
                kind: 'verwaltung',
                label: 'Verwaltung (Personal)',
                setupField: 'verwaltungGroupId',
                reviewKey: 'verwaltung',
                audienceGroupId: 'verwaltung'
            };
        }
        const g = (listCache.audienceGroups || []).find(function (x) {
            return x && x.id === id;
        });
        return {
            kind: 'audience:' + id,
            label: (g && g.label) || id,
            setupField: null,
            reviewKey: 'audience-' + id,
            audienceGroupId: id
        };
    }

    function splitEmailsByTier() {
        const api = audienceApi();
        if (listCache.memberships.length && typeof api.splitAdminEmailsByAudienceTier === 'function') {
            return api.splitAdminEmailsByAudienceTier(listCache.rows, listCache.roles, listCache.memberships);
        }
        if (typeof api.splitAdminEmailsByAudienceTier === 'function') {
            return api.splitAdminEmailsByAudienceTier(listCache.rows, listCache.roles);
        }
        return {
            schulleitung: listCache.direktion || [],
            verwaltung: listCache.members || []
        };
    }

    function emailsForSammelKind(kind) {
        const api = audienceApi();
        const groupId = audienceGroupIdFromView(kind);
        if (listCache.memberships.length && typeof api.collectEmailsForAudienceGroup === 'function') {
            return api.collectEmailsForAudienceGroup(listCache.memberships, groupId);
        }
        const split = splitEmailsByTier();
        return groupId === 'schulleitung' ? split.schulleitung || [] : split.verwaltung || [];
    }

    function activeSammelMeta() {
        return isSammelView() ? sammelMeta(activeView) : sammelMeta('verwaltung');
    }

    function renderCustomAudienceSidebar() {
        const ul = document.getElementById('vwCustomAudienceList');
        if (!ul) return;
        ul.replaceChildren();
        const custom = (listCache.audienceGroups || []).filter(function (g) {
            return g && !g.builtin;
        });
        if (!custom.length) return;
        custom.forEach(function (g) {
            const li = document.createElement('li');
            const view = 'audience:' + g.id;
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'slg-side-btn slg-side-btn--audience';
            btn.setAttribute('data-vw-kind', view);
            btn.setAttribute('aria-current', activeView === view ? 'true' : 'false');
            const count = emailsForSammelKind(view).length;
            btn.innerHTML =
                '<span class="slg-side-main"><span class="slg-side-title"><i class="bi bi-people" aria-hidden="true"></i> ' +
                escapeHtml(g.label) +
                '</span></span><span class="slg-side-count"><span class="n">' +
                String(count) +
                '</span><span class="l">Liste</span></span>';
            btn.addEventListener('click', function () {
                setActiveView(view);
            });
            li.appendChild(btn);
            ul.appendChild(li);
        });
    }

    function graphCountFor(groupId) {
        const id = String(groupId || '').trim();
        if (!id) return null;
        const n = graphMemberCounts[id];
        return typeof n === 'number' && n >= 0 ? n : null;
    }

    async function refreshGraphMemberCounts() {
        const ids = [matchedGroupIds.schulleitung, matchedGroupIds.verwaltung].filter(Boolean);
        if (!ids.length) {
            updateMismatchUi();
            return;
        }
        const gen = ++countsFetchGen;
        try {
            const token = await getGraphToken();
            if (gen !== countsFetchGen) return;
            for (let i = 0; i < ids.length; i++) {
                const gid = ids[i];
                const n = await gug().fetchGroupMemberCount(token, gid);
                if (typeof n === 'number' && n >= 0) graphMemberCounts[gid] = n;
            }
            if (gen !== countsFetchGen) return;
            updateMismatchUi();
        } catch {
            updateMismatchUi();
        }
    }

    function paintSammelKindCounts(kind) {
        const meta = sammelMeta(kind);
        const listN = emailsForSammelKind(kind).length;
        const gid = matchedGroupIds[kind];
        const groupN = graphCountFor(gid);
        const prefix = kind === 'schulleitung' ? 'Schulleitung' : 'Verwaltung';
        const wrap = document.getElementById('slg' + prefix + 'Counts');
        const listEl = document.getElementById('slg' + prefix + 'Count');
        const groupEl = document.getElementById('slg' + prefix + 'GroupCount');
        const lineEl = document.getElementById('slg' + prefix + 'Line');
        if (listEl) listEl.textContent = String(listN);
        if (groupEl) groupEl.textContent = gid ? (groupN === null ? '–' : String(groupN)) : '–';
        if (lineEl) {
            lineEl.textContent = gid ? 'Gematcht: ' + gid : 'Noch kein Match';
        }
        if (!wrap) return;
        wrap.classList.remove('is-match', 'is-mismatch');
        const known = gid && groupN !== null;
        if (known) {
            const same = listN === groupN;
            wrap.classList.add(same ? 'is-match' : 'is-mismatch');
            wrap.title = same
                ? meta.label + ': Liste und Gruppe je ' + listN + ' – stimmt überein.'
                : meta.label + ': Liste ' + listN + ' · Gruppe ' + groupN + ' Mitglieder.';
        } else {
            wrap.title = gid
                ? meta.label + ': ' + listN + ' E-Mails in der Liste (Graph lädt …).'
                : meta.label + ': ' + listN + ' E-Mails. Noch keine Gruppe gematcht.';
        }
    }

    function paintAllSammelCounts() {
        paintSammelKindCounts('schulleitung');
        paintSammelKindCounts('verwaltung');
    }

    function updateMismatchUi() {
        paintAllSammelCounts();
        if (!membershipReview) return;
        if (!isSammelView()) {
            membershipReview.updateMismatchBar([]);
            return;
        }
        const meta = activeSammelMeta();
        const gid = getActiveMatchedId();
        const listN = emailsForSammelKind(activeView).length;
        const groupN = graphCountFor(gid);
        if (gid && groupN !== null && listN !== groupN) {
            membershipReview.updateMismatchBar([
                {
                    key: meta.reviewKey,
                    label: meta.label,
                    listN: listN,
                    groupN: groupN,
                    gid: gid
                }
            ]);
        } else {
            membershipReview.updateMismatchBar([]);
        }
    }

    function getMigrationState() {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const su = window.ms365AppDataV2.getSetup() || {};
                return su.verwaltungSplitMigration || { completedAt: null, skippedAt: null };
            }
        } catch {
            /* ignore */
        }
        return { completedAt: null, skippedAt: null };
    }

    function patchMigration(patch) {
        if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.patchSetup !== 'function') return;
        window.ms365AppDataV2.patchSetup({
            verwaltungSplitMigration: Object.assign({}, getMigrationState(), patch || {})
        });
    }

    function refreshSplitMigration() {
        if (splitMigrationUi && typeof splitMigrationUi.refresh === 'function') {
            splitMigrationUi.refresh();
        }
    }

    function initSplitMigration() {
        const Ui = window.ms365VerwaltungSplitMigrationUi;
        if (!Ui || typeof Ui.mountVerwaltungSplitMigration !== 'function') return;
        splitMigrationUi = Ui.mountVerwaltungSplitMigration({
            getMigrationState: getMigrationState,
            getSplitEmails: splitEmailsByTier,
            getMatchedGroupIds: function () {
                return {
                    schulleitung: matchedGroupIds.schulleitung,
                    verwaltung: matchedGroupIds.verwaltung
                };
            },
            getGraphToken: getGraphToken,
            graphUnifiedGroups: gug(),
            patchMigration: patchMigration,
            ensureOwners: ensureOwners,
            toast: toast,
            dlgConfirm: dlgConfirm,
            onFocusSchulleitung: function () {
                setActiveView('schulleitung');
            },
            onAfterSuccess: function () {
                void refreshGraphMemberCounts();
                updateLeftListUi();
                if (getActiveMatchedId()) void live().loadGroup({ silent: true });
            }
        });
    }

    function initMembershipReview() {
        const R = window.ms365MembershipReviewUi;
        if (!R || typeof R.createMembershipReview !== 'function') return;
        membershipReview = R.createMembershipReview({
            mode: 'sammelgruppe',
            tool: 'verwaltung',
            syncLabel: function () {
                return activeSammelMeta().label;
            },
            getGraphToken: getGraphToken,
            getGroupId: getActiveMatchedId,
            getLocalEmails: function () {
                if (!isSammelView()) return [];
                return emailsForSammelKind(activeView);
            },
            getActiveReviewKey: function () {
                return activeSammelMeta().reviewKey;
            },
            getReviewTitle: function () {
                return 'Mitglieder-Abgleich: ' + activeSammelMeta().label;
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
            syncGraphMemberCount: function (gid, count) {
                const id = String(gid || '').trim();
                if (!id) return;
                const n = typeof count === 'number' ? count : -1;
                if (n < 0) return;
                graphMemberCounts[id] = n;
                updateMismatchUi();
            },
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
                            await membershipReview.loadReview(activeSammelMeta().reviewKey);
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
                onlyLocalTitle: 'Nur in der Stammdaten-Liste (Zielgruppe)',
                onlyLocalHint: 'In den Stammdaten (Zielgruppe), aber nicht in der Microsoft-365-Gruppe.',
                onlyGraphTitle: 'Nur in der Microsoft-365-Gruppe',
                onlyGraphHint:
                    'In der Gruppe online, aber nicht in der Zielgruppe der Stammdaten – Rolle/Zielgruppe prüfen.'
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
        settings.adminRoles = (listCache.roles || []).map(serializeRoleForSettings);
        settings.administration = administrationGroupsFromLists(listCache.roles || [], listCache.rows || []);
        if (typeof window.ms365TenantSettingsSave === 'function') {
            window.ms365TenantSettingsSave(settings);
        }
        readLists();
    }

    function readLists() {
        const settings = loadTenantSettings();
        const data = (settings && settings.data) || settings || {};
        const admin = data && Array.isArray(data.admin) ? data.admin : [];
        const rolesIn = data && Array.isArray(data.adminRoles) ? data.adminRoles : [];
        const adminGroups = data && Array.isArray(data.administration) ? data.administration : [];
        function metaFromAdminGroup(roleName, roleCode) {
            const g = adminGroupMetaForRole(roleName, roleCode, adminGroups);
            if (!g) {
                const api = audienceApi();
                const tier =
                    typeof api.inferAdminTierForRole === 'function'
                        ? api.inferAdminTierForRole({ name: roleName, code: roleCode })
                        : '';
                return applyRolePolicyDefaults({ name: roleName, code: roleCode, tier: tier });
            }
            return mergeRoleFromAdminGroup(
                {
                    name: roleName,
                    code: roleCode,
                    tier: g.tier || ''
                },
                adminGroups
            );
        }
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
        let roles = rolesIn
            .map(function (r) {
                const name = normStr(r && r.name);
                const code = normStr(r && r.code).toUpperCase();
                const merged = mergeRoleFromAdminGroup(
                    Object.assign({}, metaFromAdminGroup(name, code), {
                        code: code,
                        name: name,
                        tier: (r && r.tier) || metaFromAdminGroup(name, code).tier || ''
                    }),
                    adminGroups
                );
                if (r && r.seatMode) merged.seatMode = r.seatMode;
                if (r && r.m365ResourceType) merged.m365ResourceType = r.m365ResourceType;
                if (r && r.m365ResourceId) merged.m365ResourceId = r.m365ResourceId;
                if (r && r.m365ResourceEmail) merged.m365ResourceEmail = r.m365ResourceEmail;
                if (r && r.m365ResourceLabel) merged.m365ResourceLabel = r.m365ResourceLabel;
                return applyRolePolicyDefaults(merged);
            })
            .filter(function (r) {
                return r.code || r.name;
            });
        if (typeof window.ms365TenantSettingsNormalizeAdminRoles === 'function') {
            roles = window.ms365TenantSettingsNormalizeAdminRoles(roles, rows);
        }
        let audienceGroups = [];
        let memberships = [];
        const audApi = audienceApi();
        if (typeof audApi.ensureAdminAudienceOnSettings === 'function') {
            const ensured = audApi.ensureAdminAudienceOnSettings({
                admin: admin,
                adminRoles: rolesIn,
                administration: adminGroups,
                verwaltungAudienceGroups: data.verwaltungAudienceGroups,
                adminAudienceMemberships: data.adminAudienceMemberships
            });
            audienceGroups = ensured.groups || [];
            memberships = ensured.memberships || [];
        }
        listCache = {
            members: members,
            direktion: direktion,
            rows: rows,
            roles: roles,
            audienceGroups: audienceGroups,
            memberships: memberships
        };
        renderCustomAudienceSidebar();
    }

    function updateLeftListUi() {
        document.querySelectorAll('#slgListItems [data-vw-kind="schulleitung"], #slgListItems [data-vw-kind="verwaltung"]').forEach(
            function (btn) {
                const kind = btn.getAttribute('data-vw-kind');
                btn.setAttribute('aria-current', activeView === kind ? 'true' : 'false');
            }
        );
        renderRoleList();
        renderCustomAudienceSidebar();
        updateMismatchUi();
        refreshSplitMigration();
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
            btn.className = 'slg-side-btn slg-side-btn--role';
            btn.setAttribute('data-vw-kind', 'role');
            btn.setAttribute('data-vw-role-code', role.code || '');
            const on = activeView === 'role' && normStr(activeRoleCode).toUpperCase() === normStr(role.code).toUpperCase();
            btn.setAttribute('aria-current', on ? 'true' : 'false');
            const nPeople = people.length;
            const seat = roleSeatLabel(role);
            const m365 = roleM365Short(role);
            const seatBadge =
                seat === '1 Platz'
                    ? '<span class="vw-role-badge vw-role-badge--single" title="Einzelbesetzung">1×</span>'
                    : '<span class="vw-role-badge" title="Mehrere Personen">Team</span>';
            const m365Badge = m365
                ? '<span class="vw-role-badge vw-role-badge--m365" title="M365-Bezug">' + escapeHtml(m365) + '</span>'
                : '';
            btn.innerHTML =
                '<span class="slg-side-main">' +
                '<span class="slg-side-title">' +
                escapeHtml(role.name || role.code || 'Rolle') +
                '</span>' +
                '<span class="muted slg-side-meta">' +
                seatBadge +
                m365Badge +
                '<code>' +
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
        if (view === 'role') {
            activeView = 'role';
            activeRoleCode = String(roleCode || '');
        } else if (String(view).indexOf('audience:') === 0) {
            activeView = String(view);
            activeRoleCode = '';
        } else {
            const kind = normalizeSammelKind(view);
            activeView = kind || 'verwaltung';
            activeRoleCode = '';
        }
        const groupPanel = document.getElementById('vwGroupPanel');
        const rolePanel = document.getElementById('vwRolePanel');
        const headActions = document.getElementById('vwGroupHeadActions');
        const title = document.getElementById('slgDetailTitle');
        const sub = document.getElementById('slgDetailSubtitle');
        const onSammel = isSammelView();
        if (groupPanel) groupPanel.style.display = onSammel ? '' : 'none';
        if (rolePanel) rolePanel.style.display = activeView === 'role' ? '' : 'none';
        if (headActions) headActions.style.display = onSammel ? '' : 'none';
        if (onSammel) {
            const meta = sammelMeta(activeView);
            const gid = getActiveMatchedId();
            if (title) title.textContent = 'Sammelgruppe ' + meta.label;
            if (sub) {
                sub.textContent = gid
                    ? 'Gematchte Microsoft‑365‑Gruppe · ' + emailsForSammelKind(activeView).length + ' E-Mails in der Liste'
                    : 'Gruppe matchen oder anlegen';
            }
            fillCreateFormFromDraft();
            live().setMatchedMode(!!gid);
            live().fillForm(gid ? { id: gid } : null);
            if (gid) void live().loadGroup({ silent: true });
            else live().loadGroup({ silent: true });
        } else {
            const role = getActiveRole();
            if (title) title.textContent = role && role.name ? role.name : 'Rolle';
            if (sub) {
                const bits = ['Personen & Besetzung'];
                if (role && roleM365Short(role)) bits.push(roleM365Short(role));
                sub.textContent = bits.join(' · ');
            }
            fillRoleForm();
            renderRolePeopleTable();
            updateRolePeopleHint();
            updateRoleM365Ui();
        }
        updateLeftListUi();
    }

    function fillRoleForm() {
        const role = getActiveRole();
        const inpN = document.getElementById('vwRoleName');
        const inpC = document.getElementById('vwRoleCode');
        const seatEl = document.getElementById('vwRoleSeatMode');
        const kindEl = document.getElementById('vwRoleM365Kind');
        const api = rolePolicyApi();
        if (inpN) inpN.value = role ? role.name || '' : '';
        if (inpC) inpC.value = role ? role.code || '' : '';
        if (seatEl && role) {
            seatEl.value =
                typeof api.normalizeAdminRoleSeatMode === 'function' &&
                api.normalizeAdminRoleSeatMode(role.seatMode, role) === 'single'
                    ? 'single'
                    : 'multi';
        }
        if (kindEl && role) {
            kindEl.value =
                typeof api.normalizeAdminRoleM365Kind === 'function'
                    ? api.normalizeAdminRoleM365Kind(role.m365ResourceType)
                    : 'none';
        }
        updateRoleM365Ui();
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
        listCache.roles.push(applyRolePolicyDefaults({ code: code, name: label }));
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
            listCache.roles.push(applyRolePolicyDefaults({ code: d.code, name: d.name }));
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
        readRolePolicyFromForm(found);
        persistLists();
        const still = (listCache.roles || []).find(function (r) {
            return normStr(r.name).toLowerCase() === nextName.toLowerCase();
        });
        setActiveView('role', still ? still.code : found.code);
        updateRolePeopleHint();
        updateRoleM365Ui();
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
        setActiveView('verwaltung');
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
        const el = document.getElementById('slgMemberPreview');
        if (!el) return;
        el.replaceChildren();
        const emails = isSammelView() ? emailsForSammelKind(activeView) : listCache.members || [];
        if (!emails.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent = isSammelView()
                ? 'Keine E-Mails für diese Zielgruppe in den Stammdaten.'
                : 'Keine E-Mails in der Verwaltungsliste.';
            el.appendChild(p);
            return;
        }
        emails.slice(0, 30).forEach(function (em) {
            const d = document.createElement('div');
            d.textContent = em;
            d.style.padding = '4px 0';
            d.style.borderBottom = '1px solid #eef1f4';
            el.appendChild(d);
        });
        if (emails.length > 30) {
            const more = document.createElement('div');
            more.className = 'muted';
            more.style.paddingTop = '8px';
            more.textContent = '… und ' + String(emails.length - 30) + ' weitere.';
            el.appendChild(more);
        }
    }

    function setupDraftForKind(kind) {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const su = window.ms365AppDataV2.getSetup() || {};
                if (kind === 'schulleitung') return su.schulleitungDraft || {};
                return su.verwaltungDraft || {};
            }
        } catch {
            /* ignore */
        }
        return {};
    }

    function fillCreateFormFromDraft() {
        if (!isSammelView()) return;
        const d = setupDraftForKind(activeView);
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const desc = document.getElementById('slgNewDescription');
        const ct = document.getElementById('slgNewCreateTeam');
        if (activeView === 'schulleitung') {
            if (dn) dn.value = String(d.slNewDisplayName != null ? d.slNewDisplayName : 'Schulleitung');
            if (nn) nn.value = String(d.slNewMailNick != null ? d.slNewMailNick : 'schulleitung');
            if (desc) {
                desc.value = String(
                    d.slNewDescription != null ? d.slNewDescription : 'Schulleitung (MS365-Schul-Tools)'
                );
            }
            if (ct) ct.checked = !!d.slNewCreateTeam;
        } else {
            if (dn) dn.value = String(d.vwNewDisplayName != null ? d.vwNewDisplayName : 'Schulverwaltung');
            if (nn) nn.value = String(d.vwNewMailNick != null ? d.vwNewMailNick : 'verwaltung');
            if (desc) {
                desc.value = String(
                    d.vwNewDescription != null ? d.vwNewDescription : 'Schulverwaltung Personal (MS365-Schul-Tools)'
                );
            }
            if (ct) ct.checked = !!d.vwNewCreateTeam;
        }
    }

    function applyCreateDefaults() {
        fillCreateFormFromDraft();
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const desc = document.getElementById('slgNewDescription');
        if (activeView === 'schulleitung') {
            if (dn && !dn.value) dn.value = 'Schulleitung';
            if (nn && !nn.value) nn.value = 'schulleitung';
            if (desc && !desc.value) desc.value = 'Schulleitung (MS365-Schul-Tools)';
        } else {
            if (dn && !dn.value) dn.value = 'Schulverwaltung';
            if (nn && !nn.value) nn.value = 'verwaltung';
            if (desc && !desc.value) desc.value = 'Schulverwaltung Personal (MS365-Schul-Tools)';
        }
    }

    function createDraftFromForm() {
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const desc = document.getElementById('slgNewDescription');
        const ct = document.getElementById('slgNewCreateTeam');
        if (activeView === 'schulleitung') {
            return {
                slNewDisplayName: dn ? dn.value : '',
                slNewMailNick: nn ? nn.value : '',
                slNewDescription: desc ? desc.value : '',
                slNewCreateTeam: ct ? !!ct.checked : false
            };
        }
        return {
            vwNewDisplayName: dn ? dn.value : '',
            vwNewMailNick: nn ? nn.value : '',
            vwNewDescription: desc ? desc.value : '',
            vwNewCreateTeam: ct ? !!ct.checked : false
        };
    }

    function getActiveMatchedId() {
        if (!isSammelView()) return null;
        const meta = activeSammelMeta();
        const gid = meta.audienceGroupId;
        if (gid === 'schulleitung') return matchedGroupIds.schulleitung;
        if (gid === 'verwaltung') return matchedGroupIds.verwaltung;
        return (matchedGroupIds.custom && matchedGroupIds.custom[gid]) || null;
    }

    function setActiveMatchedId(id) {
        if (!isSammelView()) return;
        const meta = activeSammelMeta();
        const gid = meta.audienceGroupId;
        const val = id ? String(id) : null;
        if (gid === 'schulleitung') matchedGroupIds.schulleitung = val;
        else if (gid === 'verwaltung') matchedGroupIds.verwaltung = val;
        else {
            if (!matchedGroupIds.custom) matchedGroupIds.custom = {};
            matchedGroupIds.custom[gid] = val;
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                const su = window.ms365AppDataV2.getSetup() || {};
                const map = Object.assign({}, su.matched && su.matched.verwaltungAudienceGroupIds);
                if (val) map[gid] = val;
                else delete map[gid];
                window.ms365AppDataV2.patchSetup({
                    matched: { verwaltungAudienceGroupIds: map }
                });
            }
        }
        live().resetCaches();
    }

    async function runSyncMembers() {
        const gid = getActiveMatchedId();
        if (!gid) {
            toast('Zuerst eine Gruppe matchen oder anlegen.');
            return;
        }
        const meta = activeSammelMeta();
        const emails = emailsForSammelKind(meta.kind);
        if (!emails.length) {
            toast('Keine E-Mails für „' + meta.label + '“ in den Stammdaten.');
            return;
        }
        clearSyncLog();
        appendSyncLog('Start: ' + meta.label + ' (' + emails.length + ' Adressen) …', '');
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
                    'Abgleich mit Stammdaten-Liste: +' + joinEmails.length + ' / −' + leaveEmails.length + '.',
                    ''
                );
            }
            if (joinEmails.length) {
                const r = await gug().syncEmailsToGroup(token, gid, joinEmails, meta.label, appendSyncLog);
                appendSyncLog('Aufnehmen: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            }
            if (leaveEmails.length && typeof gug().removeEmailsFromGroup === 'function') {
                const r = await gug().removeEmailsFromGroup(token, gid, leaveEmails, meta.label, appendSyncLog);
                appendSyncLog('Entfernen: ' + r.ok + ' OK, übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            }
            if (!joinEmails.length && !leaveEmails.length) {
                appendSyncLog('Keine Änderungen gegenüber der Stammdaten-Liste.', 'ok');
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
        let slDraft = setupDraftForKind('schulleitung');
        let vwDraft = setupDraftForKind('verwaltung');
        if (activeView === 'schulleitung') {
            slDraft = Object.assign({}, slDraft, createDraftFromForm());
        } else if (activeView === 'verwaltung') {
            vwDraft = Object.assign({}, vwDraft, createDraftFromForm());
        }
        return {
            kind: STORAGE_KEY,
            savedAt: new Date().toISOString(),
            matched: {
                schulleitungGroupId: matchedGroupIds.schulleitung,
                verwaltungGroupId: matchedGroupIds.verwaltung
            },
            schulleitungDraft: slDraft,
            verwaltungDraft: vwDraft
        };
    }

    function applyStateObject(o) {
        if (!o || typeof o !== 'object') return;
        matchedGroupIds = { schulleitung: null, verwaltung: null, custom: {} };
        if (o.matched && typeof o.matched === 'object') {
            if (o.matched.schulleitungGroupId) {
                matchedGroupIds.schulleitung = String(o.matched.schulleitungGroupId);
            }
            if (o.matched.verwaltungGroupId) {
                matchedGroupIds.verwaltung = String(o.matched.verwaltungGroupId);
            }
            const audMap = o.matched.verwaltungAudienceGroupIds;
            if (audMap && typeof audMap === 'object') {
                Object.keys(audMap).forEach(function (k) {
                    const id = audMap[k] ? String(audMap[k]).trim() : '';
                    if (id) matchedGroupIds.custom[k] = id;
                });
            }
        } else if (o.verwaltungGroupId) {
            matchedGroupIds.verwaltung = String(o.verwaltungGroupId);
        }
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            const patch = {
                matched: {
                    schulleitungGroupId: matchedGroupIds.schulleitung,
                    verwaltungGroupId: matchedGroupIds.verwaltung,
                    verwaltungAudienceGroupIds: Object.assign({}, matchedGroupIds.custom)
                }
            };
            if (o.schulleitungDraft) patch.schulleitungDraft = o.schulleitungDraft;
            if (o.verwaltungDraft) patch.verwaltungDraft = o.verwaltungDraft;
            else if (o.vwNewDisplayName !== undefined) {
                patch.verwaltungDraft = {
                    vwNewDisplayName: o.vwNewDisplayName,
                    vwNewMailNick: o.vwNewMailNick,
                    vwNewDescription: o.vwNewDescription,
                    vwNewCreateTeam: o.vwNewCreateTeam
                };
            }
            window.ms365AppDataV2.patchSetup(patch);
        }
        live().resetCaches();
        gd().clearSearchResults();
        fillCreateFormFromDraft();
        const gid = getActiveMatchedId();
        live().setMatchedMode(!!gid);
        live().fillForm(gid ? { id: gid } : null);
        updateLeftListUi();
        gd().setTab('general');
    }

    function saveState() {
        try {
            const obj = buildStateObject();
            localStorage.setItem(STORAGE_KEY, JSON.stringify(obj));
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: {
                        schulleitungGroupId: obj.matched.schulleitungGroupId,
                        verwaltungGroupId: obj.matched.verwaltungGroupId
                    },
                    schulleitungDraft: obj.schulleitungDraft,
                    verwaltungDraft: obj.verwaltungDraft
                });
            }
        } catch {
            // ignore
        }
    }

    function loadState() {
        let rawLocal = null;
        let localObj = null;
        try {
            rawLocal = localStorage.getItem(STORAGE_KEY);
            if (rawLocal) localObj = JSON.parse(rawLocal);
        } catch {
            rawLocal = null;
            localObj = null;
        }
        const localMatched =
            localObj && localObj.matched && typeof localObj.matched === 'object' ? localObj.matched : null;
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const su = window.ms365AppDataV2.getSetup() || {};
                const merged = su.matched && typeof su.matched === 'object' ? Object.assign({}, su.matched) : {};
                if (localMatched) {
                    if (localMatched.verwaltungGroupId && !merged.verwaltungGroupId) {
                        merged.verwaltungGroupId = String(localMatched.verwaltungGroupId).trim();
                    }
                    if (localMatched.schulleitungGroupId && !merged.schulleitungGroupId) {
                        merged.schulleitungGroupId = String(localMatched.schulleitungGroupId).trim();
                    }
                }
                applyStateObject({
                    matched: merged,
                    schulleitungDraft:
                        su.schulleitungDraft ||
                        (localObj && localObj.schulleitungDraft) ||
                        undefined,
                    verwaltungDraft:
                        su.verwaltungDraft || (localObj && localObj.verwaltungDraft) || undefined
                });
                return;
            }
        } catch {
            // ignore
        }
        try {
            if (!localObj) return;
            applyStateObject(localObj);
        } catch {
            // ignore
        }
    }

    function clearStorage() {
        try {
            localStorage.removeItem(STORAGE_KEY);
            matchedGroupIds = { schulleitung: null, verwaltung: null };
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: { schulleitungGroupId: null, verwaltungGroupId: null }
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
                    refreshSplitMigration();
                },
                persistUnmatch: function () {
                    setActiveMatchedId(null);
                    saveState();
                    refreshSplitMigration();
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
        document.querySelectorAll('#slgListItems [data-vw-kind="schulleitung"], #slgListItems [data-vw-kind="verwaltung"]').forEach(
            function (btn) {
                btn.addEventListener('click', function () {
                    setActiveView(btn.getAttribute('data-vw-kind'));
                });
            }
        );
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
        onClick('vwBtnSplitMigration', function () {
            patchMigration({ skippedAt: null, completedAt: null });
            if (splitMigrationUi && typeof splitMigrationUi.resetPreview === 'function') {
                splitMigrationUi.resetPreview();
            }
            refreshSplitMigration();
            toast('Assistent „Aufteilen“ – Hinweis oben.');
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
        const kindParam = new URLSearchParams(window.location.search).get('kind');
        let initialView = 'verwaltung';
        if (kindParam && String(kindParam).indexOf('audience:') === 0) initialView = String(kindParam);
        else if (kindParam === 'schulleitung' || kindParam === 'verwaltung') initialView = kindParam;
        setActiveView(initialView);
        if (!getActiveMatchedId()) {
            live().setMatchedMode(false);
            applyCreateDefaults();
        } else {
            void live().loadGroup({ silent: true });
        }
        void refreshGraphMemberCounts();
        (function trySplitUi(attempts) {
            if (window.ms365VerwaltungSplitMigrationUi) {
                initSplitMigration();
                refreshSplitMigration();
                return;
            }
            if (attempts > 40) return;
            setTimeout(function () {
                trySplitUi(attempts + 1);
            }, 50);
        })(0);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMCont