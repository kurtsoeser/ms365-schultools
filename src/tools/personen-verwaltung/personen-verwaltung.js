/**
 * Personen-Verwaltung – Entry (Analyse 02 Phase B).
 */
import {
    GRAPH_SCOPES,
    getGraphToken,
    graphRequest,
    graphJson,
    graphDelete,
    fetchAllPages,
    sleep,
    odataEscape,
    appendLog,
    clearLog
} from './personen-verwaltung-graph.js';
import {
    graphErrorFriendly,
    norm,
    compareStrings,
    readSortFromSelect,
    formatPhones,
    formatDate,
    groupTypeLabel,
    userTypeLabel,
    sanitizeMailNickname,
    isGuid,
    isDuplicateMemberError,
    assignedSkuIdsOfUser,
    skuLookupFromSubscribed,
    Lic,
    userLicenseSummary
} from './personen-verwaltung-logic.js';
import { dlgConfirm, dlgPrompt } from '../../shared/utils/dialog.js';

import {
    init as initCreateUi,
    openCreateModal,
    closeCreateModal,
    submitCreateUser,
    updateDeleteModalUi,
    openDeleteModal,
    closeDeleteModal,
    syncDeleteConfirmButton,
    submitDeleteUser
} from './personen-verwaltung-create-ui.js';
import {
    init as initLicensePanel,
    renderLicenseTab,
    loadSubscribedSkus,
    ensureUsageLocation,
    assignSelectedLicense,
    removeLicense,
    saveUsageLocation
} from './personen-verwaltung-license-panel.js';
import {
    init as initSearchUi,
    getVisibleRows,
    refreshDepartmentFilter,
    refreshLicenseFilter,
    updateStatsPanel,
    updateProgressLine,
    renderUserTree
} from './personen-verwaltung-search-ui.js';
import {
    init as initGroupsUi,
    loadGroupsForSelected,
    searchGroupsForAdd,
    addSelectedUserToGroup,
    removeUserFromGroup
} from './personen-verwaltung-groups-ui.js';
import {
    init as initProfileUi,
    renderProfileTab,
    refreshUserFromGraph,
    mergeUserIntoList,
    saveProfilePatch,
    resetProfileFromGraph
} from './personen-verwaltung-profile-ui.js';

import { pv } from './personen-verwaltung-state.js';
const USER_LIST_SELECT = pv.USER_LIST_SELECT;
const USER_REFRESH_SELECT = pv.USER_REFRESH_SELECT;
const GROUP_MEMBEROF_SELECT = pv.GROUP_MEMBEROF_SELECT;
const AD_FLAGS_KEY = pv.AD_FLAGS_KEY;
const SESSION_CACHE_KEY = pv.SESSION_CACHE_KEY;
const SESSION_CACHE_MAX_AGE_MS = pv.SESSION_CACHE_MAX_AGE_MS;

function saveUsersToSession() {
    try {
        sessionStorage.setItem(
            SESSION_CACHE_KEY,
            JSON.stringify({
                savedAt: Date.now(),
                users: pv.loadedUsers,
                skus: pv.subscribedSkus,
                skusOk: pv.subscribedSkusOk
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
        pv.loadedUsers = obj.users;
        pv.loadedUsers = applyAdFlagsToUsers(pv.loadedUsers);
        pv.subscribedSkus = Array.isArray(obj.skus) ? obj.skus : [];
        pv.subscribedSkusOk = !!obj.skusOk;
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

function loadAdUserFlags() {
    try {
        const raw = localStorage.getItem(AD_FLAGS_KEY);
        if (!raw) return {};
        const obj = JSON.parse(raw);
        if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return {};
        return obj;
    } catch {
        return {};
    }
}

function saveAdUserFlags(map) {
    try {
        localStorage.setItem(AD_FLAGS_KEY, JSON.stringify(map && typeof map === 'object' ? map : {}));
    } catch {
        /* ignore */
    }
}

function patchAdUserFlag(userId, patch) {
    const id = String(userId || '').trim();
    if (!id) return;
    const map = loadAdUserFlags();
    const prev = map[id] || { flagged: false, note: '', flaggedAt: '' };
    const flagged = patch && patch.flagged !== undefined ? !!patch.flagged : !!prev.flagged;
    const note = patch && patch.note !== undefined ? String(patch.note) : String(prev.note || '');
    if (!flagged && !String(note || '').trim()) {
        delete map[id];
    } else {
        map[id] = {
            flagged: flagged,
            note: note,
            flaggedAt:
                flagged && !prev.flagged
                    ? new Date().toISOString()
                    : prev.flaggedAt || (flagged ? new Date().toISOString() : '')
        };
    }
    saveAdUserFlags(map);
}

function applyAdFlagsToUsers(users) {
    const flags = loadAdUserFlags();
    return (Array.isArray(users) ? users : []).map(function (u) {
        if (!u || !u.id) return u;
        const f = flags[String(u.id)] || null;
        const next = Object.assign({}, u);
        next.onPremisesSyncEnabled = u.onPremisesSyncEnabled === true;
        next.adFlagged = !!(f && f.flagged);
        next.adFlagNote = f && f.note ? String(f.note) : '';
        next.adFlaggedAt = f && f.flaggedAt ? String(f.flaggedAt) : '';
        return next;
    });
}

function formatAdSyncDate(iso) {
    const s = String(iso || '').trim();
    if (!s) return '–';
    try {
        const d = new Date(s);
        if (isNaN(d.getTime())) return s;
        return d.toLocaleString('de-AT');
    } catch {
        return s;
    }
}

function fillUserAdPanels(u) {
    const adPanel = document.getElementById('pvAdSyncPanel');
    const cloudPanel = document.getElementById('pvCloudFlagPanel');
    if (!u) {
        if (adPanel) adPanel.style.display = 'none';
        if (cloudPanel) cloudPanel.style.display = 'none';
        return;
    }
    const isAd = u.onPremisesSyncEnabled === true;
    if (adPanel) adPanel.style.display = isAd ? '' : 'none';
    if (cloudPanel) cloudPanel.style.display = isAd ? 'none' : '';
    if (isAd) {
        const sam = document.getElementById('pvAdSam');
        const dom = document.getElementById('pvAdDomain');
        const last = document.getElementById('pvAdLastSync');
        const note = document.getElementById('pvAdFlagNote');
        const flagged = document.getElementById('pvAdFlagged');
        if (sam) sam.value = String(u.onPremisesSamAccountName || '–');
        if (dom) dom.value = String(u.onPremisesDomainName || '–');
        if (last) last.value = formatAdSyncDate(u.onPremisesLastSyncDateTime);
        if (note) note.value = String(u.adFlagNote || '');
        if (flagged) flagged.checked = !!u.adFlagged;
    } else {
        const note = document.getElementById('pvCloudFlagNote');
        const flagged = document.getElementById('pvCloudFlagged');
        if (note) note.value = String(u.adFlagNote || '');
        if (flagged) flagged.checked = !!u.adFlagged;
    }
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

/* Graph-Client → personen-verwaltung-graph.js */

/* Logic → personen-verwaltung-logic.js */

function getSelectedUser() {
    if (!pv.selectedUserId) return null;
    return pv.loadedUsers.find(function (x) {
        return x.id === pv.selectedUserId;
    }) || null;
}

function updateDetailActionButtons() {
    const save = document.getElementById('pvBtnSave');
    const saveBottom = document.getElementById('pvBtnSaveBottom');
    const cancel = document.getElementById('pvBtnCancelEdit');
    const del = document.getElementById('pvBtnDelete');
    const hasSel = !!pv.selectedUserId;
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


/* getVisibleRows… ausgelagert */
/* Profil-UI → personen-verwaltung-profile-ui.js */


/* openCreateModal… ausgelagert */
/* renderGroupsTable… → personen-verwaltung-groups-ui.js */
function setTab(tab) {
    if (tab === 'gruppen') pv.activeTab = 'gruppen';
    else if (tab === 'lizenzen') pv.activeTab = 'lizenzen';
    else pv.activeTab = 'profil';

    const rows = [
        ['profil', 'pvPanelProfil', 'pvTabProfil'],
        ['lizenzen', 'pvPanelLizenzen', 'pvTabLizenzen'],
        ['gruppen', 'pvPanelGruppen', 'pvTabGruppen']
    ];
    rows.forEach(function (row) {
        const on = pv.activeTab === row[0];
        const p = document.getElementById(row[1]);
        const b = document.getElementById(row[2]);
        if (p) {
            p.classList.toggle('active', on);
            p.setAttribute('aria-hidden', on ? 'false' : 'true');
        }
        if (b) b.setAttribute('aria-selected', on ? 'true' : 'false');
    });

    if (pv.activeTab === 'gruppen' && pv.selectedUserId) {
        if (pv.cachedGroupsForSelection === null) {
            loadGroupsForSelected();
        }
    }
    if (pv.activeTab === 'lizenzen' && pv.selectedUserId) {
        renderLicenseTab();
    }
}

function selectUser(userId) {
    pv.selectedUserId = userId || null;
    pv.cachedGroupsForSelection = null;
    pv.activeTab = 'profil';
    pv.profileEditMode = !!pv.selectedUserId;
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

    if (!pv.selectedUserId) {
        if (hint) hint.style.display = '';
        if (detail) detail.style.display = 'none';
        fillUserAdPanels(null);
        updateDetailActionButtons();
        renderUserTree();
        return;
    }

    const u = getSelectedUser();

    if (hint) hint.style.display = 'none';
    if (detail) detail.style.display = '';
    if (title) title.textContent = u && u.displayName ? String(u.displayName) : '(ohne Anzeigename)';

    fillUserAdPanels(u || null);
    renderProfileTab(u || null, true);
    updateDetailActionButtons();
    setTab(pv.pendingTabAfterSelect || 'profil');
    pv.pendingTabAfterSelect = '';
    renderUserTree();
}

async function loadUsers() {
    const btn = document.getElementById('pvBtnLoad');
    const btnCsv = document.getElementById('pvBtnCsv');
    const progress = document.getElementById('pvProgress');
    if (btn) btn.disabled = true;
    if (btnCsv) btnCsv.disabled = true;
    clearLog();
    pv.loadedUsers = [];
    pv.selectedUserId = null;
    pv.cachedGroupsForSelection = null;
    pv.profileEditMode = false;
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
        pv.loadedUsers = applyAdFlagsToUsers(users);
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
        if (btnCsv) btnCsv.disabled = !pv.loadedUsers.length;
        renderUserTree();
    }
}

function exportCsv() {
    if (!pv.loadedUsers.length) {
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
        'license',
        'onPremisesSyncEnabled',
        'onPremisesSamAccountName',
        'onPremisesDomainName',
        'onPremisesLastSyncDateTime',
        'onPremisesSecurityIdentifier',
        'adFlagged',
        'adFlagNote'
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
            } else if (key === 'onPremisesSyncEnabled' || key === 'adFlagged') {
                v = u[key] === true ? 'true' : 'false';
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

function exportAdUserReport(kind) {
    if (!pv.loadedUsers.length) {
        toast('Keine Daten – zuerst Personen einlesen.');
        return;
    }
    const rows =
        kind === 'flagged'
            ? pv.loadedUsers.filter(function (u) {
                  return u && u.adFlagged;
              })
            : pv.loadedUsers.filter(function (u) {
                  return u && u.onPremisesSyncEnabled === true;
              });
    if (!rows.length) {
        toast(kind === 'flagged' ? 'Keine markierten Konten.' : 'Keine AD‑Sync‑Konten gefunden.');
        return;
    }
    const headers = [
        'displayName',
        'userPrincipalName',
        'mail',
        'department',
        'jobTitle',
        'accountEnabled',
        'userType',
        'onPremisesSamAccountName',
        'onPremisesDomainName',
        'onPremisesLastSyncDateTime',
        'onPremisesSecurityIdentifier',
        'adFlagged',
        'adFlagNote',
        'id'
    ];
    const list = rows;
    const lines = [headers.join(';')];
    for (let i = 0; i < list.length; i++) {
        const u = list[i];
        lines.push(
            headers
                .map(function (key) {
                    let v = u[key];
                    if (key === 'adFlagged' || key === 'accountEnabled') {
                        v = u[key] === true ? 'true' : u[key] === false ? 'false' : '';
                    }
                    if (v === undefined || v === null) v = '';
                    return '"' + String(v).replace(/"/g, '""') + '"';
                })
                .join(';')
        );
    }
    const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = kind === 'flagged' ? 'ad-admin-markierte-personen.csv' : 'ad-sync-personen.csv';
    a.click();
    URL.revokeObjectURL(a.href);
    appendLog(
        (kind === 'flagged' ? 'Markierte' : 'AD‑Sync') + ' Bericht: ' + list.length + ' Zeile(n).',
        'ok'
    );
    toast(list.length + ' Zeile(n) exportiert.');
}

function saveFlagForSelectedUser(flaggedId, noteId) {
    const u = getSelectedUser();
    if (!u) return;
    const flaggedEl = document.getElementById(flaggedId);
    const noteEl = document.getElementById(noteId);
    const flagged = !!(flaggedEl && flaggedEl.checked);
    const note = noteEl ? String(noteEl.value || '') : '';
    patchAdUserFlag(u.id, { flagged: flagged, note: note });
    pv.loadedUsers = applyAdFlagsToUsers(pv.loadedUsers);
    saveUsersToSession();
    updateStatsPanel();
    selectUser(u.id);
    toast(flagged ? 'Für lokalen Admin markiert.' : 'Markierung entfernt.');
}

function bind() {
        const api = {
            toast,
            appendLog,
            clearLog,
            getSelectedUser,
            selectUser,
            mergeUserIntoList,
            updateDetailActionButtons,
            updateStatsPanel,
            updateProgressLine,
            renderUserTree,
            refreshDepartmentFilter,
            refreshLicenseFilter,
            setTab,
            renderProfileTab,
            loadGroupsForSelected,
            renderLicenseTab,
            saveUsersToSession,
            getVisibleRows,
            refreshUserFromGraph,
            fillUserAdPanels
        };
        initProfileUi(api);
        initCreateUi(api);
        initLicensePanel(api);
        initSearchUi(api);
        initGroupsUi(api);



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
    const fad = document.getElementById('pvFilterAdSync');
    const fs = document.getElementById('pvSortKey');
    if (ft) ft.addEventListener('change', reRender);
    if (fa) fa.addEventListener('change', reRender);
    if (fd) fd.addEventListener('change', reRender);
    if (fl) fl.addEventListener('change', reRender);
    if (fad) fad.addEventListener('change', reRender);
    if (fs) fs.addEventListener('change', reRender);

    document.getElementById('pvAdSummary')?.addEventListener('click', function (ev) {
        const t = ev.target && ev.target.closest ? ev.target.closest('[data-pv-ad]') : null;
        if (!t) return;
        const src = String(t.getAttribute('data-pv-ad') || '');
        const sel = document.getElementById('pvFilterAdSync');
        if (!sel || !src) return;
        sel.value = src;
        reRender();
    });
    document.getElementById('pvBtnExportAdSync')?.addEventListener('click', function () {
        exportAdUserReport('adSync');
    });
    document.getElementById('pvBtnExportAdFlagged')?.addEventListener('click', function () {
        exportAdUserReport('flagged');
    });
    document.getElementById('pvAdFlagSave')?.addEventListener('click', function () {
        saveFlagForSelectedUser('pvAdFlagged', 'pvAdFlagNote');
    });
    document.getElementById('pvCloudFlagSave')?.addEventListener('click', function () {
        saveFlagForSelectedUser('pvCloudFlagged', 'pvCloudFlagNote');
    });

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
        pv.cachedGroupsForSelection = null;
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
        appendLog('Benutzerliste aus Sitzungs-Cache wiederhergestellt (' + pv.loadedUsers.length + ' Person(en)).', 'ok');
    }

    updateDetailActionButtons();
    renderUserTree();

    try {
        const q = new URLSearchParams(window.location.search);
        const tab = String(q.get('tab') || '').toLowerCase();
        if (tab === 'lizenzen' || tab === 'gruppen' || tab === 'profil') pv.pendingTabAfterSelect = tab;
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
