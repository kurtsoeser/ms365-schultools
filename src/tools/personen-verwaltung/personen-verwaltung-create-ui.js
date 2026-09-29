/**
 * Create/Delete-Modals (Analyse 02 Phase B).
 */
import { pv } from './personen-verwaltung-state.js';
import { GRAPH_SCOPES, getGraphToken, graphJson, graphDelete, fetchAllPages, odataEscape } from './personen-verwaltung-graph.js';
import {
    graphErrorFriendly,
    norm,
    sanitizeMailNickname,
    isGuid,
    isDuplicateMemberError,
    assignedSkuIdsOfUser,
    skuLookupFromSubscribed,
    Lic,
    userLicenseSummary,
    groupTypeLabel,
    formatDate
} from './personen-verwaltung-logic.js';
import { dlgConfirm } from '../../shared/utils/dialog.js';
import { readInputTrim } from './personen-verwaltung-logic.js';

/** @type {any} */
let host = null;

export function init(a) {
    host = a;
}

export function openCreateModal() {
    const bd = document.getElementById('pvModalCreateBackdrop');
    if (!bd) return;
    const ids = ['pvCreateUpn', 'pvCreateDisplayName', 'pvCreateMailNick', 'pvCreatePassword', 'pvCreateGiven', 'pvCreateSurname', 'pvCreateMail'];
    for (let i = 0; i < ids.length; i++) {
        const el = document.getElementById(ids[i]);
        if (el) el.value = '';
    }
    const f = document.getElementById('pvCreateForcePw');
    if (f) f.checked = true;
    const e = document.getElementById('pvCreateEnabled');
    if (e) e.checked = true;
    bd.classList.add('active');
    bd.setAttribute('aria-hidden', 'false');
}

export function closeCreateModal() {
    const bd = document.getElementById('pvModalCreateBackdrop');
    if (!bd) return;
    bd.classList.remove('active');
    bd.setAttribute('aria-hidden', 'true');
}

export async function submitCreateUser() {
    const upn = readInputTrim(document.getElementById('pvCreateUpn'));
    const displayName = readInputTrim(document.getElementById('pvCreateDisplayName'));
    let mailNick = readInputTrim(document.getElementById('pvCreateMailNick'));
    const password = String(document.getElementById('pvCreatePassword')?.value || '');
    const givenName = readInputTrim(document.getElementById('pvCreateGiven'));
    const surname = readInputTrim(document.getElementById('pvCreateSurname'));
    const mail = readInputTrim(document.getElementById('pvCreateMail'));
    const forcePw = document.getElementById('pvCreateForcePw') ? document.getElementById('pvCreateForcePw').checked : true;
    const enabled = document.getElementById('pvCreateEnabled') ? document.getElementById('pvCreateEnabled').checked : true;

    if (!upn || !displayName || !password) {
        host.toast('UPN, Anzeigename und Kennwort sind Pflichtfelder.');
        return;
    }
    mailNick = sanitizeMailNickname(mailNick, upn);
    if (!mailNick) {
        host.toast('Mail-Nickname ungültig oder leer.');
        return;
    }

    const body = {
        accountEnabled: enabled,
        displayName: displayName,
        mailNickname: mailNick,
        userPrincipalName: upn,
        passwordProfile: {
            password: password,
            forceChangePasswordNextSignIn: !!forcePw
        }
    };
    if (givenName) body.givenName = givenName;
    if (surname) body.surname = surname;
    if (mail) body.mail = mail;

    const sub = document.getElementById('pvModalCreateSubmit');
    if (sub) sub.disabled = true;
    try {
        const token = await getGraphToken();
        const created = await graphJson('POST', '/users', token, body);
        const id = created && created.id ? created.id : null;
        host.appendLog('Benutzer angelegt: ' + (created.userPrincipalName || upn), 'ok');
        host.toast('Benutzer angelegt. Als Nächstes: Nutzungsort und Lizenz zuweisen.');
        closeCreateModal();
        if (id) {
            pv.pendingTabAfterSelect = 'lizenzen';
            try {
                const fresh = await host.refreshUserFromGraph(token, id);
                host.mergeUserIntoList(fresh);
                host.selectUser(id);
            } catch {
                pv.loadedUsers.push(created);
                host.refreshDepartmentFilter();
                host.updateStatsPanel();
                if (id) host.selectUser(id);
            }
        }
        host.renderUserTree();
    } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        host.appendLog('Anlegen: ' + msg, 'err');
        host.toast(msg);
    } finally {
        if (sub) sub.disabled = false;
    }
}

export function updateDeleteModalUi() {
    const hard = document.getElementById('pvDeleteHard') && document.getElementById('pvDeleteHard').checked;
    const warn = document.getElementById('pvDeleteHardWarn');
    const intro = document.getElementById('pvDeleteSoftIntro');
    const sub = document.getElementById('pvModalDeleteSubmit');
    if (warn) warn.style.display = hard ? 'block' : 'none';
    if (intro) intro.style.opacity = hard ? '0.55' : '1';
    if (sub) {
        sub.textContent = hard ? 'Endgültig löschen' : 'Konto deaktivieren';
        sub.className = hard ? 'btn btn-danger' : 'btn btn-success';
    }
}

export function openDeleteModal() {
    const u = host.getSelectedUser();
    if (!u) return;
    if (u.onPremisesSyncEnabled === true) {
        const hardChk = document.getElementById('pvDeleteHard');
        // Hard-Delete aus der Cloud geht bei AD-Sync nicht – Modal nur für Deaktivieren erlauben
        if (hardChk) {
            hardChk.checked = false;
            hardChk.disabled = true;
        }
        host.toast(
            'AD‑Sync‑Konto: Endgültig löschen nur im lokalen AD. Deaktivieren in der Cloud ggf. möglich – SAM: ' +
                (u.onPremisesSamAccountName || '–')
        );
    } else {
        const hardChk = document.getElementById('pvDeleteHard');
        if (hardChk) hardChk.disabled = false;
    }
    const bd = document.getElementById('pvModalDeleteBackdrop');
    const echo = document.getElementById('pvDeleteUpnEcho');
    const inp = document.getElementById('pvDeleteConfirmInput');
    const sub = document.getElementById('pvModalDeleteSubmit');
    const hardChk2 = document.getElementById('pvDeleteHard');
    if (hardChk2 && !u.onPremisesSyncEnabled) hardChk2.checked = false;
    updateDeleteModalUi();
    if (echo) echo.textContent = u.userPrincipalName || u.mail || u.id;
    if (inp) inp.value = '';
    if (sub) sub.disabled = true;
    if (bd) {
        bd.classList.add('active');
        bd.setAttribute('aria-hidden', 'false');
    }
}

export function closeDeleteModal() {
    const bd = document.getElementById('pvModalDeleteBackdrop');
    if (!bd) return;
    bd.classList.remove('active');
    bd.setAttribute('aria-hidden', 'true');
}

export function syncDeleteConfirmButton() {
    const u = host.getSelectedUser();
    const inp = document.getElementById('pvDeleteConfirmInput');
    const sub = document.getElementById('pvModalDeleteSubmit');
    if (!sub || !inp || !u) return;
    const ok = readInputTrim(inp) === String(u.userPrincipalName || '').trim();
    sub.disabled = !ok;
}

export async function submitDeleteUser() {
    const u = host.getSelectedUser();
    if (!u) return;
    const inp = document.getElementById('pvDeleteConfirmInput');
    if (!inp || readInputTrim(inp) !== String(u.userPrincipalName || '').trim()) {
        host.toast('UPN stimmt nicht überein.');
        return;
    }
    const hard = document.getElementById('pvDeleteHard') && document.getElementById('pvDeleteHard').checked;
    if (hard && u.onPremisesSyncEnabled === true) {
        host.toast('AD‑Sync‑Konten können nicht aus der Cloud gelöscht werden – bitte lokal im AD löschen.');
        return;
    }
    const sub = document.getElementById('pvModalDeleteSubmit');
    if (sub) sub.disabled = true;
    try {
        const token = await getGraphToken();
        if (!hard) {
            if (u.accountEnabled === false) {
                host.toast('Konto ist bereits deaktiviert.');
                closeDeleteModal();
                return;
            }
            await graphJson('PATCH', '/users/' + encodeURIComponent(u.id), token, { accountEnabled: false });
            const fresh = await host.refreshUserFromGraph(token, u.id);
            host.mergeUserIntoList(fresh);
            host.appendLog('Konto deaktiviert: ' + (fresh.userPrincipalName || u.id), 'ok');
            host.toast('Konto deaktiviert.');
            closeDeleteModal();
            pv.profileEditMode = false;
            pv.cachedGroupsForSelection = null;
            host.selectUser(fresh.id);
            host.updateStatsPanel();
            return;
        }

        await graphDelete('/users/' + encodeURIComponent(u.id), token);
        pv.loadedUsers = pv.loadedUsers.filter(function (x) {
            return x.id !== u.id;
        });
        host.appendLog('Benutzer gelöscht (DELETE): ' + (u.userPrincipalName || u.id), 'ok');
        host.toast('Dauerhaft gelöscht.');
        closeDeleteModal();
        pv.selectedUserId = null;
        pv.cachedGroupsForSelection = null;
        pv.profileEditMode = false;
        const hint = document.getElementById('pvHint');
        const detail = document.getElementById('pvDetail');
        if (hint) hint.style.display = '';
        if (detail) detail.style.display = 'none';
        host.updateDetailActionButtons();
        host.refreshDepartmentFilter();
        host.updateStatsPanel();
        host.renderUserTree();
    } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        host.appendLog((hard ? 'Löschen: ' : 'Deaktivieren: ') + msg, 'err');
        host.toast(msg);
    } finally {
        if (sub) sub.disabled = false;
        syncDeleteConfirmButton();
    }
}

