/**
 * Create/Delete-Modals (Analyse 02 Phase B).
 */
import { pv } from './personen-verwaltung-state.js';
import { getGraphToken, graphJson, graphDelete } from './personen-verwaltung-graph.js';
import { sanitizeMailNickname } from './personen-verwaltung-logic.js';

import { readInputTrim } from './personen-verwaltung-logic.js';

/** @type {any} */
let host = null;

/** Nutzer hat Anzeigename / Nickname manuell geändert → kein Auto-Overwrite. */
let displayNameTouched = false;
let mailNickTouched = false;
let createWired = false;

export function init(a) {
    host = a;
}

function el(id) {
    return document.getElementById(id);
}

function suggestDisplayName() {
    const given = readInputTrim(el('pvCreateGiven'));
    const surname = readInputTrim(el('pvCreateSurname'));
    return [given, surname].filter(Boolean).join(' ').trim();
}

function suggestMailNickFromUpn(upn) {
    return sanitizeMailNickname('', upn);
}

/** Kennwort das übliche Richtlinien (Länge, Groß/Klein, Ziffer, Sonderzeichen) erfüllt. */
export function suggestCreatePassword() {
    const upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
    const lower = 'abcdefghijkmnopqrstuvwxyz';
    const digits = '23456789';
    const special = '!@#$%&*-_+';
    const all = upper + lower + digits + special;
    const pick = function (set) {
        return set.charAt(Math.floor(Math.random() * set.length));
    };
    let out = pick(upper) + pick(lower) + pick(digits) + pick(special);
    for (let i = 0; i < 10; i++) out += pick(all);
    return out
        .split('')
        .sort(function () {
            return Math.random() - 0.5;
        })
        .join('');
}

function syncDerivedFields() {
    if (!displayNameTouched) {
        const dn = el('pvCreateDisplayName');
        const suggested = suggestDisplayName();
        if (dn && suggested) dn.value = suggested;
    }
    if (!mailNickTouched) {
        const nick = el('pvCreateMailNick');
        const upn = readInputTrim(el('pvCreateUpn'));
        const suggested = suggestMailNickFromUpn(upn);
        if (nick && suggested) nick.value = suggested;
    }
}

function wireCreateAutofill() {
    if (createWired) return;
    createWired = true;

    const given = el('pvCreateGiven');
    const surname = el('pvCreateSurname');
    const display = el('pvCreateDisplayName');
    const upn = el('pvCreateUpn');
    const nick = el('pvCreateMailNick');
    const pwGen = el('pvCreatePwGen');
    const form = el('pvCreateForm');

    function onNameInput() {
        syncDerivedFields();
    }
    if (given) given.addEventListener('input', onNameInput);
    if (surname) surname.addEventListener('input', onNameInput);
    if (display) {
        display.addEventListener('input', function () {
            displayNameTouched = true;
        });
    }
    if (upn) {
        upn.addEventListener('input', function () {
            syncDerivedFields();
        });
    }
    if (nick) {
        nick.addEventListener('input', function () {
            mailNickTouched = true;
        });
    }
    if (pwGen) {
        pwGen.addEventListener('click', function () {
            const pw = el('pvCreatePassword');
            if (!pw) return;
            pw.value = suggestCreatePassword();
            pw.focus();
            pw.select();
        });
    }
    const usage = el('pvCreateUsageLocation');
    if (usage) {
        usage.addEventListener('input', function () {
            const start = usage.selectionStart;
            const end = usage.selectionEnd;
            usage.value = String(usage.value || '')
                .toUpperCase()
                .replace(/[^A-Z]/g, '')
                .slice(0, 2);
            if (typeof start === 'number') usage.setSelectionRange(start, end);
        });
    }
    if (form) {
        form.addEventListener('submit', function (ev) {
            ev.preventDefault();
            void submitCreateUser();
        });
    }
}

export function openCreateModal() {
    const bd = el('pvModalCreateBackdrop');
    if (!bd) return;
    wireCreateAutofill();
    displayNameTouched = false;
    mailNickTouched = false;

    const ids = [
        'pvCreateUpn',
        'pvCreateDisplayName',
        'pvCreateMailNick',
        'pvCreatePassword',
        'pvCreateGiven',
        'pvCreateSurname',
        'pvCreateMail'
    ];
    for (let i = 0; i < ids.length; i++) {
        const node = el(ids[i]);
        if (node) node.value = '';
    }
    const usage = el('pvCreateUsageLocation');
    if (usage) usage.value = 'AT';
    const f = el('pvCreateForcePw');
    if (f) f.checked = true;
    const e = el('pvCreateEnabled');
    if (e) e.checked = true;

    bd.classList.add('active');
    bd.setAttribute('aria-hidden', 'false');
    const focusEl = el('pvCreateGiven') || el('pvCreateUpn');
    if (focusEl) {
        setTimeout(function () {
            focusEl.focus();
        }, 30);
    }
}

export function closeCreateModal() {
    const bd = el('pvModalCreateBackdrop');
    if (!bd) return;
    bd.classList.remove('active');
    bd.setAttribute('aria-hidden', 'true');
}

export async function submitCreateUser() {
    const upn = readInputTrim(el('pvCreateUpn'));
    const displayName = readInputTrim(el('pvCreateDisplayName'));
    let mailNick = readInputTrim(el('pvCreateMailNick'));
    const password = String(el('pvCreatePassword')?.value || '');
    const givenName = readInputTrim(el('pvCreateGiven'));
    const surname = readInputTrim(el('pvCreateSurname'));
    const mail = readInputTrim(el('pvCreateMail'));
    let usageLocation = readInputTrim(el('pvCreateUsageLocation')).toUpperCase();
    const forcePw = el('pvCreateForcePw') ? el('pvCreateForcePw').checked : true;
    const enabled = el('pvCreateEnabled') ? el('pvCreateEnabled').checked : true;

    if (!upn || !displayName || !password) {
        host.toast('UPN, Anzeigename und Kennwort sind Pflichtfelder.');
        return;
    }
    mailNick = sanitizeMailNickname(mailNick, upn);
    if (!mailNick) {
        host.toast('Mail-Nickname ungültig oder leer.');
        return;
    }
    if (usageLocation && !/^[A-Z]{2}$/.test(usageLocation)) {
        host.toast('Nutzungsort: zwei Buchstaben, z. B. AT.');
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
    if (usageLocation) body.usageLocation = usageLocation;

    const sub = el('pvModalCreateSubmit');
    if (sub) sub.disabled = true;
    try {
        const token = await getGraphToken();
        const created = await graphJson('POST', '/users', token, body);
        const id = created && created.id ? created.id : null;
        host.appendLog(
            'Benutzer angelegt: ' +
                (created.userPrincipalName || upn) +
                (usageLocation ? ' · Nutzungsort ' + usageLocation : ''),
            'ok'
        );
        host.toast(
            usageLocation
                ? 'Konto angelegt. Nutzungsort ' + usageLocation + ' – bitte Lizenz zuweisen.'
                : 'Konto angelegt. Als Nächstes: Nutzungsort und Lizenz zuweisen.'
        );
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
    const hard = el('pvDeleteHard') && el('pvDeleteHard').checked;
    const warn = el('pvDeleteHardWarn');
    const intro = el('pvDeleteSoftIntro');
    const sub = el('pvModalDeleteSubmit');
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
        const hardChk = el('pvDeleteHard');
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
        const hardChk = el('pvDeleteHard');
        if (hardChk) hardChk.disabled = false;
    }
    const bd = el('pvModalDeleteBackdrop');
    const echo = el('pvDeleteUpnEcho');
    const inp = el('pvDeleteConfirmInput');
    const sub = el('pvModalDeleteSubmit');
    const hardChk2 = el('pvDeleteHard');
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
    const bd = el('pvModalDeleteBackdrop');
    if (!bd) return;
    bd.classList.remove('active');
    bd.setAttribute('aria-hidden', 'true');
}

export function syncDeleteConfirmButton() {
    const u = host.getSelectedUser();
    const inp = el('pvDeleteConfirmInput');
    const sub = el('pvModalDeleteSubmit');
    if (!sub || !inp || !u) return;
    const ok = readInputTrim(inp) === String(u.userPrincipalName || '').trim();
    sub.disabled = !ok;
}

export async function submitDeleteUser() {
    const u = host.getSelectedUser();
    if (!u) return;
    const inp = el('pvDeleteConfirmInput');
    if (!inp || readInputTrim(inp) !== String(u.userPrincipalName || '').trim()) {
        host.toast('UPN stimmt nicht überein.');
        return;
    }
    const hard = el('pvDeleteHard') && el('pvDeleteHard').checked;
    if (hard && u.onPremisesSyncEnabled === true) {
        host.toast('AD‑Sync‑Konten können nicht aus der Cloud gelöscht werden – bitte lokal im AD löschen.');
        return;
    }
    const sub = el('pvModalDeleteSubmit');
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
        const hint = el('pvHint');
        const detail = el('pvDetail');
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
