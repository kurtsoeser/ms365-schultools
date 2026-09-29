/**
 * Profil-Tab UI (Analyse 02 Phase B).
 */
import { pv } from './personen-verwaltung-state.js';
import { getGraphToken, graphJson } from './personen-verwaltung-graph.js';
import { graphErrorFriendly, formatPhones, formatDate, userTypeLabel } from './personen-verwaltung-logic.js';

const USER_REFRESH_SELECT = pv.USER_REFRESH_SELECT;

/** @type {any} */
let host = null;

export function init(a) {
    host = a;
}

export function dispVal(v) {
    if (v === undefined || v === null || v === '') return '–';
    return String(v);
}

export function addProfileTextField(root, label, fieldKey, value, editable, fullWidth) {
    const wrap = document.createElement('div');
    wrap.className =
        'field ' + (editable ? 'field-editable' : 'field-readonly') + (fullWidth ? ' field-full' : '');
    const lab = document.createElement('label');
    lab.setAttribute('for', 'pv_f_' + fieldKey);
    lab.textContent = label;
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.id = 'pv_f_' + fieldKey;
    inp.dataset.pvField = fieldKey;
    inp.readOnly = !editable;
    inp.autocomplete = 'off';
    if (editable && (value === undefined || value === null || value === '')) {
        inp.value = '';
        inp.placeholder = '–';
    } else {
        inp.value = dispVal(value);
    }
    wrap.appendChild(lab);
    wrap.appendChild(inp);
    root.appendChild(wrap);
}

export function addProfileAccountEnabled(root, u, editable) {
    const wrap = document.createElement('div');
    wrap.className = 'field ' + (editable ? 'field-editable' : 'field-readonly');
    const lab = document.createElement('label');
    lab.setAttribute('for', 'pv_f_accountEnabled');
    lab.textContent = 'Konto aktiv';
    wrap.appendChild(lab);
    if (editable) {
        const row = document.createElement('div');
        row.className = 'pv-checkbox-row';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.id = 'pv_f_accountEnabled';
        cb.dataset.pvField = 'accountEnabled';
        cb.checked = u.accountEnabled !== false;
        const l2 = document.createElement('label');
        l2.htmlFor = 'pv_f_accountEnabled';
        l2.style.margin = '0';
        l2.style.fontWeight = '600';
        l2.textContent = 'Konto ist aktiviert';
        row.appendChild(cb);
        row.appendChild(l2);
        wrap.appendChild(row);
    } else {
        const inp = document.createElement('input');
        inp.type = 'text';
        inp.readOnly = true;
        inp.value = u.accountEnabled === false ? 'Nein' : u.accountEnabled === true ? 'Ja' : '–';
        if (inp.value === '–') inp.style.color = 'var(--muted)';
        wrap.appendChild(inp);
    }
    root.appendChild(wrap);
}

export function renderProfileTab(u, editable) {
    const root = document.getElementById('pvProfileFields');
    if (!root) return;
    root.replaceChildren();
    if (!u) return;

    addProfileTextField(root, 'Anzeigename', 'displayName', u.displayName, editable, false);
    addProfileTextField(root, 'Alias (Mail-Nickname)', 'mailNickname', u.mailNickname, editable, false);
    addProfileTextField(root, 'Vorname', 'givenName', u.givenName, editable, false);
    addProfileTextField(root, 'Nachname', 'surname', u.surname, editable, false);
    addProfileTextField(
        root,
        'Benutzername (UPN)' + (String(u.userType).toLowerCase() === 'guest' ? ' (Gast)' : ''),
        'userPrincipalName',
        u.userPrincipalName,
        editable,
        false
    );
    addProfileTextField(root, 'E-Mail (SMTP)', 'mail', u.mail, editable, false);
    addProfileTextField(root, 'Position', 'jobTitle', u.jobTitle, editable, false);
    addProfileTextField(root, 'Abteilung', 'department', u.department, editable, false);
    addProfileTextField(root, 'Firma', 'companyName', u.companyName, editable, false);
    addProfileTextField(root, 'Bürostandort', 'officeLocation', u.officeLocation, editable, false);
    addProfileTextField(root, 'Straße', 'streetAddress', u.streetAddress, editable, true);
    addProfileTextField(root, 'PLZ', 'postalCode', u.postalCode, editable, false);
    addProfileTextField(root, 'Ort', 'city', u.city, editable, false);
    addProfileTextField(root, 'Land', 'country', u.country, editable, false);
    addProfileTextField(root, 'Mobiltelefon', 'mobilePhone', u.mobilePhone, editable, false);
    const bp0 = u.businessPhones && u.businessPhones[0] ? u.businessPhones[0] : '';
    addProfileTextField(root, 'Geschäftstelefon (1. Zeile)', 'businessPhone0', bp0, editable, false);
    addProfileTextField(root, 'Sprache (z. B. de-AT)', 'preferredLanguage', u.preferredLanguage, editable, false);
    addProfileAccountEnabled(root, u, editable);
    addProfileTextField(root, 'Kontotyp', '_userType', userTypeLabel(u.userType), false, false);
    addProfileTextField(root, 'Objekt-ID', '_id', u.id, false, true);
    addProfileTextField(root, 'Erstellt', '_created', formatDate(u.createdDateTime), false, false);
    const licSum = userLicenseSummary(u);
    const licText = licSum
        ? licSum.hasAny
            ? (licSum.licenses || [])
                  .map(function (l) {
                      return l.name || l.shortLabel;
                  })
                  .filter(Boolean)
                  .join(', ')
            : 'Keine'
        : '–';
    addProfileTextField(root, 'Lizenzen (Übersicht)', '_licenses', licText, false, true);
}

export function readInputTrim(el) {
    if (!el) return '';
    return String(el.value || '').trim();
}

export function buildPatchFromForm(u) {
    const root = document.getElementById('pvProfileFields');
    if (!root) return null;

    function get(field) {
        const el = root.querySelector('[data-pv-field="' + field + '"]');
        if (!el) return undefined;
        if (el.type === 'checkbox') return !!el.checked;
        const t = readInputTrim(el);
        return t === '' || t === '–' ? '' : t;
    }

    const patch = {};
    const strFields = [
        'displayName',
        'givenName',
        'surname',
        'userPrincipalName',
        'mail',
        'mailNickname',
        'jobTitle',
        'department',
        'companyName',
        'officeLocation',
        'streetAddress',
        'city',
        'postalCode',
        'country',
        'mobilePhone',
        'preferredLanguage'
    ];
    for (let i = 0; i < strFields.length; i++) {
        const k = strFields[i];
        let nv = get(k);
        if (nv === undefined) continue;
        if (k === 'mailNickname' && nv) {
            nv = sanitizeMailNickname(nv, '');
        }
        const ov = u[k] == null ? '' : String(u[k]);
        if (String(nv) !== ov) {
            patch[k] = nv === '' ? null : nv;
        }
    }

    const accEl = root.querySelector('[data-pv-field="accountEnabled"]');
    if (accEl && accEl.type === 'checkbox') {
        const nv = !!accEl.checked;
        const ov = u.accountEnabled !== false;
        if (nv !== ov) patch.accountEnabled = nv;
    }

    const bpNew = get('businessPhone0');
    if (bpNew !== undefined) {
        const ov = u.businessPhones && u.businessPhones[0] ? String(u.businessPhones[0]) : '';
        if (String(bpNew) !== ov) {
            patch.businessPhones = bpNew === '' ? [] : [bpNew];
        }
    }

    return patch;
}

export async function refreshUserFromGraph(token, userId) {
    const path =
        '/users/' + encodeURIComponent(userId) + '?$select=' + encodeURIComponent(USER_REFRESH_SELECT);
    return graphJson('GET', path, token, undefined);
}

export function mergeUserIntoList(updated) {
    const withFlags = applyAdFlagsToUsers([updated])[0] || updated;
    const idx = pv.loadedUsers.findIndex(function (x) {
        return x.id === withFlags.id;
    });
    if (idx === -1) {
        pv.loadedUsers.push(withFlags);
    } else {
        pv.loadedUsers[idx] = Object.assign({}, pv.loadedUsers[idx], withFlags);
    }
    host.refreshDepartmentFilter();
    host.refreshLicenseFilter();
    host.updateStatsPanel();
}

export async function saveProfilePatch() {
    const u = host.getSelectedUser();
    if (!u) return;
    const root = document.getElementById('pvProfileFields');
    const dnEl = root && root.querySelector('[data-pv-field="displayName"]');
    if (dnEl && readInputTrim(dnEl) === '') {
        host.toast('Anzeigename darf nicht leer sein.');
        return;
    }
    const upnEl = root && root.querySelector('[data-pv-field="userPrincipalName"]');
    if (upnEl && readInputTrim(upnEl) === '') {
        host.toast('UPN darf nicht leer sein.');
        return;
    }

    const patch = buildPatchFromForm(u);
    if (!patch || Object.keys(patch).length === 0) {
        host.toast('Keine Änderungen.');
        return;
    }
    const saveBtns = [
        document.getElementById('pvBtnSave'),
        document.getElementById('pvBtnSaveBottom')
    ].filter(Boolean);
    saveBtns.forEach(function (b) {
        b.disabled = true;
    });
    try {
        const token = await getGraphToken();
        await graphJson('PATCH', '/users/' + encodeURIComponent(u.id), token, patch);
        const fresh = await host.refreshUserFromGraph(token, u.id);
        host.mergeUserIntoList(fresh);
        host.appendLog('Profil gespeichert (PATCH).', 'ok');
        host.toast('Gespeichert.');
        pv.profileEditMode = true;
        host.updateDetailActionButtons();
        host.renderProfileTab(fresh, true);
        host.renderUserTree();
    } catch (e) {
        const msg = graphErrorFriendly(e);
        host.appendLog('PATCH: ' + msg, 'err');
        host.toast(msg);
    } finally {
        saveBtns.forEach(function (b) {
            b.disabled = false;
        });
    }
}

export async function resetProfileFromGraph() {
    const u = host.getSelectedUser();
    if (!u) return;
    try {
        const token = await getGraphToken();
        const fresh = await host.refreshUserFromGraph(token, u.id);
        host.mergeUserIntoList(fresh);
        pv.profileEditMode = true;
        host.renderProfileTab(fresh, true);
        host.renderUserTree();
        host.appendLog('Profil neu geladen.', 'ok');
        host.toast('Profil neu geladen.');
    } catch (e) {
        const msg = graphErrorFriendly(e);
        host.appendLog('Profil laden: ' + msg, 'err');
        host.toast(msg);
        host.renderProfileTab(u, true);
    }
}
