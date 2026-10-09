/**
 * Matrix-UI: Gruppen/Personen × Planer-Rollen (Freistellungen-Setup Schritt 6).
 */
import { pickEntraGroup } from '../../shared/entra-group-picker.js';
import { pickEntraUser } from '../../shared/entra-user-picker.js';
import {
    FR_PLANNER_GRANT_COLS,
    normalizePlannerGrantRows,
    mergeDuplicateGrantRows
} from './freistellung-planner-grant-matrix.js';

function toastErr(e) {
    const msg = e && e.message ? e.message : String(e);
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else window.alert(msg);
}

function readRolesFromTr(tr) {
    /** @type {{ direktion: boolean, kv: boolean, schueler: boolean }} */
    const roles = { direktion: false, kv: false, schueler: false };
    tr.querySelectorAll('[data-fr-grant-role]').forEach(function (cb) {
        const key = String(cb.getAttribute('data-fr-grant-role') || '');
        if (key && cb.checked) roles[key] = true;
    });
    return roles;
}

function readRowFromTr(tr) {
    const type = String(tr.getAttribute('data-fr-principal-type') || 'group');
    const labelEl = tr.querySelector('[data-fr-grant-label]');
    const idEl = tr.querySelector('[data-fr-grant-id]');
    const mailEl = tr.querySelector('[data-fr-grant-mail]');
    return {
        principalType: type === 'user' ? 'user' : 'group',
        groupId: idEl ? String(idEl.value || '').trim() : '',
        groupLabel: labelEl ? String(labelEl.value || '').trim() : '',
        mail: mailEl ? String(mailEl.value || '').trim().toLowerCase() : '',
        displayName: labelEl ? String(labelEl.value || '').trim() : '',
        roles: readRolesFromTr(tr)
    };
}

function buildRoleCell(roleKey, checked) {
    const td = document.createElement('td');
    td.className = 'fr-grant-matrix__cell';
    const col = FR_PLANNER_GRANT_COLS.find((c) => c.key === roleKey);
    const wrap = document.createElement('label');
    wrap.className = 'fr-grant-matrix__role';
    wrap.title = col ? col.label + (col.hint ? ' – ' + col.hint : '') : roleKey;
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.className = 'fr-grant-matrix__cb';
    cb.setAttribute('data-fr-grant-role', roleKey);
    cb.setAttribute('aria-label', col ? col.label : roleKey);
    cb.checked = !!checked;
    wrap.appendChild(cb);
    td.appendChild(wrap);
    return td;
}

function renderPrincipalCell(tr, row) {
    const td = document.createElement('td');
    td.className = 'fr-grant-matrix__principal';
    const isUser = row.principalType === 'user';
    tr.setAttribute('data-fr-principal-type', isUser ? 'user' : 'group');
    const wrap = document.createElement('div');
    wrap.className = 'fr-grant-matrix__principal-inner';
    const badge = document.createElement('span');
    badge.className = 'fr-grant-matrix__type';
    badge.textContent = isUser ? 'Person' : 'Gruppe';
    const label = document.createElement('input');
    label.type = 'text';
    label.readOnly = true;
    label.className = 'fr-grant-matrix__label';
    label.setAttribute('data-fr-grant-label', '1');
    label.placeholder = isUser ? 'Person wählen …' : 'Gruppe wählen …';
    label.value = isUser ? row.displayName || row.mail || '' : row.groupLabel || '';
    const hiddenId = document.createElement('input');
    hiddenId.type = 'hidden';
    hiddenId.setAttribute('data-fr-grant-id', '1');
    hiddenId.value = isUser ? '' : row.groupId || '';
    const hiddenMail = document.createElement('input');
    hiddenMail.type = 'hidden';
    hiddenMail.setAttribute('data-fr-grant-mail', '1');
    hiddenMail.value = isUser ? row.mail || '' : '';
    const pick = document.createElement('button');
    pick.type = 'button';
    pick.className = 'btn btn-sm';
    pick.setAttribute('data-fr-grant-pick', '1');
    pick.innerHTML = '<i class="bi bi-search"></i>';
    pick.title = isUser ? 'Person wählen' : 'Entra-Gruppe wählen';
    const clear = document.createElement('button');
    clear.type = 'button';
    clear.className = 'btn btn-sm alt';
    clear.setAttribute('data-fr-grant-clear', '1');
    clear.innerHTML = '<i class="bi bi-x-lg"></i>';
    clear.title = 'Zeile leeren';
    wrap.appendChild(badge);
    wrap.appendChild(label);
    wrap.appendChild(hiddenId);
    wrap.appendChild(hiddenMail);
    wrap.appendChild(pick);
    wrap.appendChild(clear);
    td.appendChild(wrap);
    tr.appendChild(td);
}

export function appendGrantRow(tbody, row) {
    const tr = document.createElement('tr');
    tr.setAttribute('data-fr-grant-row', '1');
    renderPrincipalCell(tr, row || { principalType: 'group', roles: {} });
    FR_PLANNER_GRANT_COLS.forEach(function (col) {
        tr.appendChild(buildRoleCell(col.key, row && row.roles && row.roles[col.key]));
    });
    const tdAct = document.createElement('td');
    tdAct.className = 'fr-grant-matrix__actions';
    const rm = document.createElement('button');
    rm.type = 'button';
    rm.className = 'btn btn-sm alt';
    rm.setAttribute('data-fr-grant-remove', '1');
    rm.title = 'Zeile entfernen';
    rm.innerHTML = '<i class="bi bi-trash"></i>';
    tdAct.appendChild(rm);
    tr.appendChild(tdAct);
    tbody.appendChild(tr);
}

function renderHead(thead) {
    if (!thead) return;
    const tr = document.createElement('tr');
    const th0 = document.createElement('th');
    th0.textContent = 'Gruppe / Person';
    tr.appendChild(th0);
    FR_PLANNER_GRANT_COLS.forEach(function (col) {
        const th = document.createElement('th');
        th.textContent = col.label;
        th.title = col.hint;
        tr.appendChild(th);
    });
    const thA = document.createElement('th');
    thA.className = 'fr-grant-matrix__actions-head';
    tr.appendChild(thA);
    thead.innerHTML = '';
    thead.appendChild(tr);
}

export function readGrantRowsFromDom(tbodyId) {
    const tbody = document.getElementById(tbodyId || 'frPermMatrixBody');
    if (!tbody) return [];
    const rows = [];
    tbody.querySelectorAll('tr[data-fr-grant-row]').forEach(function (tr) {
        rows.push(readRowFromTr(tr));
    });
    return normalizePlannerGrantRows(rows);
}

/**
 * @param {string} tbodyId
 * @param {import('./freistellung-planner-grant-matrix.js').FrPlannerGrantRow[]} initialRows
 * @param {() => void} [onChange]
 */
export function initFreistellungPlannerGrantMatrixUi(tbodyId, initialRows, onChange) {
    const id = tbodyId || 'frPermMatrixBody';
    const tbody = document.getElementById(id);
    if (!tbody) return;
    const table = tbody.closest('table.fr-grant-matrix');
    renderHead(table ? table.querySelector('thead') : null);
    tbody.innerHTML = '';
    const rows = mergeDuplicateGrantRows(initialRows || []);
    if (rows.length) {
        rows.forEach(function (row) {
            appendGrantRow(tbody, row);
        });
    } else {
        appendGrantRow(tbody, { principalType: 'group', roles: { admin: true } });
        appendGrantRow(tbody, { principalType: 'group', roles: { direktion: true } });
        appendGrantRow(tbody, { principalType: 'group', roles: { kv: true } });
        appendGrantRow(tbody, { principalType: 'group', roles: { schueler: true } });
    }

    if (tbody.dataset.frGrantWired === '1') return;
    tbody.dataset.frGrantWired = '1';

    const persist = function () {
        if (typeof onChange === 'function') onChange();
    };

    tbody.addEventListener('change', function (ev) {
        const t = ev.target;
        if (!t || !t.matches) return;
        if (t.matches('[data-fr-grant-role]')) persist();
    });

    tbody.addEventListener('click', function (ev) {
        const pick = ev.target.closest('[data-fr-grant-pick]');
        if (pick) {
            const tr = pick.closest('tr[data-fr-grant-row]');
            if (!tr) return;
            const isUser = tr.getAttribute('data-fr-principal-type') === 'user';
            if (isUser) {
                pickEntraUser({ title: 'Person für Freistellungs-Planer' })
                    .then(function (sel) {
                        if (!sel) return;
                        const mail = String(sel.mail || sel.userPrincipalName || '').trim().toLowerCase();
                        const labelEl = tr.querySelector('[data-fr-grant-label]');
                        const mailEl = tr.querySelector('[data-fr-grant-mail]');
                        if (labelEl) labelEl.value = sel.displayName || mail;
                        if (mailEl) mailEl.value = mail;
                        persist();
                    })
                    .catch(toastErr);
            } else {
                pickEntraGroup({ title: 'Entra-Gruppe für Freistellungs-Planer' })
                    .then(function (sel) {
                        if (!sel) return;
                        const labelEl = tr.querySelector('[data-fr-grant-label]');
                        const idEl = tr.querySelector('[data-fr-grant-id]');
                        if (labelEl) labelEl.value = sel.label || sel.displayName || '';
                        if (idEl) idEl.value = sel.id || '';
                        persist();
                    })
                    .catch(toastErr);
            }
            return;
        }
        const clear = ev.target.closest('[data-fr-grant-clear]');
        if (clear) {
            const tr = clear.closest('tr[data-fr-grant-row]');
            if (!tr) return;
            tr.querySelectorAll('[data-fr-grant-label], [data-fr-grant-id], [data-fr-grant-mail]').forEach(function (el) {
                el.value = '';
            });
            persist();
            return;
        }
        const rm = ev.target.closest('[data-fr-grant-remove]');
        if (rm) {
            const tr = rm.closest('tr[data-fr-grant-row]');
            if (tr) tr.remove();
            persist();
        }
    });

    const addGroupBtn = document.getElementById('frPermMatrixAddGroup');
    if (addGroupBtn && addGroupBtn.dataset.frGrantAddWired !== '1') {
        addGroupBtn.dataset.frGrantAddWired = '1';
        addGroupBtn.addEventListener('click', function () {
            appendGrantRow(tbody, { principalType: 'group', roles: {} });
            persist();
        });
    }
    const addPersonBtn = document.getElementById('frPermMatrixAddPerson');
    if (addPersonBtn && addPersonBtn.dataset.frGrantAddWired !== '1') {
        addPersonBtn.dataset.frGrantAddWired = '1';
        addPersonBtn.addEventListener('click', function () {
            appendGrantRow(tbody, { principalType: 'user', roles: {} });
            persist();
        });
    }
}
