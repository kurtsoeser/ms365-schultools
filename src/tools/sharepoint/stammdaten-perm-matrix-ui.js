/**
 * UI: Entra-Gruppen als Zeilen, Listen als Spalten (Stammdaten-Berechtigungen).
 */
import { pickEntraGroup } from '../../shared/entra-group-picker.js';
import { savePermissionsConfig, grantRowsForUi, loadPermissionsConfig } from './stammdaten-liste-permissions.js';
import {
    PERM_LEVEL_OPTIONS,
    STAMMDATEN_PERM_MATRIX_COLS,
    normalizeGrantRows
} from './stammdaten-liste-perm-matrix.js';

function levelToSelectValue(level) {
    return level === null || level === undefined ? '' : String(level);
}

function selectValueToLevel(value) {
    const v = String(value || '').trim();
    if (!v) return null;
    return v;
}

function toastErr(e) {
    const msg = e && e.message ? e.message : String(e);
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else window.alert(msg);
}

function buildCellSelect(listKey, level) {
    const sel = document.createElement('select');
    sel.className = 'sps-perm-matrix__select';
    sel.setAttribute('data-sps-perm-list', listKey);
    sel.title = 'Berechtigung für diese Liste';
    const opt0 = document.createElement('option');
    opt0.value = '';
    opt0.textContent = '—';
    sel.appendChild(opt0);
    PERM_LEVEL_OPTIONS.forEach(function (o) {
        const opt = document.createElement('option');
        opt.value = o.value;
        opt.textContent = o.label;
        sel.appendChild(opt);
    });
    sel.value = levelToSelectValue(level);
    return sel;
}

function readRowFromTr(tr) {
    const labelEl = tr.querySelector('[data-sps-grant-label]');
    const idEl = tr.querySelector('[data-sps-grant-id]');
    const groupLabel = labelEl ? String(labelEl.value || '').trim() : '';
    const groupId = idEl ? String(idEl.value || '').trim() : '';
    /** @type {Record<string, string|null>} */
    const cells = {};
    tr.querySelectorAll('select[data-sps-perm-list]').forEach(function (sel) {
        const key = String(sel.getAttribute('data-sps-perm-list') || '').trim();
        if (!key) return;
        cells[key] = selectValueToLevel(sel.value);
    });
    return { groupId, groupLabel, cells };
}

function persistFromTbody(tbody) {
    const rows = [];
    tbody.querySelectorAll('tr[data-sps-grant-row]').forEach(function (tr) {
        rows.push(readRowFromTr(tr));
    });
    const grantRows = normalizeGrantRows(rows);
    savePermissionsConfig({ grantRows });
    return grantRows;
}

function renderGroupCell(tr, row) {
    const td = document.createElement('td');
    td.className = 'sps-perm-matrix__group';
    const wrap = document.createElement('div');
    wrap.className = 'sps-perm-matrix__group-inner';
    const label = document.createElement('input');
    label.type = 'text';
    label.readOnly = true;
    label.className = 'sps-perm-matrix__group-label';
    label.setAttribute('data-sps-grant-label', '1');
    label.placeholder = 'Gruppe wählen …';
    label.value = row.groupLabel || '';
    const hidden = document.createElement('input');
    hidden.type = 'hidden';
    hidden.setAttribute('data-sps-grant-id', '1');
    hidden.value = row.groupId || '';
    const pick = document.createElement('button');
    pick.type = 'button';
    pick.className = 'btn btn-sm';
    pick.setAttribute('data-sps-grant-pick', '1');
    pick.innerHTML = '<i class="bi bi-search"></i>';
    pick.title = 'Entra-Gruppe wählen';
    const clear = document.createElement('button');
    clear.type = 'button';
    clear.className = 'btn btn-sm alt';
    clear.setAttribute('data-sps-grant-clear', '1');
    clear.innerHTML = '<i class="bi bi-x-lg"></i>';
    clear.title = 'Gruppe entfernen';
    wrap.appendChild(label);
    wrap.appendChild(hidden);
    wrap.appendChild(pick);
    wrap.appendChild(clear);
    td.appendChild(wrap);
    tr.appendChild(td);
}

function appendGrantRow(tbody, row) {
    const tr = document.createElement('tr');
    tr.setAttribute('data-sps-grant-row', '1');
    renderGroupCell(tr, row);
    STAMMDATEN_PERM_MATRIX_COLS.forEach(function (col) {
        const td = document.createElement('td');
        td.className = 'sps-perm-matrix__cell';
        const level = row.cells && row.cells[col.key] != null ? row.cells[col.key] : null;
        td.appendChild(buildCellSelect(col.key, level));
        tr.appendChild(td);
    });
    const tdAct = document.createElement('td');
    tdAct.className = 'sps-perm-matrix__actions';
    const rm = document.createElement('button');
    rm.type = 'button';
    rm.className = 'btn btn-sm alt';
    rm.setAttribute('data-sps-grant-remove', '1');
    rm.title = 'Zeile entfernen';
    rm.innerHTML = '<i class="bi bi-trash"></i>';
    tdAct.appendChild(rm);
    tr.appendChild(tdAct);
    tbody.appendChild(tr);
}

function renderMatrixHead(thead) {
    if (!thead) return;
    const tr = document.createElement('tr');
    const th0 = document.createElement('th');
    th0.textContent = 'Entra-Gruppe';
    tr.appendChild(th0);
    STAMMDATEN_PERM_MATRIX_COLS.forEach(function (col) {
        const th = document.createElement('th');
        th.textContent = col.label;
        tr.appendChild(th);
    });
    const thAct = document.createElement('th');
    thAct.className = 'sps-perm-matrix__actions-head';
    thAct.textContent = '';
    tr.appendChild(thAct);
    thead.innerHTML = '';
    thead.appendChild(tr);
}

function renderMatrixBody(tbody) {
    const rows = grantRowsForUi();
    tbody.innerHTML = '';
    rows.forEach(function (row) {
        appendGrantRow(tbody, row);
    });
}

function wireTbody(tbody) {
    if (!tbody || tbody.dataset.spsGrantWired === '1') return;
    tbody.dataset.spsGrantWired = '1';

    tbody.addEventListener('change', function (ev) {
        const t = ev.target;
        if (!t || !t.matches || !t.matches('select[data-sps-perm-list]')) return;
        persistFromTbody(tbody);
    });

    tbody.addEventListener('click', function (ev) {
        const pick = ev.target.closest('[data-sps-grant-pick]');
        if (pick) {
            const tr = pick.closest('tr[data-sps-grant-row]');
            if (!tr) return;
            pickEntraGroup({ title: 'Entra-Gruppe für Listen-Berechtigung' })
                .then(function (sel) {
                    if (!sel) return;
                    const labelEl = tr.querySelector('[data-sps-grant-label]');
                    const idEl = tr.querySelector('[data-sps-grant-id]');
                    if (labelEl) labelEl.value = sel.label || sel.displayName || '';
                    if (idEl) idEl.value = sel.id || '';
                    persistFromTbody(tbody);
                })
                .catch(toastErr);
            return;
        }
        const clear = ev.target.closest('[data-sps-grant-clear]');
        if (clear) {
            const tr = clear.closest('tr[data-sps-grant-row]');
            if (!tr) return;
            const labelEl = tr.querySelector('[data-sps-grant-label]');
            const idEl = tr.querySelector('[data-sps-grant-id]');
            if (labelEl) labelEl.value = '';
            if (idEl) idEl.value = '';
            persistFromTbody(tbody);
            return;
        }
        const rm = ev.target.closest('[data-sps-grant-remove]');
        if (rm) {
            const tr = rm.closest('tr[data-sps-grant-row]');
            if (!tr) return;
            tr.remove();
            persistFromTbody(tbody);
        }
    });
}

function matrixTableFromTbody(tbody) {
    return tbody ? tbody.closest('table.sps-perm-matrix') : null;
}

/**
 * @param {string} [tbodyId]
 */
export function initStammdatenPermMatrixUi(tbodyId) {
    const id = tbodyId || 'spsPermMatrixBody';
    const tbody = document.getElementById(id);
    if (!tbody) return;
    const table = matrixTableFromTbody(tbody);
    if (table) {
        const thead = table.querySelector('thead');
        renderMatrixHead(thead);
    }
    renderMatrixBody(tbody);
    wireTbody(tbody);

    const addBtnId = tbody.getAttribute('data-sps-add-btn') || 'spsPermMatrixAddRow';
    const addBtn = document.getElementById(addBtnId);
    if (addBtn && addBtn.dataset.spsGrantAddWired !== '1') {
        addBtn.dataset.spsGrantAddWired = '1';
        addBtn.addEventListener('click', function () {
            appendGrantRow(tbody, { groupId: '', groupLabel: '', cells: {} });
            persistFromTbody(tbody);
        });
    }
}

export function readGrantRowsFromDom(tbodyId) {
    const tbody = document.getElementById(tbodyId || 'spsPermMatrixBody');
    if (!tbody) return normalizeGrantRows(loadPermissionsConfig().grantRows);
    return persistFromTbody(tbody);
}

/** @deprecated Nutze readGrantRowsFromDom */
export function readListPermProfilesFromDom() {
    return readGrantRowsFromDom();
}
