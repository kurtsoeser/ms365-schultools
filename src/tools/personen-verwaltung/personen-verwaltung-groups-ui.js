/**
 * Gruppen-Tab UI (Analyse 02 Phase B).
 */
import { pv } from './personen-verwaltung-state.js';
import { GRAPH_SCOPES, getGraphToken, graphJson, odataEscape } from './personen-verwaltung-graph.js';
import {
    graphErrorFriendly,
    norm,
    isGuid,
    isDuplicateMemberError,
    groupTypeLabel
} from './personen-verwaltung-logic.js';
import { dlgConfirm } from '../../shared/utils/dialog.js';

const GROUP_MEMBEROF_SELECT = pv.GROUP_MEMBEROF_SELECT;

/** @type {any} */
let host = null;

export function init(a) {
    host = a;
}

export function renderGroupsTable(groups) {
    const tbody = document.getElementById('pvGroupsTbody');
    if (!tbody) return;
    tbody.replaceChildren();

    if (!groups || !groups.length) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.colSpan = 4;
        td.style.color = '#6c757d';
        td.textContent = 'Keine direkten Gruppenmitgliedschaften gefunden.';
        tr.appendChild(td);
        tbody.appendChild(tr);
        return;
    }

    const sorted = groups.slice().sort(function (a, b) {
        return compareStrings(a.displayName, b.displayName);
    });

    for (let i = 0; i < sorted.length; i++) {
        const g = sorted[i];
        const tr = document.createElement('tr');
        const tdN = document.createElement('td');
        tdN.textContent = g.displayName || '–';
        const tdM = document.createElement('td');
        tdM.textContent = g.mail || g.mailNickname || '–';
        tdM.style.wordBreak = 'break-all';
        tdM.style.fontSize = '0.9em';
        const tdT = document.createElement('td');
        tdT.textContent = groupTypeLabel(g);
        tdT.style.fontSize = '0.88em';
        const tdAct = document.createElement('td');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn small-btn';
        btn.setAttribute('data-pv-group-remove', g.id || '');
        btn.textContent = 'Entfernen';
        tdAct.appendChild(btn);
        tr.appendChild(tdN);
        tr.appendChild(tdM);
        tr.appendChild(tdT);
        tr.appendChild(tdAct);
        tbody.appendChild(tr);
    }
}



export function memberGroupIds() {
    const set = new Set();
    (pv.cachedGroupsForSelection || []).forEach(function (g) {
        const id = String((g && g.id) || '').toLowerCase();
        if (id) set.add(id);
    });
    return set;
}

export function fillGroupSearchResults(groups) {
    const container = document.getElementById('pvGroupSearchResults');
    if (!container) return;
    container.replaceChildren();
    const already = memberGroupIds();
    const filtered = (groups || []).filter(function (g) {
        return g && g.id && !already.has(String(g.id).toLowerCase());
    });
    if (!filtered.length) {
        const hint = document.createElement('span');
        hint.className = 'pv-group-checklist-hint';
        hint.textContent = groups && groups.length ? '(alle Treffer bereits Mitglied)' : '(keine Treffer)';
        container.appendChild(hint);
        return;
    }
    // Alle auswählen-Zeile
    const selAllRow = document.createElement('label');
    selAllRow.className = 'pv-group-checklist-selectall';
    const selAllCb = document.createElement('input');
    selAllCb.type = 'checkbox';
    selAllCb.style.width = '16px';
    selAllCb.style.height = '16px';
    selAllCb.style.cursor = 'pointer';
    selAllCb.style.accentColor = 'var(--brand1)';
    selAllCb.setAttribute('aria-label', 'Alle auswählen');
    selAllRow.appendChild(selAllCb);
    const selAllTxt = document.createElement('span');
    selAllTxt.textContent = 'Alle auswählen (' + filtered.length + ')';
    selAllRow.appendChild(selAllTxt);
    container.appendChild(selAllRow);

    filtered.forEach(function (g) {
        const mail = g.mail || g.mailNickname || '';
        const label = document.createElement('label');
        label.className = 'pv-group-checklist-item';
        label.setAttribute('data-pv-group-id', g.id);
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = g.id;
        cb.addEventListener('change', function () {
            label.classList.toggle('is-checked', cb.checked);
            // Alle-auswählen-Checkbox aktualisieren
            const allCbs = container.querySelectorAll('.pv-group-checklist-item input[type="checkbox"]');
            const checkedCount = container.querySelectorAll('.pv-group-checklist-item input[type="checkbox"]:checked').length;
            selAllCb.indeterminate = checkedCount > 0 && checkedCount < allCbs.length;
            selAllCb.checked = checkedCount === allCbs.length;
        });
        label.appendChild(cb);
        const txt = document.createElement('span');
        txt.textContent = (g.displayName || g.id) + (mail ? ' · ' + mail : '') + ' · ' + groupTypeLabel(g);
        label.appendChild(txt);
        container.appendChild(label);
    });

    selAllCb.addEventListener('change', function () {
        const allCbs = container.querySelectorAll('.pv-group-checklist-item input[type="checkbox"]');
        allCbs.forEach(function (cb) {
            cb.checked = selAllCb.checked;
            cb.closest('.pv-group-checklist-item').classList.toggle('is-checked', selAllCb.checked);
        });
    });
}

export async function searchDirectoryGroups(token, queryRaw) {
    const q = String(queryRaw || '').trim();
    if (!q) return [];
    const select = 'id,displayName,mail,mailNickname,groupTypes,securityEnabled,mailEnabled';
    if (isGuid(q)) {
        try {
            const g = await graphJson(
                'GET',
                '/groups/' + encodeURIComponent(q) + '?$select=' + encodeURIComponent(select),
                token
            );
            return g && g.id ? [g] : [];
        } catch {
            return [];
        }
    }
    try {
        const phrase = q.replace(/"/g, '\\"').replace(/\r?\n/g, ' ').trim();
        const aqs =
            '(displayName:' + phrase + ' OR mail:' + phrase + ' OR mailNickname:' + phrase + ')';
        const path =
            '/groups?$search=' +
            encodeURIComponent('"' + aqs + '"') +
            '&$select=' +
            encodeURIComponent(select) +
            '&$top=25';
        const data = await graphJson('GET', path, token, undefined, { ConsistencyLevel: 'eventual' });
        return Array.isArray(data.value) ? data.value : [];
    } catch {
        // Fallback ohne $search
    }
    const esc = odataEscape(q);
    const filter =
        "startswith(displayName,'" +
        esc +
        "') or startswith(mailNickname,'" +
        esc +
        "') or startswith(mail,'" +
        esc +
        "')";
    const path =
        '/groups?$filter=' +
        encodeURIComponent(filter) +
        '&$select=' +
        encodeURIComponent(select) +
        '&$top=25';
    const data = await graphJson('GET', path, token);
    return Array.isArray(data.value) ? data.value : [];
}

export async function fetchUserGroups(token, userId) {
    const path =
        '/users/' +
        encodeURIComponent(userId) +
        '/memberOf/microsoft.graph.group?$select=' +
        encodeURIComponent(GROUP_MEMBEROF_SELECT) +
        '&$top=999';
    return fetchAllPages(token, path, undefined);
}

export async function loadGroupsForSelected() {
    const prog = document.getElementById('pvGroupsProgress');
    if (!pv.selectedUserId) return;
    if (prog) prog.textContent = 'Lade Gruppen …';

    const tbody = document.getElementById('pvGroupsTbody');
    if (tbody) {
        tbody.replaceChildren();
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.colSpan = 4;
        td.style.color = '#6c757d';
        td.textContent = 'Lade …';
        tr.appendChild(td);
        tbody.appendChild(tr);
    }

    try {
        const token = await getGraphToken();
        const groups = await fetchUserGroups(token, pv.selectedUserId);
        pv.cachedGroupsForSelection = groups;
        renderGroupsTable(groups);
        if (prog) prog.textContent = groups.length ? groups.length + ' Gruppe(n).' : 'Keine Einträge.';
        host.appendLog('Gruppen für ausgewählte Person: ' + groups.length, 'ok');
    } catch (e) {
        pv.cachedGroupsForSelection = [];
        renderGroupsTable([]);
        const msg = graphErrorFriendly(e);
        if (prog) prog.textContent = 'Fehler: ' + msg;
        host.appendLog('Gruppen laden: ' + msg, 'err');
        host.toast('Gruppen: ' + msg);
    }
}

export async function searchGroupsForAdd() {
    const inp = document.getElementById('pvGroupSearch');
    const q = inp && inp.value ? String(inp.value).trim() : '';
    if (!q) {
        host.toast('Bitte einen Gruppennamen, Alias oder eine ID eingeben.');
        return;
    }
    const btn = document.getElementById('pvGroupSearchBtn');
    const status = document.getElementById('pvGroupsProgress');
    if (btn) btn.disabled = true;
    try {
        const token = await getGraphToken();
        const list = await searchDirectoryGroups(token, q);
        fillGroupSearchResults(list);
        if (status) {
            status.textContent = list.length
                ? 'Suche: ' + list.length + ' Treffer.'
                : 'Suche: keine Treffer.';
        }
    } catch (e) {
        const msg = graphErrorFriendly(e);
        fillGroupSearchResults([]);
        if (status) status.textContent = 'Suche: ' + msg;
        host.toast(msg);
    } finally {
        if (btn) btn.disabled = false;
    }
}

export async function addSelectedUserToGroup() {
    const u = host.getSelectedUser();
    const container = document.getElementById('pvGroupSearchResults');
    const checkedBoxes = container
        ? Array.from(container.querySelectorAll('.pv-group-checklist-item input[type="checkbox"]:checked'))
        : [];
    if (!u || !checkedBoxes.length) {
        host.toast('Bitte zuerst mindestens eine Gruppe aus den Treffern auswählen.');
        return;
    }
    if (pv.groupBusy) return;
    const groups = checkedBoxes.map(function (cb) {
        const lbl = cb.closest('.pv-group-checklist-item');
        return { id: cb.value, label: lbl ? (lbl.querySelector('span') ? lbl.querySelector('span').textContent : cb.value) : cb.value };
    });
    const groupNames = groups.map(function (g) { return '· ' + g.label; }).join('\n');
    if (
        !(await dlgConfirm(
            'Diese Person zu ' + groups.length + ' Gruppe(n) hinzufügen?\n\n' +
                (u.displayName || u.userPrincipalName || '') +
                '\n\n' + groupNames,
            { title: 'Zu Gruppen hinzufügen', okText: 'Hinzufügen' }
        ))
    ) {
        return;
    }
    const btn = document.getElementById('pvGroupAddBtn');
    pv.groupBusy = true;
    if (btn) btn.disabled = true;
    try {
        const token = await getGraphToken();
        let ok = 0, fail = 0;
        for (const g of groups) {
            try {
                await graphJson('POST', '/groups/' + encodeURIComponent(g.id) + '/members/$ref', token, {
                    '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + u.id
                });
                host.appendLog('Mitglied hinzugefügt: ' + (u.displayName || u.id) + ' → ' + g.label, 'ok');
                ok++;
            } catch (e) {
                if (isDuplicateMemberError(e)) {
                    host.appendLog('Bereits Mitglied: ' + g.label, 'warn');
                    ok++;
                } else {
                    host.appendLog('Fehler bei ' + g.label + ': ' + graphErrorFriendly(e), 'err');
                    fail++;
                }
            }
        }
        host.toast(ok + ' Gruppe(n) hinzugefügt' + (fail ? ', ' + fail + ' Fehler.' : '.'));
        pv.cachedGroupsForSelection = null;
        if (container) {
            container.replaceChildren();
            const hint = document.createElement('span');
            hint.className = 'pv-group-checklist-hint';
            hint.textContent = '(zuerst suchen)';
            container.appendChild(hint);
        }
        await host.loadGroupsForSelected();
    } catch (e) {
        const msg = graphErrorFriendly(e);
        host.appendLog('Gruppe hinzufügen: ' + msg, 'err');
        host.toast(msg);
    } finally {
        pv.groupBusy = false;
        if (btn) btn.disabled = false;
    }
}

export async function removeUserFromGroup(groupIdRaw) {
    const u = host.getSelectedUser();
    const groupId = String(groupIdRaw || '').trim();
    if (!u || !groupId || pv.groupBusy) return;
    const g = (pv.cachedGroupsForSelection || []).find(function (x) {
        return x && x.id === groupId;
    });
    const label = (g && (g.displayName || g.mail)) || groupId;
    if (
        !(await dlgConfirm(
            'Mitgliedschaft entfernen?\n\n' +
                (u.displayName || u.userPrincipalName || '') +
                '\n← ' +
                label,
            { title: 'Aus Gruppe entfernen', okText: 'Entfernen', danger: true }
        ))
    ) {
        return;
    }
    pv.groupBusy = true;
    try {
        const token = await getGraphToken();
        await graphJson(
            'DELETE',
            '/groups/' + encodeURIComponent(groupId) + '/members/' + encodeURIComponent(u.id) + '/$ref',
            token
        );
        host.appendLog('Mitglied entfernt: ' + (u.displayName || u.id) + ' ← ' + label, 'ok');
        host.toast('Aus der Gruppe entfernt.');
        pv.cachedGroupsForSelection = null;
        await host.loadGroupsForSelected();
    } catch (e) {
        const msg = graphErrorFriendly(e);
        host.appendLog('Gruppe entfernen: ' + msg, 'err');
        host.toast(msg);
    } finally {
        pv.groupBusy = false;
    }
}


/* renderLicenseTab… ausgelagert */
