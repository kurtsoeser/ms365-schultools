/**
 * Suche/Filter/Baum/Stats (Analyse 02 Phase B).
 */
import { pv } from './personen-verwaltung-state.js';
import {
    norm,
    compareStrings,
    readSortFromSelect,
    userTypeLabel,
    userLicenseSummary,
    Lic
} from './personen-verwaltung-logic.js';


/** @type {any} */
let host = null;

export function init(a) {
    host = a;
}

export function getVisibleRows() {
    const filterInp = document.getElementById('pvFilterText');
    const q = filterInp && filterInp.value ? norm(filterInp.value) : '';

    const typeSel = document.getElementById('pvFilterUserType');
    const typeVal = typeSel && typeSel.value ? String(typeSel.value) : '';

    const accSel = document.getElementById('pvFilterAccount');
    const accVal = accSel && accSel.value !== '' ? String(accSel.value) : '';

    const depSel = document.getElementById('pvFilterDepartment');
    const depVal = depSel && depSel.value ? String(depSel.value) : '';

    const licSel = document.getElementById('pvFilterLicense');
    const licVal = licSel && licSel.value ? String(licSel.value) : '';

    const adSel = document.getElementById('pvFilterAdSync');
    const adVal = adSel && adSel.value ? String(adSel.value) : '';

    let rows = pv.loadedUsers.slice();

    if (typeVal) {
        rows = rows.filter(function (u) {
            return String(u.userType || '') === typeVal;
        });
    }

    if (accVal === '1') {
        rows = rows.filter(function (u) {
            return u.accountEnabled === true;
        });
    } else if (accVal === '0') {
        rows = rows.filter(function (u) {
            return u.accountEnabled === false;
        });
    }

    if (depVal) {
        rows = rows.filter(function (u) {
            return String(u.department || '').trim() === depVal;
        });
    }

    if (licVal) {
        const api = Lic();
        if (api && typeof api.userMatchesLicenseFilter === 'function') {
            rows = rows.filter(function (u) {
                return api.userMatchesLicenseFilter(u, licVal);
            });
        }
    }

    if (adVal === 'adSync') {
        rows = rows.filter(function (u) {
            return u.onPremisesSyncEnabled === true;
        });
    } else if (adVal === 'cloud') {
        rows = rows.filter(function (u) {
            return u.onPremisesSyncEnabled !== true;
        });
    } else if (adVal === 'flagged') {
        rows = rows.filter(function (u) {
            return !!u.adFlagged;
        });
    }

    if (q) {
        rows = rows.filter(function (u) {
            const blob = [
                u.displayName,
                u.givenName,
                u.surname,
                u.mail,
                u.userPrincipalName,
                u.department,
                u.jobTitle,
                u.id,
                u.officeLocation,
                u.companyName,
                u.onPremisesSamAccountName,
                u.onPremisesDomainName,
                u.adFlagNote
            ]
                .map(function (x) {
                    return norm(x);
                })
                .join(' ');
            const sum = userLicenseSummary(u);
            const lic = sum && sum.primaryLabel ? norm(sum.primaryLabel) : '';
            return blob.indexOf(q) !== -1 || (lic && lic.indexOf(q) !== -1);
        });
    }

    const sortState = readSortFromSelect();
    const key = sortState.key || 'displayName';
    const dir = sortState.dir === 'desc' ? -1 : 1;

    rows.sort(function (ua, ub) {
        return compareStrings(ua[key] || '', ub[key] || '') * dir;
    });

    return rows;
}

export function refreshDepartmentFilter() {
    const sel = document.getElementById('pvFilterDepartment');
    if (!sel) return;
    const current = sel.value;
    const set = new Set();
    for (let i = 0; i < pv.loadedUsers.length; i++) {
        const d = pv.loadedUsers[i].department;
        if (d && String(d).trim()) set.add(String(d).trim());
    }
    const list = Array.from(set).sort(function (a, b) {
        return compareStrings(a, b);
    });
    sel.replaceChildren();
    const o0 = document.createElement('option');
    o0.value = '';
    o0.textContent = '(alle)';
    sel.appendChild(o0);
    for (let j = 0; j < list.length; j++) {
        const o = document.createElement('option');
        o.value = list[j];
        o.textContent = list[j];
        sel.appendChild(o);
    }
    if (current && set.has(current)) sel.value = current;
}

export function refreshLicenseFilter() {
    const sel = document.getElementById('pvFilterLicense');
    if (!sel) return;
    const api = Lic();
    const current = sel.value;
    sel.replaceChildren();
    const opts =
        api && typeof api.buildLicenseFilterOptions === 'function'
            ? api.buildLicenseFilterOptions(pv.loadedUsers)
            : [{ value: '', label: '(alle Lizenzen)' }];
    for (let i = 0; i < opts.length; i++) {
        const o = document.createElement('option');
        o.value = opts[i].value;
        o.textContent = opts[i].label;
        sel.appendChild(o);
    }
    const values = {};
    for (let j = 0; j < opts.length; j++) values[opts[j].value] = true;
    if (current && values[current]) sel.value = current;
}

export function updateStatsPanel() {
    const total = pv.loadedUsers.length;
    let members = 0;
    let guests = 0;
    let active = 0;
    let adSync = 0;
    let flagged = 0;
    for (let i = 0; i < pv.loadedUsers.length; i++) {
        const u = pv.loadedUsers[i];
        if (String(u.userType || '').toLowerCase() === 'guest') guests++;
        else members++;
        if (u.accountEnabled === true) active++;
        if (u.onPremisesSyncEnabled === true) adSync++;
        if (u.adFlagged) flagged++;
    }
    const el = function (id, val) {
        const n = document.getElementById(id);
        if (n) n.textContent = val;
    };
    el('pvStatTotal', total ? String(total) : '–');
    el('pvStatMember', total ? String(members) : '–');
    el('pvStatGuest', total ? String(guests) : '–');
    el('pvStatActive', total ? String(active) : '–');
    const sum = document.getElementById('pvAdSummary');
    if (sum) {
        if (!total) {
            sum.style.display = 'none';
            sum.textContent = '';
        } else {
            sum.style.display = '';
            sum.innerHTML =
                '<i class="bi bi-hdd-network" aria-hidden="true"></i> ' +
                '<button type="button" class="pv-ad-filter-link" data-pv-ad="adSync" title="Filter: mit lokalem AD synchronisiert">' +
                '<strong>' +
                String(adSync) +
                '</strong> mit lokalem AD synchronisiert</button>' +
                ' · ' +
                '<button type="button" class="pv-ad-filter-link" data-pv-ad="flagged" title="Filter: markierte Konten">' +
                '<strong>' +
                String(flagged) +
                '</strong> markiert</button>';
        }
    }
}

export function updateProgressLine() {
    const progress = document.getElementById('pvProgress');
    if (!progress) return;
    if (!pv.loadedUsers.length) {
        progress.textContent = '';
        return;
    }
    const visible = host.getVisibleRows();
    const base = 'Geladen: ' + pv.loadedUsers.length + ' Person(en).';
    if (visible.length !== pv.loadedUsers.length) {
        progress.textContent = base + ' Angezeigt: ' + visible.length + ' Treffer.';
    } else {
        progress.textContent = base;
    }
}

export function renderUserTree() {
    const tree = document.getElementById('pvTree');
    if (!tree) return;
    tree.replaceChildren();
    const rows = host.getVisibleRows();

    if (!rows.length) {
        const li = document.createElement('li');
        const p = document.createElement('p');
        p.className = 'muted';
        p.style.margin = '0';
        p.style.padding = '14px 12px';
        p.textContent = pv.loadedUsers.length ? 'Keine Treffer für die Filter.' : 'Noch keine Daten – „Personen einlesen“ wählen.';
        li.appendChild(p);
        tree.appendChild(li);
        host.updateProgressLine();
        return;
    }

    for (let i = 0; i < rows.length; i++) {
        const u = rows[i];
        const li = document.createElement('li');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'pv-tree-row';
        btn.dataset.pvSelectUser = u.id || '';
        btn.setAttribute('aria-current', pv.selectedUserId && u.id === pv.selectedUserId ? 'true' : 'false');

        const isGuest = String(u.userType || '').toLowerCase() === 'guest';
        const iconWrap = document.createElement('span');
        iconWrap.className = 'pv-tree-icon';
        const icon = document.createElement('i');
        icon.className = isGuest ? 'bi bi-person-badge' : 'bi bi-person-fill';
        icon.setAttribute('aria-hidden', 'true');
        iconWrap.appendChild(icon);

        const main = document.createElement('div');
        main.className = 'pv-tree-main';
        const title = document.createElement('div');
        title.className = 'pv-tree-title';
        title.textContent = u.displayName || u.userPrincipalName || u.mail || '(ohne Namen)';
        const sub = document.createElement('div');
        sub.className = 'pv-tree-sub';
        sub.textContent = u.userPrincipalName || u.mail || u.id || '';
        main.appendChild(title);
        main.appendChild(sub);

        const meta = document.createElement('div');
        meta.className = 'pv-tree-meta';
        const pillType = document.createElement('span');
        pillType.className = 'pill' + (isGuest ? '' : ' ok');
        pillType.textContent = userTypeLabel(u.userType);
        meta.appendChild(pillType);
        if (u.accountEnabled === false) {
            const pillOff = document.createElement('span');
            pillOff.className = 'pill err';
            pillOff.textContent = 'Inaktiv';
            meta.appendChild(pillOff);
        }
        if (u.onPremisesSyncEnabled === true) {
            const pillAd = document.createElement('span');
            pillAd.className = 'pill warn';
            pillAd.title =
                'Aus lokalem Active Directory synchronisiert' +
                (u.onPremisesSamAccountName ? ' · SAM: ' + u.onPremisesSamAccountName : '');
            pillAd.textContent = 'AD‑Sync';
            meta.appendChild(pillAd);
        }
        if (u.adFlagged) {
            const pillFlag = document.createElement('span');
            pillFlag.className = 'pill';
            pillFlag.title = u.adFlagNote ? String(u.adFlagNote) : 'Für lokalen Admin markiert';
            pillFlag.textContent = 'Markiert';
            meta.appendChild(pillFlag);
        }
        if (u.department) {
            const pillDep = document.createElement('span');
            pillDep.className = 'pill muted-pill';
            pillDep.textContent = String(u.department).trim();
            meta.appendChild(pillDep);
        }

        btn.appendChild(iconWrap);
        btn.appendChild(main);
        btn.appendChild(meta);
        li.appendChild(btn);
        tree.appendChild(li);
    }
    host.updateProgressLine();
}

