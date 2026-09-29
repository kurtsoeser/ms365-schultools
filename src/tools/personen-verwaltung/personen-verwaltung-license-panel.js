/**
 * Lizenz-Panel (Analyse 02 Phase B).
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


/** @type {any} */
let host = null;

export function init(a) {
    host = a;
}

export function renderLicenseTab() {
    const u = host.getSelectedUser();
    const hint = document.getElementById('pvLicHint');
    const usage = document.getElementById('pvLicUsageLocation');
    const tbody = document.getElementById('pvLicAssignedBody');
    const sel = document.getElementById('pvLicAssignSelect');
    const status = document.getElementById('pvLicStatus');
    if (!u) return;
    if (usage) usage.value = String(u.usageLocation || '').toUpperCase();
    if (hint) {
        hint.textContent = pv.subscribedSkusOk
            ? 'Zuweisen und Entziehen über Microsoft Graph (assignLicense). Nutzungsort ist ein zweistelliger Ländercode (Österreich: AT).'
            : 'Mandanten-SKUs konnten nicht gelesen werden (Organization.Read.All). Zuweisen über den Education-Katalog ist möglich; Graph lehnt unbekannte SKUs ab.';
    }
    if (status) status.textContent = '';

    const api = Lic();
    const lookup = skuLookupFromSubscribed();
    const sum = api && typeof api.summarizeUserLicenses === 'function' ? api.summarizeUserLicenses(u, lookup) : null;
    const licenses = sum && Array.isArray(sum.licenses) ? sum.licenses : [];

    if (tbody) {
        tbody.replaceChildren();
        if (!licenses.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 3;
            td.style.color = '#6c757d';
            td.textContent = 'Keine Lizenz zugewiesen.';
            tr.appendChild(td);
            tbody.appendChild(tr);
        } else {
            licenses.forEach(function (lic) {
                const tr = document.createElement('tr');
                const tdName = document.createElement('td');
                tdName.textContent = lic.name || lic.shortLabel || lic.skuId;
                const tdSku = document.createElement('td');
                const code = document.createElement('code');
                code.textContent = String(lic.skuPartNumber || lic.skuId || '').slice(0, 42);
                tdSku.appendChild(code);
                const tdAct = document.createElement('td');
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn small-btn';
                btn.setAttribute('data-pv-lic-remove', lic.skuId);
                btn.textContent = 'Entziehen';
                tdAct.appendChild(btn);
                tr.appendChild(tdName);
                tr.appendChild(tdSku);
                tr.appendChild(tdAct);
                tbody.appendChild(tr);
            });
        }
    }

    if (sel) {
        const assigned = assignedSkuIdsOfUser(u);
        const opts =
            api && typeof host.buildAssignableSkuOptions === 'function'
                ? host.buildAssignableSkuOptions(pv.subscribedSkus, assigned, {
                      fallbackCatalog: !pv.subscribedSkusOk
                  })
                : [];
        sel.replaceChildren();
        const o0 = document.createElement('option');
        o0.value = '';
        o0.textContent = opts.length ? '(Lizenz wählen)' : '(keine freie Lizenz)';
        sel.appendChild(o0);
        opts.forEach(function (o) {
            const opt = document.createElement('option');
            opt.value = o.skuId;
            const rest = o.remaining == null ? '' : ' · ' + o.remaining + ' frei';
            opt.textContent = (o.name || o.shortLabel) + rest;
            opt.disabled = !!o.disabled;
            sel.appendChild(opt);
        });
    }
}

export async function loadSubscribedSkus(token) {
    try {
        const data = await graphJson(
            'GET',
            '/pv.subscribedSkus?$select=' +
                encodeURIComponent('skuId,skuPartNumber,prepaidUnits,consumedUnits,capabilityStatus'),
            token
        );
        pv.subscribedSkus = Array.isArray(data.value) ? data.value : [];
        pv.subscribedSkusOk = true;
        host.appendLog('Mandanten-Lizenzen: ' + pv.subscribedSkus.length + ' SKU(s).', 'ok');
    } catch (e) {
        pv.subscribedSkus = [];
        pv.subscribedSkusOk = false;
        host.appendLog('Mandanten-Lizenzen nicht lesbar: ' + graphErrorFriendly(e), 'warn');
    }
}

export async function ensureUsageLocation(token, u, locationHint) {
    const cur = String((u && u.usageLocation) || '').trim().toUpperCase();
    if (/^[A-Z]{2}$/.test(cur)) return cur;
    let next = String(locationHint || '').trim().toUpperCase();
    if (!/^[A-Z]{2}$/.test(next)) {
        const asked = await dlgPrompt(
            'Für die Lizenzzuweisung braucht das Konto einen Nutzungsort (Ländercode, z. B. AT).',
            'AT',
            { title: 'Nutzungsort', inputLabel: 'Ländercode', okText: 'Setzen' }
        );
        if (asked == null) return '';
        next = String(asked).trim().toUpperCase();
    }
    if (!/^[A-Z]{2}$/.test(next)) {
        host.toast('Ungültiger Ländercode (zwei Buchstaben, z. B. AT).');
        return '';
    }
    await graphJson('PATCH', '/users/' + encodeURIComponent(u.id), token, { usageLocation: next });
    u.usageLocation = next;
    host.appendLog('Nutzungsort gesetzt: ' + next, 'ok');
    return next;
}

export async function assignSelectedLicense() {
    const u = host.getSelectedUser();
    const sel = document.getElementById('pvLicAssignSelect');
    const skuId = sel && sel.value ? String(sel.value).toLowerCase() : '';
    if (!u || !skuId) {
        host.toast('Bitte eine Lizenz wählen.');
        return;
    }
    if (pv.licenseBusy) return;
    const optLabel = sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].textContent : skuId;
    const isGuest = String(u.userType || '').toLowerCase() === 'guest';
    if (
        isGuest &&
        !(await dlgConfirm(
            'Gäste erhalten selten Education-Lizenzen. Trotzdem zuweisen?\n\n' + optLabel,
            { title: 'Lizenz zuweisen', okText: 'Zuweisen' }
        ))
    ) {
        return;
    }
    if (
        !isGuest &&
        !(await dlgConfirm('Lizenz zuweisen?\n\n' + optLabel + '\n\n' + (u.displayName || u.userPrincipalName || ''), {
            title: 'Lizenz zuweisen',
            okText: 'Zuweisen'
        }))
    ) {
        return;
    }
    const usageEl = document.getElementById('pvLicUsageLocation');
    const btn = document.getElementById('pvLicAssignBtn');
    pv.licenseBusy = true;
    if (btn) btn.disabled = true;
    try {
        const token = await getGraphToken();
        const loc = await ensureUsageLocation(token, u, usageEl ? usageEl.value : '');
        if (!loc) return;
        await graphJson('POST', '/users/' + encodeURIComponent(u.id) + '/assignLicense', token, {
            addLicenses: [{ skuId: skuId, disabledPlans: [] }],
            removeLicenses: []
        });
        const fresh = await host.refreshUserFromGraph(token, u.id);
        host.mergeUserIntoList(fresh);
        host.appendLog('Lizenz zugewiesen: ' + optLabel, 'ok');
        host.toast('Lizenz zugewiesen.');
        host.renderProfileTab(fresh, pv.profileEditMode);
        host.renderUserTree();
        await loadSubscribedSkus(token);
        host.renderLicenseTab();
    } catch (e) {
        const msg = graphErrorFriendly(e);
        host.appendLog('Lizenz zuweisen: ' + msg, 'err');
        host.toast(msg);
    } finally {
        pv.licenseBusy = false;
        if (btn) btn.disabled = false;
    }
}

export async function removeLicense(skuIdRaw) {
    const u = host.getSelectedUser();
    const skuId = String(skuIdRaw || '').toLowerCase();
    if (!u || !skuId || pv.licenseBusy) return;
    const api = Lic();
    const lookup = skuLookupFromSubscribed();
    const info = api && typeof api.resolveSku === 'function' ? api.resolveSku(skuId) : { name: skuId };
    const lookupPart = lookup.get(skuId);
    const label =
        api && lookupPart
            ? api.resolveSku(skuId, lookupPart.skuPartNumber).name
            : info.name || skuId;
    if (
        !(await dlgConfirm('Lizenz entziehen?\n\n' + label + '\n\n' + (u.displayName || u.userPrincipalName || ''), {
            title: 'Lizenz entziehen',
            okText: 'Entziehen',
            danger: true
        }))
    ) {
        return;
    }
    pv.licenseBusy = true;
    try {
        const token = await getGraphToken();
        await graphJson('POST', '/users/' + encodeURIComponent(u.id) + '/assignLicense', token, {
            addLicenses: [],
            removeLicenses: [skuId]
        });
        const fresh = await host.refreshUserFromGraph(token, u.id);
        host.mergeUserIntoList(fresh);
        host.appendLog('Lizenz entzogen: ' + label, 'ok');
        host.toast('Lizenz entzogen.');
        host.renderProfileTab(fresh, pv.profileEditMode);
        host.renderUserTree();
        await loadSubscribedSkus(token);
        host.renderLicenseTab();
    } catch (e) {
        const msg = graphErrorFriendly(e);
        host.appendLog('Lizenz entziehen: ' + msg, 'err');
        host.toast(msg);
    } finally {
        pv.licenseBusy = false;
    }
}

export async function saveUsageLocation() {
    const u = host.getSelectedUser();
    const inp = document.getElementById('pvLicUsageLocation');
    if (!u || !inp) return;
    const next = String(inp.value || '').trim().toUpperCase();
    if (!/^[A-Z]{2}$/.test(next)) {
        host.toast('Ländercode: zwei Buchstaben, z. B. AT.');
        return;
    }
    try {
        const token = await getGraphToken();
        await graphJson('PATCH', '/users/' + encodeURIComponent(u.id), token, { usageLocation: next });
        const fresh = await host.refreshUserFromGraph(token, u.id);
        host.mergeUserIntoList(fresh);
        host.appendLog('Nutzungsort gespeichert: ' + next, 'ok');
        host.toast('Nutzungsort gespeichert.');
        host.renderLicenseTab();
    } catch (e) {
        const msg = graphErrorFriendly(e);
        host.appendLog('Nutzungsort: ' + msg, 'err');
        host.toast(msg);
    }
}

