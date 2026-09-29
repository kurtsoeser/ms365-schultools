/**
 * Tenant-Detail / Archiv-UI für Schulstruktur-Sync (Analyse 02 Phase B).
 * Move-first aus schulstruktur-sync.js.
 */

import { getEl } from '../../shared/utils/dom.js';

export function tenantRowToLiveGroup(group) {
    if (!group) return null;
    const typ = String(group.typ || '');
    const unified = typ === 'Team' || typ === 'Gruppe';
    const hasTeam = typ === 'Team' || group.hasTeamsForArchive === true;
    return {
        id: group.id,
        displayName: group.bezeichnung,
        mail: group.mail,
        mailNickname: group.alias,
        description: group.description,
        visibility: group.visibility,
        expirationDateTime: group.expirationDateTime,
        createdDateTime: group.createdDateTime,
        groupTypes: unified ? ['Unified'] : [],
        resourceProvisioningOptions: hasTeam ? ['Team'] : []
    };
}

export function applyTenantArchiveUi(group) {
    const archWrap = getEl('slgTeamArchiveWrap');
    const archSel = getEl('slgArchiveState');
    const archHint = getEl('slgArchiveHint');
    const archSpo = getEl('slgArchiveSpoReadonly');
    if (!archWrap || !archSel || !archHint || !archSpo) return;
    if (!group) {
        archWrap.style.display = 'none';
        return;
    }
    const typ = String(group.typ || '');
    const unifiedLike = typ === 'Team' || typ === 'Gruppe';
    archWrap.style.display = unifiedLike ? '' : 'none';
    if (!unifiedLike) {
        archSpo.checked = false;
        return;
    }
    const st = group.teamIsArchived;
    let cap = group.hasTeamsForArchive;
    if (cap === undefined && (st === true || st === false)) {
        cap = true;
    }
    const loadingCap = cap === undefined && st === undefined;
    if (loadingCap) {
        archSel.disabled = true;
        archSel.value = 'active';
        archHint.style.display = '';
        archHint.textContent = 'Teams-Anbindung und Archiv-Status werden ermittelt …';
        archSpo.disabled = true;
        archSpo.checked = false;
    } else if (cap === false) {
        archSel.disabled = true;
        archSel.value = 'active';
        archHint.style.display = '';
        archHint.textContent =
            'Kein Microsoft Teams an dieser Microsoft 365-Gruppe (nur Gruppe ohne Team/Kursteam) – Teams-Archivierung ist nicht verfügbar.';
        archSpo.disabled = true;
        archSpo.checked = false;
    } else if (cap === true && (st === true || st === false)) {
        archSel.disabled = false;
        archSel.value = st ? 'archived' : 'active';
        archHint.style.display = 'none';
        archHint.textContent = '';
        archSpo.disabled = archSel.value !== 'archived';
    } else {
        archSel.disabled = true;
        archSel.value = 'active';
        archHint.style.display = '';
        archHint.textContent =
            'Teams ist vorhanden, der Archiv-Status konnte nicht gelesen werden. Bitte „Neu laden“ oder Berechtigungen prüfen.';
        archSpo.disabled = true;
        archSpo.checked = false;
    }
}

/** Wird in bind() gesetzt, damit Details auch nach verzögertem Script-Load eingehängt werden. */
let ensureTenantGroupDetailMounted = function () {
    return !!(window.ms365GroupDetail && document.getElementById('slgLiveName'));
};

export function setEnsureTenantGroupDetailMounted(fn) {
    ensureTenantGroupDetailMounted = typeof fn === 'function' ? fn : ensureTenantGroupDetailMounted;
}

export function getEnsureTenantGroupDetailMounted() {
    return ensureTenantGroupDetailMounted;
}

export function formatAdSyncDate(iso) {
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

export function fillTenantAdSyncPanels(group) {
    const adPanel = getEl('ssTenantAdSyncPanel');
    const cloudPanel = getEl('ssTenantCloudFlagPanel');
    if (!group) {
        if (adPanel) adPanel.style.display = 'none';
        if (cloudPanel) cloudPanel.style.display = 'none';
        return;
    }
    const isAd = !!group.onPremisesSyncEnabled;
    if (adPanel) adPanel.style.display = isAd ? '' : 'none';
    if (cloudPanel) cloudPanel.style.display = isAd ? 'none' : '';
    if (isAd) {
        const sam = getEl('ssAdSam');
        const dom = getEl('ssAdDomain');
        const last = getEl('ssAdLastSync');
        const note = getEl('ssAdFlagNote');
        const flagged = getEl('ssAdFlagged');
        if (sam) sam.value = String(group.onPremisesSamAccountName || '–');
        if (dom) dom.value = String(group.onPremisesDomainName || '–');
        if (last) last.value = formatAdSyncDate(group.onPremisesLastSyncDateTime);
        if (note) note.value = String(group.adFlagNote || '');
        if (flagged) flagged.checked = !!group.adFlagged;
    } else {
        const note = getEl('ssCloudFlagNote');
        const flagged = getEl('ssCloudFlagged');
        if (note) note.value = String(group.adFlagNote || '');
        if (flagged) flagged.checked = !!group.adFlagged;
    }
    const delBtn = getEl('slgBtnDeleteGroup');
    if (delBtn) {
        if (isAd) {
            delBtn.disabled = true;
            delBtn.title = 'AD‑gesyncte Gruppen können nicht aus der Cloud gelöscht werden – bitte lokal im AD löschen.';
        } else {
            delBtn.disabled = false;
            delBtn.title = '';
        }
    }
}

export function showTenantDetail(group) {
    const hint = getEl('ssHint');
    const detail = getEl('ssDetail');
    const tenantDetail = getEl('ssTenantDetail');
    if (hint) hint.style.display = group ? 'none' : '';
    if (detail) detail.style.display = 'none';
    if (tenantDetail) tenantDetail.style.display = group ? '' : 'none';
    if (group) ensureTenantGroupDetailMounted();
    const L = window.ms365SlgLiveDetails;
    if (!group) {
        if (L) {
            L.resetCaches();
            L.fillForm(null);
        }
        applyTenantArchiveUi(null);
        fillTenantAdSyncPanels(null);
        return;
    }
    applyTenantArchiveUi(group);
    fillTenantAdSyncPanels(group);
    if (!L) return;
    const idEl = getEl('slgLiveId');
    const already = idEl && String(idEl.value || '') === String(group.id || '');
    if (!already) {
        L.fillForm(tenantRowToLiveGroup(group));
        L.setMatchedMode(true);
    }
    const art = getEl('slgLiveArt');
    if (art) art.value = String(group.typ || '');
    const entra = getEl('slgBtnOpenEntra');
    if (entra) entra.disabled = !group.id;
}

