/**
 * Stammdaten: Schul-Organigramm als eingebetteter Tab (lazy mount).
 */
import { mountSchulOrganigramm } from '../tools/schul-organigramm/schul-organigramm-ui.js';

const TAB_BTN_ID = 'tabMainOrganigramm';
const PANEL_ID = 'panelMainOrganigramm';
const MOUNT_ID = 'tenantSchulOrganigrammApp';

let api = null;

function ensureMounted() {
    const mount = document.getElementById(MOUNT_ID);
    if (!mount) return;
    if (api && typeof api.refresh === 'function') {
        api.refresh();
        return;
    }
    try {
        api = mountSchulOrganigramm(mount, { embedded: true });
    } catch (err) {
        console.error('Schul-Organigramm (Stammdaten): Mount fehlgeschlagen', err);
        mount.innerHTML =
            '<p class="muted" role="alert">Organigramm konnte nicht geladen werden. Bitte Seite neu laden oder <a href="tools/schul-organigramm.html">Vollbild-Ansicht</a> öffnen.</p>';
    }
}

function afterShow() {
    ensureMounted();
}

function isOrganigrammHash() {
    const raw = String(location.hash || '')
        .replace(/^#/, '')
        .trim()
        .toLowerCase();
    return (
        raw === 'organigramm' ||
        raw === 'schul-organigramm' ||
        raw === 'schulorganigramm' ||
        raw === 'org'
    );
}

export function initTenantSchulOrganigrammEmbed() {
    const panel = document.getElementById(PANEL_ID);
    const tabBtn = document.getElementById(TAB_BTN_ID);
    if (!panel || !tabBtn) return;

    tabBtn.addEventListener('click', afterShow);

    const prevActivate = window.ms365TenantTabActivate;
    if (typeof prevActivate === 'function') {
        window.ms365TenantTabActivate = function (activeBtnId) {
            prevActivate(activeBtnId);
            if (activeBtnId === TAB_BTN_ID) afterShow();
        };
    }

    window.addEventListener('ms365-register-tab-request', function (ev) {
        const id = ev && ev.detail && ev.detail.tabBtnId;
        if (id === TAB_BTN_ID) afterShow();
    });

    if (panel.classList.contains('active') || isOrganigrammHash()) {
        afterShow();
    }
}
