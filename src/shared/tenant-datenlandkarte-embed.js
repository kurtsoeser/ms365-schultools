/**
 * Stammdaten: Datenlandkarte als eingebetteter Tab (lazy mount).
 */
import { mountDatenlandkarte } from '../tools/datenlandkarte/datenlandkarte.js';

const TAB_BTN_ID = 'tabMainDatenlandkarte';
const PANEL_ID = 'panelMainDatenlandkarte';
const MOUNT_ID = 'tenantDatenlandkarteApp';

let mounted = false;

function ensureMounted() {
    if (mounted) return;
    const mount = document.getElementById(MOUNT_ID);
    if (!mount) return;
    try {
        mountDatenlandkarte(mount, { embedded: true });
        mounted = true;
    } catch (err) {
        console.error('Datenlandkarte (Stammdaten): Mount fehlgeschlagen', err);
        mount.innerHTML =
            '<p class="muted" role="alert">Datenlandkarte konnte nicht geladen werden. Bitte Seite neu laden oder <a href="tools/datenlandkarte.html">Vollbild-Ansicht</a> öffnen.</p>';
    }
}

function afterShow() {
    ensureMounted();
    requestAnimationFrame(() => {
        if (typeof window.ms365DatenlandkarteFit === 'function') {
            window.ms365DatenlandkarteFit();
        }
        if (typeof window.ms365DatenlandkarteRefresh === 'function') {
            window.ms365DatenlandkarteRefresh();
        }
    });
}

function isDatenlandkarteHash() {
    const raw = String(location.hash || '')
        .replace(/^#/, '')
        .trim()
        .toLowerCase();
    return raw === 'datenlandkarte' || raw === 'landkarte' || raw === 'daten-landkarte';
}

export function initTenantDatenlandkarteEmbed() {
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

    if (panel.classList.contains('active') || isDatenlandkarteHash()) {
        afterShow();
    }
}
