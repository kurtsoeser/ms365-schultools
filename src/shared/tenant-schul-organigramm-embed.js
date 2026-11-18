/**
 * Stammdaten: Schul-Organigramm als eingebetteter Tab (lazy mount).
 * Live-Refresh läuft im Mount (ms365-school-data-written / ms365-school-year-changed).
 * Anzeigen = nur Panel wechseln – kein blindes Full-Refresh (sonst wirkt die UI träge).
 */
import { mountSchulOrganigramm } from '../tools/schul-organigramm/schul-organigramm-ui.js';

const TAB_BTN_ID = 'tabMainOrganigramm';
const PANEL_ID = 'panelMainOrganigramm';
const MOUNT_ID = 'tenantSchulOrganigrammApp';

let api = null;

/**
 * @param {{ refresh?: boolean }} [opts]
 */
function ensureMounted(opts) {
    const mount = document.getElementById(MOUNT_ID);
    if (!mount) return;
    const forceRefresh = !!(opts && opts.refresh);
    if (api) {
        if (forceRefresh && typeof api.refresh === 'function') {
            api.refresh();
            return;
        }
        if (typeof api.resume === 'function') {
            api.resume();
            return;
        }
        return;
    }
    try {
        api = mountSchulOrganigramm(mount, { embedded: true });
    } catch (err) {
        console.error('Schul-Organigramm (Stammdaten): Mount fehlgeschlagen', err);
        mount.innerHTML =
            '<p class="muted" role="alert">Organigramm konnte nicht geladen werden. Bitte Seite neu laden.</p>';
    }
}

function afterShow() {
    // Einmal mounten; wiederholte Tab-/Header-Wechsel nur soft fortsetzen.
    ensureMounted({ refresh: false });
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
