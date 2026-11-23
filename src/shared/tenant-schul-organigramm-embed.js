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
            '<div class="ts-org-empty" role="status">' +
            '<p><strong>Organigramm noch nicht bereit.</strong></p>' +
            '<p class="muted">Oft fehlen noch Rollen oder Klassen in den Stammdaten – das ist kein Seitenfehler.</p>' +
            '<p class="ts-org-empty__actions">' +
            '<a class="btn fb-btn fb-btn--ghost" href="#verwaltung" data-ms365-register-tab-link="tabMainVerwaltung">' +
            '<i class="bi bi-building" aria-hidden="true"></i> Zur Verwaltung</a>' +
            '<a class="btn fb-btn fb-btn--ghost" href="#klassen" data-ms365-register-tab-link="tabMainKlassen">' +
            '<i class="bi bi-people" aria-hidden="true"></i> Zu den Klassen</a>' +
            '</p></div>';
        bindOrgEmptyTabLinks(mount);
    }
}

/**
 * @param {HTMLElement} mount
 */
function bindOrgEmptyTabLinks(mount) {
    if (!mount) return;
    mount.querySelectorAll('[data-ms365-register-tab-link]').forEach(function (a) {
        a.addEventListener('click', function (ev) {
            const id = a.getAttribute('data-ms365-register-tab-link');
            if (!id) return;
            ev.preventDefault();
            if (typeof window.ms365TenantTabActivate === 'function') {
                window.ms365TenantTabActivate(id);
                return;
            }
            const btn = document.getElementById(id);
            if (btn) btn.click();
        });
    });
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
