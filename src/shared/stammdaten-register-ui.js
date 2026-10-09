/**
 * Stammdaten: Tab-Hashes und Dashboard-Status (keine Modi mehr).
 */
import { parseRegisterHash, buildRegisterHash, isRegisterModeHash, BTN_TO_TAB_HASH } from './stammdaten-register-mode.js';
import { renderRegisterLayerStatus } from './stammdaten-register-status.js';

function $(id) {
    return document.getElementById(id);
}

function activateTab(tabBtnId) {
    if (!tabBtnId || !window.ms365TenantTabActivate) return;
    window.ms365TenantTabActivate(tabBtnId);
}

function applyHashFromLocation() {
    const parsed = parseRegisterHash(location.hash);
    if (parsed.tabBtnId) activateTab(parsed.tabBtnId);
}

function wireRegisterTabLinks() {
    document.querySelectorAll('[data-ms365-register-tab-link]').forEach(function (link) {
        if (link.dataset.registerTabLinkBound === '1') return;
        link.dataset.registerTabLinkBound = '1';
        link.addEventListener('click', function (ev) {
            const tabBtnId = String(link.getAttribute('data-ms365-register-tab-link') || '').trim();
            if (!tabBtnId) return;
            ev.preventDefault();
            try {
                window.dispatchEvent(
                    new CustomEvent('ms365-register-tab-request', { detail: { tabBtnId: tabBtnId } })
                );
            } catch {
                activateTab(tabBtnId);
            }
        });
    });
}

function wireQuickTabNav() {
    document.querySelectorAll('[data-tenant-quick-tab]').forEach(function (link) {
        if (link.dataset.quickBound === '1') return;
        link.dataset.quickBound = '1';
        link.addEventListener('click', function (ev) {
            ev.preventDefault();
            const raw = String(link.getAttribute('data-tenant-quick-tab') || link.getAttribute('href') || '')
                .replace(/^#/, '')
                .toLowerCase();
            if (!raw) return;
            const parsed = parseRegisterHash('#' + raw);
            if (parsed.tabBtnId) activateTab(parsed.tabBtnId);
            const hash = buildRegisterHash(parsed.tabBtnId);
            if (hash) {
                try {
                    history.replaceState(null, '', '#' + hash);
                } catch {
                    location.hash = hash;
                }
            }
        });
    });
}

function wireTabHashUpdates() {
    const tabs = $('tenantMainTabs');
    if (!tabs) return;
    tabs.querySelectorAll('.tab-btn[role="tab"]').forEach(function (btn) {
        if (btn.dataset.registerHashBound === '1') return;
        btn.dataset.registerHashBound = '1';
        btn.addEventListener('click', function () {
            const id = btn.id;
            if (!id || !BTN_TO_TAB_HASH[id]) return;
            const hash = buildRegisterHash(id);
            if (!hash) return;
            try {
                history.replaceState(null, '', '#' + hash);
            } catch {
                location.hash = hash;
            }
        });
    });
}

function refreshLayerStatus() {
    renderRegisterLayerStatus($('tenantRegisterLayerGrid'), { compact: false });
    renderRegisterLayerStatus($('dashboardRegisterStatusMount'), { compact: true });
}

export function mountDashboardRegisterStatus() {
    const el = $('dashboardRegisterStatusMount');
    if (!el || el.dataset.mounted === '1') return;
    el.dataset.mounted = '1';
    refreshLayerStatus();
    window.addEventListener('ms365-spo-sync-status', refreshLayerStatus);
    window.addEventListener('ms365-tenant-settings-changed', refreshLayerStatus);
}

export function initTenantSchulregister() {
    if (!$('tenantMainTabs')) return;

    document.title = 'MS365-Schul-Tools – Stammdaten (Stammdaten)';
    const hint = $('tenantHeaderModeHint');
    if (hint) hint.hidden = true;

    wireQuickTabNav();
    wireRegisterTabLinks();
    wireTabHashUpdates();
    applyHashFromLocation();

    window.addEventListener('hashchange', applyHashFromLocation);
    window.addEventListener('ms365-register-tab-request', function (ev) {
        const id = ev && ev.detail && ev.detail.tabBtnId;
        if (!id) return;
        activateTab(id);
        const hash = buildRegisterHash(id);
        if (hash) {
            try {
                history.replaceState(null, '', '#' + hash);
            } catch {
                location.hash = hash;
            }
        }
    });
    window.addEventListener('ms365-register-mode-request', function (ev) {
        const m = ev && ev.detail && ev.detail.mode;
        if (m === 'sync') activateTab('tabMainStammdaten');
        else if (m === 'import') activateTab('tabMainSchueler');
        else activateTab('tabMainStammdaten');
    });

    const refreshBtn = $('tenantStatusRefresh');
    if (refreshBtn && refreshBtn.dataset.registerLayerBound !== '1') {
        refreshBtn.dataset.registerLayerBound = '1';
        refreshBtn.addEventListener('click', function () {
            refreshLayerStatus();
            if (typeof window.ms365RenderTenantStatusOverview === 'function') {
                window.ms365RenderTenantStatusOverview();
            }
        });
    }
}

/** Für tenant.html Tab-Skript: Tab-Hash nicht doppelt auswerten. */
export function tenantResolveHashTabId(hash) {
    const raw = String(hash || '')
        .replace(/^#/, '')
        .trim()
        .toLowerCase();
    if (isRegisterModeHash(raw)) return '';
    const parsed = parseRegisterHash('#' + raw);
    return parsed.tabBtnId || '';
}
