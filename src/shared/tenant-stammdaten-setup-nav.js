/**
 * Stammdaten-Tab: linke Unternavigation (Schulprofil, E-Mail, SharePoint-Ziele).
 */
const PANE_IDS = ['profile', 'email', 'spo-lists', 'spo-planer', 'spo-it', 'spo-hub'];
const STORAGE_KEY = 'ms365-stammdaten-setup-pane-v1';

function $(id) {
    return document.getElementById(id);
}

function root() {
    return $('tenantStammdatenSetup');
}

function navButtons(setup) {
    return Array.from((setup || root())?.querySelectorAll('[data-stammdaten-pane]') || []);
}

function panels(setup) {
    return Array.from((setup || root())?.querySelectorAll('[data-stammdaten-pane-panel]') || []);
}

function isPaneAvailable(paneId, setup) {
    const btn = (setup || root())?.querySelector('[data-stammdaten-pane="' + paneId + '"]');
    if (!btn) return false;
    if (btn.hidden || btn.getAttribute('aria-hidden') === 'true') return false;
    return true;
}

function readStoredPane() {
    try {
        const v = String(localStorage.getItem(STORAGE_KEY) || '').trim();
        return PANE_IDS.includes(v) ? v : '';
    } catch {
        return '';
    }
}

function storePane(paneId) {
    try {
        localStorage.setItem(STORAGE_KEY, paneId);
    } catch {
        /* ignore */
    }
}

export function activateStammdatenSetupPane(paneId, opts) {
    const setup = root();
    if (!setup) return;
    let next = String(paneId || '').trim();
    if (!PANE_IDS.includes(next) || !isPaneAvailable(next, setup)) {
        next = PANE_IDS.find(function (id) {
            return isPaneAvailable(id, setup);
        });
        if (!next) return;
    }

    navButtons(setup).forEach(function (btn) {
        const id = String(btn.getAttribute('data-stammdaten-pane') || '');
        const on = id === next;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-current', on ? 'page' : 'false');
    });

    panels(setup).forEach(function (panel) {
        const id = String(panel.getAttribute('data-stammdaten-pane-panel') || '');
        const on = id === next;
        panel.hidden = !on;
        if (on) panel.removeAttribute('aria-hidden');
        else panel.setAttribute('aria-hidden', 'true');
    });

    setup.setAttribute('data-active-pane', next);
    if (!opts || opts.persist !== false) storePane(next);

    if (opts && opts.scroll) {
        try {
            setup.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        } catch {
            /* ignore */
        }
    }
}

export function refreshStammdatenSetupNavAvailability() {
    const setup = root();
    if (!setup) return;
    const active = String(setup.getAttribute('data-active-pane') || '').trim();
    const stored = readStoredPane();
    if (active && isPaneAvailable(active, setup)) return;
    if (stored && isPaneAvailable(stored, setup)) {
        activateStammdatenSetupPane(stored, { persist: false });
        return;
    }
    activateStammdatenSetupPane('profile', { persist: false });
}

export function mountTenantStammdatenSetupNav() {
    const setup = root();
    if (!setup || setup.dataset.setupNavBound === '1') {
        refreshStammdatenSetupNavAvailability();
        return;
    }
    setup.dataset.setupNavBound = '1';

    navButtons(setup).forEach(function (btn) {
        btn.addEventListener('click', function () {
            const id = String(btn.getAttribute('data-stammdaten-pane') || '');
            activateStammdatenSetupPane(id);
        });
    });

    const initial = readStoredPane() || 'profile';
    activateStammdatenSetupPane(initial, { persist: false });
}
