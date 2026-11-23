/**
 * Stammdaten-Tab: linke Unternavigation (Schulprofil, E-Mail, SharePoint-Ziele).
 */
const PANE_IDS = ['profile', 'email', 'spo-lists', 'spo-planer', 'spo-it', 'spo-hub'];
const STORAGE_KEY = 'ms365-stammdaten-setup-pane-v1';

const STATUS_LABEL = {
    ok: 'OK',
    warn: 'Teilweise',
    empty: 'Offen'
};

/** Hash in der URL für Setup-Panes (Register-Tab bleibt Stammdaten). */
function hashForPane(paneId) {
    const id = String(paneId || '').trim();
    if (id === 'profile') return 'stammdaten';
    if (PANE_IDS.includes(id)) return id;
    return '';
}

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

function fieldValue(id) {
    const el = $(id);
    return el ? String(el.value || '').trim() : '';
}

function hasUrl(id) {
    return /^https?:\/\//i.test(fieldValue(id));
}

/**
 * Einrichtungsstatus je Sub-Tab (für Nav-Badges).
 * @returns {Record<string, 'ok'|'warn'|'empty'>}
 */
export function computeStammdatenSetupPaneStatus() {
    const name = fieldValue('schoolName');
    const domain = fieldValue('schoolEmailDomain');
    let profile = 'empty';
    if (name && domain) profile = 'ok';
    else if (name || domain) profile = 'warn';

    const email = fieldValue('tenantSchoolEmailPattern') || 'vorname.nachname';
    const emailStatus = email ? 'ok' : 'empty';

    const lists = hasUrl('tenantIntranetSiteUrl') ? 'ok' : 'empty';
    const planer = hasUrl('tenantPlanerDefaultSiteUrl')
        ? 'ok'
        : lists === 'ok'
          ? 'warn'
          : 'empty';
    const it = hasUrl('tenantItLibrarySiteUrl') ? 'ok' : 'empty';
    const hub = hasUrl('tenantSchoolIntranetSiteUrl')
        ? 'ok'
        : lists === 'ok'
          ? 'warn'
          : 'empty';

    return {
        profile,
        email: emailStatus,
        'spo-lists': lists,
        'spo-planer': planer,
        'spo-it': it,
        'spo-hub': hub
    };
}

export function refreshStammdatenSetupNavStatus() {
    const setup = root();
    if (!setup) return;
    const statusMap = computeStammdatenSetupPaneStatus();
    PANE_IDS.forEach(function (paneId) {
        const badge = setup.querySelector('[data-setup-status-badge="' + paneId + '"]');
        const btn = setup.querySelector('[data-stammdaten-pane="' + paneId + '"]');
        if (!badge || !btn) return;
        const tone = statusMap[paneId] || 'empty';
        badge.hidden = false;
        badge.dataset.tone = tone;
        badge.textContent = STATUS_LABEL[tone] || tone;
        btn.setAttribute('data-setup-status', tone);
        badge.setAttribute(
            'title',
            tone === 'ok'
                ? 'Einrichtung vorhanden'
                : tone === 'warn'
                  ? 'Teilweise eingerichtet'
                  : 'Noch offen'
        );
    });
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
        if (on) btn.setAttribute('aria-current', 'page');
        else btn.removeAttribute('aria-current');
    });

    panels(setup).forEach(function (panel) {
        const id = String(panel.getAttribute('data-stammdaten-pane-panel') || '');
        const on = id === next;
        panel.hidden = !on;
        panel.classList.toggle('is-pane-active', on);
        if (on) {
            panel.removeAttribute('aria-hidden');
            panel.removeAttribute('hidden');
        } else {
            panel.setAttribute('aria-hidden', 'true');
            panel.setAttribute('hidden', '');
        }
    });

    setup.setAttribute('data-active-pane', next);
    setup.setAttribute('data-setup-ready', '1');
    if (!opts || opts.persist !== false) storePane(next);
    refreshStammdatenSetupNavStatus();

    if (opts && opts.updateHash) {
        const hash = hashForPane(next);
        if (hash) {
            try {
                const cur = String(location.hash || '').replace(/^#/, '');
                if (cur !== hash) history.replaceState(null, '', '#' + hash);
            } catch {
                /* ignore */
            }
        }
    }

    if (opts && opts.scroll) {
        try {
            const activePanel = setup.querySelector(
                '[data-stammdaten-pane-panel="' + next + '"].is-pane-active'
            );
            (activePanel || setup).scrollIntoView({ block: 'nearest', behavior: 'smooth' });
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
    if (active && isPaneAvailable(active, setup)) {
        refreshStammdatenSetupNavStatus();
        return;
    }
    if (stored && isPaneAvailable(stored, setup)) {
        activateStammdatenSetupPane(stored, { persist: false });
        return;
    }
    activateStammdatenSetupPane('profile', { persist: false });
}

/** Hash `#spo-planer` / `#profile` … → Setup-Pane (Playbook Deep-Links). */
export function paneIdFromLocationHash(hash) {
    const raw = String(hash || (typeof location !== 'undefined' ? location.hash : '') || '')
        .replace(/^#/, '')
        .trim()
        .toLowerCase();
    if (PANE_IDS.includes(raw)) return raw;
    if (raw === 'planer-apps' || raw === 'planer' || raw === 'tool-site') return 'spo-planer';
    if (raw === 'schulprofil' || raw === 'profil') return 'profile';
    // #stammdaten = Register-Tab, nicht zwingend Pane „profile“ zurücksetzen
    return '';
}

export function mountTenantStammdatenSetupNav() {
    const setup = root();
    if (!setup) return;

    if (setup.dataset.setupNavBound === '1') {
        refreshStammdatenSetupNavAvailability();
        const fromHash = paneIdFromLocationHash();
        if (fromHash) activateStammdatenSetupPane(fromHash, { persist: false, scroll: true });
        return;
    }
    setup.dataset.setupNavBound = '1';

    // Delegation: überlebt DOM-Änderungen an den Buttons, fängt Klicks auf Badge/Icon
    setup.addEventListener('click', function (ev) {
        const t = ev.target;
        if (!t || !t.closest) return;
        const btn = t.closest('[data-stammdaten-pane]');
        if (!btn || !setup.contains(btn)) return;
        if (btn.hidden || btn.getAttribute('aria-hidden') === 'true') return;
        ev.preventDefault();
        ev.stopPropagation();
        const id = String(btn.getAttribute('data-stammdaten-pane') || '');
        activateStammdatenSetupPane(id, { persist: true, updateHash: true });
    });

    const statusInputs = [
        'schoolName',
        'schoolEmailDomain',
        'tenantSchoolEmailPattern',
        'tenantIntranetSiteUrl',
        'tenantPlanerDefaultSiteUrl',
        'tenantItLibrarySiteUrl',
        'tenantSchoolIntranetSiteUrl'
    ];
    statusInputs.forEach(function (id) {
        const el = $(id);
        if (!el || el.dataset.setupStatusBound === '1') return;
        el.dataset.setupStatusBound = '1';
        el.addEventListener('input', refreshStammdatenSetupNavStatus);
        el.addEventListener('change', refreshStammdatenSetupNavStatus);
    });

    window.addEventListener('ms365-tenant-settings-changed', function () {
        refreshStammdatenSetupNavStatus();
    });

    window.addEventListener('hashchange', function () {
        const fromHash = paneIdFromLocationHash();
        if (fromHash) activateStammdatenSetupPane(fromHash, { persist: false, scroll: true });
    });

    const fromHash = paneIdFromLocationHash();
    const initial = fromHash || readStoredPane() || 'profile';
    activateStammdatenSetupPane(initial, { persist: !fromHash, scroll: !!fromHash });
}
