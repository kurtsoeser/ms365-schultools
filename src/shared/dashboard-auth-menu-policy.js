/**
 * Konto-Menü: Aktionsprotokoll nur für Schul-IT (Global Admin / Betreiber).
 * Tenant-ID und „+ Schuljahr“ gehören in Stammdaten, nicht ins Kontomenü.
 * Portal-Rolle wird für alle angemeldeten Nutzer angezeigt.
 * Bereich Portal/IT/Schüler nur bei klarem IT-Zugang (Schul-IT, Global Admin, Betreiber).
 * Schüler/in, Lehrkraft, Direktion, Sekretariat und Schul-Administration bleiben im Portal.
 * Zusätzlich: kompakter Portal/IT-Chip im App-Header, ebenfalls nur für diesen IT-Zugang.
 */
import { getDashboardPersonas } from './dashboard-persona-session.js';
import { primaryPersonaHint, primaryPersonaLabel } from './app-role-switcher-policy.js';
import {
    layoutModeLabel,
    resolveLayoutModeFromPersonas,
    writeItPreviewChoice
} from './dashboard-layout-mode.js';

const DASH_AUTH_LAYOUT_IDS = ['full', 'preview-lehrer', 'preview-schueler'];
const HEADER_SHELL_LAYOUT_IDS = ['full', 'preview-lehrer'];

/**
 * Umschalter Portal / IT / Schüler-Vorschau.
 * Nur ein klarer IT-Zugang: Schul-IT-Kontakt, Entra-Global-Admin oder Betreiber.
 * Schüler/in, Lehrkraft, Direktion, Sekretariat und Schul-Administration haben keinen Wechsler.
 *
 * @param {import('./dashboard-audience-resolve.js').DashboardPersonas|null|undefined} personas
 * @returns {boolean}
 */
export function canUseItShellSwitch(personas) {
    if (!personas || !personas.loggedIn) return false;
    return !!(
        personas.isIt ||
        personas.designatedSchoolIt ||
        personas.globalAdmin ||
        personas.operator
    );
}

function livePersonasAllowShellSwitch() {
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365DashboardAudience &&
            typeof window.ms365DashboardAudience.getPersonas === 'function'
        ) {
            const fromDash = window.ms365DashboardAudience.getPersonas();
            if (fromDash) return canUseItShellSwitch(fromDash);
        }
    } catch {
        /* ignore */
    }
    try {
        if (typeof window !== 'undefined' && window.__ms365DashboardPersonasLast) {
            return canUseItShellSwitch(window.__ms365DashboardPersonasLast);
        }
    } catch {
        /* ignore */
    }
    return false;
}

export function applyAuthMenuItAudience(showItExtras) {
    const actionLog = document.getElementById('ms365AuthActionLogLink');
    const itTourBtn = document.getElementById('ms365AuthItTourBtn');
    const hide = !showItExtras;
    if (actionLog) actionLog.hidden = hide;
    if (itTourBtn) itTourBtn.hidden = hide;
    // Alte Menü-Markup-Versionen: Tenant-Zeile falls noch vorhanden dauerhaft ausblenden.
    const tenantRow = document.querySelector('.ms365-auth-menu__ctx-row--tenant');
    if (tenantRow) tenantRow.hidden = true;
}

/**
 * @param {import('./dashboard-audience-resolve.js').DashboardPersonas|null|undefined} personas
 */
export function syncAuthMenuPersonaStatus(personas) {
    const section = document.getElementById('ms365AuthPersonaSection');
    const valueEl = document.getElementById('ms365AuthPersonaValue');
    const hintEl = document.getElementById('ms365AuthPersonaHint');
    if (!section || !valueEl) return;

    const loggedIn = !!(personas && personas.loggedIn);
    section.hidden = !loggedIn;
    if (!loggedIn) {
        valueEl.textContent = '–';
        if (hintEl) {
            hintEl.hidden = true;
            hintEl.textContent = '';
        }
        return;
    }

    const label = primaryPersonaLabel(personas) || '–';
    valueEl.textContent = label;
    // Hinweise bleiben aus – Kontomenü soll ohne Scroll kompakt bleiben.
    if (hintEl) {
        hintEl.hidden = true;
        hintEl.textContent = '';
        const tip = primaryPersonaHint(personas);
        if (tip) valueEl.title = tip;
    }
}

/**
 * @param {import('./dashboard-audience-resolve.js').DashboardPersonas|null|undefined} personas
 * @param {import('./dashboard-layout-mode.js').DashboardLayoutMode} [layoutMode]
 */
export function syncAuthMenuDashViewSection(personas, layoutMode) {
    const section = document.getElementById('ms365AuthDashViewSection');
    if (!section) return;

    const show = canUseItShellSwitch(personas);
    section.hidden = !show;

    const mode =
        layoutMode ||
        (personas && personas.loggedIn
            ? resolveLayoutModeFromPersonas(personas)
            : 'preview-lehrer');
    const inPortalPreview = mode === 'preview-lehrer' || mode === 'preview-schueler';

    section.querySelectorAll('[data-dash-layout]').forEach((btn) => {
        const id = btn.getAttribute('data-dash-layout');
        const on = id === mode;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-checked', on ? 'true' : 'false');
    });

    const hint = document.getElementById('ms365AuthDashViewHint');
    if (hint) {
        hint.hidden = true;
        hint.textContent = '';
    }

    const label = document.getElementById('ms365AuthDashViewLabel');
    if (label && show) {
        label.textContent = 'Bereich';
    }

    const trigger = document.getElementById('ms365AuthBadge');
    if (trigger) {
        trigger.classList.toggle('ms365-auth-menu__trigger--dash-preview', show && inPortalPreview);
        trigger.setAttribute('title', show ? layoutModeLabel(mode, personas) : 'Konto');
    }

    const werkLink = document.getElementById('ms365AuthDashWerkzeugeLink');
    if (werkLink && !werkLink.dataset.hrefFixed) {
        werkLink.dataset.hrefFixed = '1';
        try {
            const p = String(window.location.pathname || '');
            werkLink.href = /\/tools\//i.test(p)
                ? '../dashboard-werkzeug-zugriff.html'
                : 'dashboard-werkzeug-zugriff.html';
        } catch {
            /* ignore */
        }
    }

    ensureAuthMenuDashViewBindings();
    syncHeaderShellSwitch(personas, mode);
}

/**
 * Kompakter Portal/IT-Umschalter oben rechts (nur Schul-IT).
 * Schüler-Vorschau bleibt im Kontomenü.
 * @param {HTMLElement|null} [container]
 */
export function ensureHeaderShellSwitch(container) {
    if (typeof document === 'undefined') return null;
    const host =
        container ||
        document.getElementById('adminAppTopActions') ||
        document.getElementById('ms365HeaderRightCluster');
    if (!host) return null;

    let switcher = document.getElementById('ms365HeaderShellSwitch');
    if (!switcher) {
        switcher = document.createElement('div');
        switcher.id = 'ms365HeaderShellSwitch';
        switcher.className = 'ms365-header-shell-switch';
        switcher.setAttribute('role', 'group');
        switcher.setAttribute('aria-label', 'Bereich Portal oder IT');
        switcher.hidden = true;
        switcher.innerHTML =
            '<button type="button" class="ms365-header-shell-switch__btn" data-dash-layout="preview-lehrer" aria-pressed="false" title="Portal" aria-label="Portal">' +
            '<i class="bi bi-grid" aria-hidden="true"></i>' +
            '<span class="ms365-header-shell-switch__label">Portal</span></button>' +
            '<button type="button" class="ms365-header-shell-switch__btn" data-dash-layout="full" aria-pressed="false" title="IT-Backend" aria-label="IT-Backend">' +
            '<i class="bi bi-grid-3x3-gap-fill" aria-hidden="true"></i>' +
            '<span class="ms365-header-shell-switch__label">IT</span></button>';
    }

    const auth = document.getElementById('ms365AuthWidget');
    if (auth && auth.parentElement === host) {
        if (switcher.parentElement !== host || switcher.nextElementSibling !== auth) {
            host.insertBefore(switcher, auth);
        }
    } else if (switcher.parentElement !== host) {
        host.appendChild(switcher);
    }

    if (switcher.dataset.bound !== '1') {
        switcher.dataset.bound = '1';
        switcher.querySelectorAll('[data-dash-layout]').forEach((btn) => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                const m = btn.getAttribute('data-dash-layout');
                if (!m || !HEADER_SHELL_LAYOUT_IDS.includes(m)) return;
                commitItShellLayout(/** @type {'full'|'preview-lehrer'} */ (m));
            });
        });
    }
    return switcher;
}

/**
 * @param {import('./dashboard-audience-resolve.js').DashboardPersonas|null|undefined} personas
 * @param {import('./dashboard-layout-mode.js').DashboardLayoutMode} [layoutMode]
 */
export function syncHeaderShellSwitch(personas, layoutMode) {
    const switcher = ensureHeaderShellSwitch();
    if (!switcher) return;

    const show = canUseItShellSwitch(personas);
    switcher.hidden = !show;
    if (!show) {
        switcher.setAttribute('aria-hidden', 'true');
        return;
    }
    switcher.removeAttribute('aria-hidden');

    const mode =
        layoutMode ||
        (personas && personas.loggedIn
            ? resolveLayoutModeFromPersonas(personas)
            : 'preview-lehrer');
    const portalOn = mode === 'preview-lehrer' || mode === 'preview-schueler';
    const itOn = mode === 'full';

    switcher.querySelectorAll('[data-dash-layout]').forEach((btn) => {
        const id = btn.getAttribute('data-dash-layout');
        const on = id === 'full' ? itOn : portalOn;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
    switcher.setAttribute('title', layoutModeLabel(mode, personas));
}

function refreshToolChromeAfterShellChange() {
    return import('./frontend-planner-chrome-policy.js')
        .then((m) => {
            if (m && typeof m.refreshFrontendPlannerChromeAudience === 'function') {
                return m.refreshFrontendPlannerChromeAudience();
            }
        })
        .catch(() => {});
}

/**
 * Shell-Wechsel aus dem Auth-Menü (Dashboard + Tool-Seiten).
 * @param {'full'|'preview-lehrer'|'preview-schueler'} mode
 */
export function commitItShellLayout(mode) {
    if (!DASH_AUTH_LAYOUT_IDS.includes(mode)) return;
    if (!livePersonasAllowShellSwitch()) return;

    try {
        if (
            window.ms365DashboardAudience &&
            typeof window.ms365DashboardAudience.setLayoutMode === 'function'
        ) {
            // Dashboard schreibt Shell + dispatcht ms365-it-shell-changed
            window.ms365DashboardAudience.setLayoutMode(mode);
            void refreshToolChromeAfterShellChange();
            void refreshAuthMenuAudiencePolicy();
            return;
        }
    } catch {
        /* Tool-Seite ohne Dashboard */
    }

    writeItPreviewChoice(mode);
    try {
        window.dispatchEvent(new CustomEvent('ms365-it-shell-changed', { detail: { mode } }));
    } catch {
        /* ignore */
    }
    void refreshToolChromeAfterShellChange();
    void refreshAuthMenuAudiencePolicy();
}

function ensureAuthMenuDashViewBindings() {
    const section = document.getElementById('ms365AuthDashViewSection');
    if (!section || section.dataset.bound === '1') return;
    section.dataset.bound = '1';
    section.querySelectorAll('[data-dash-layout]').forEach((btn) => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const m = btn.getAttribute('data-dash-layout');
            if (!m || !DASH_AUTH_LAYOUT_IDS.includes(m)) return;
            commitItShellLayout(/** @type {'full'|'preview-lehrer'|'preview-schueler'} */ (m));
        });
    });
}

/**
 * @param {import('./dashboard-audience-resolve.js').DashboardPersonas|null|undefined} personas
 * @param {import('./dashboard-layout-mode.js').DashboardLayoutMode} [layoutMode]
 */
export function syncAuthMenuAudienceUi(personas, layoutMode) {
    applyAuthMenuItAudience(!!(personas && personas.isIt));
    syncAuthMenuPersonaStatus(personas);
    syncAuthMenuDashViewSection(personas, layoutMode);
    syncHeaderShellSwitch(personas, layoutMode);
}

export async function refreshAuthMenuAudiencePolicy() {
    try {
        if (typeof window.ms365AuthIsLoggedIn !== 'function' || !window.ms365AuthIsLoggedIn()) {
            syncAuthMenuAudienceUi(null);
            return;
        }
    } catch {
        syncAuthMenuAudienceUi(null);
        return;
    }

    try {
        const personas = await getDashboardPersonas({ force: false });
        let layoutMode;
        try {
            if (
                window.ms365DashboardAudience &&
                typeof window.ms365DashboardAudience.getLayoutMode === 'function'
            ) {
                layoutMode = window.ms365DashboardAudience.getLayoutMode();
            }
        } catch {
            layoutMode = undefined;
        }
        if (!layoutMode && personas && personas.loggedIn) {
            layoutMode = resolveLayoutModeFromPersonas(personas);
        }
        syncAuthMenuAudienceUi(personas, layoutMode);
    } catch {
        syncAuthMenuAudienceUi(null);
    }
}

export function bootAuthMenuAudiencePolicy() {
    if (typeof window !== 'undefined' && window.__ms365AuthMenuPolicyBoot) return;
    if (typeof window !== 'undefined') window.__ms365AuthMenuPolicyBoot = true;

    refreshAuthMenuAudiencePolicy();
    window.addEventListener('ms365-auth-state-changed', () => {
        refreshAuthMenuAudiencePolicy();
    });
    window.addEventListener('ms365-dashboard-persona-ready', () => {
        refreshAuthMenuAudiencePolicy();
    });
    window.addEventListener('ms365-auth-widget-ready', () => {
        refreshAuthMenuAudiencePolicy();
    });
    window.addEventListener('ms365-it-shell-changed', () => {
        refreshAuthMenuAudiencePolicy();
    });
    // Nach License/Operator-Nachzug: Bereich Portal/IT erneut einblenden
    window.addEventListener('ms365-license-changed', () => {
        refreshAuthMenuAudiencePolicy();
    });
}
