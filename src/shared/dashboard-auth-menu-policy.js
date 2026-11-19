/**
 * Konto-Menü: Tenant-ID und Aktionsprotokoll nur für Schul-IT (Global Admin / Betreiber).
 * Portal-Rolle wird für alle angemeldeten Nutzer angezeigt (ohne Umschalter für Nicht-IT).
 * Bereich Portal/IT/Schüler auch auf Tool-Seiten (nicht nur Dashboard).
 */
import { getDashboardPersonas } from './dashboard-persona-session.js';
import { primaryPersonaHint, primaryPersonaLabel } from './app-role-switcher-policy.js';
import {
    layoutModeLabel,
    resolveLayoutModeFromPersonas,
    writeItPreviewChoice
} from './dashboard-layout-mode.js';

const DASH_AUTH_LAYOUT_IDS = ['full', 'preview-lehrer', 'preview-schueler'];

export function applyAuthMenuItAudience(showItExtras) {
    const tenantRow = document.querySelector('.ms365-auth-menu__ctx-row--tenant');
    const actionLog = document.getElementById('ms365AuthActionLogLink');
    const hide = !showItExtras;
    if (tenantRow) tenantRow.hidden = hide;
    if (actionLog) actionLog.hidden = hide;
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
    const hint = primaryPersonaHint(personas);
    if (hintEl) {
        if (hint) {
            hintEl.hidden = false;
            hintEl.textContent = hint;
        } else {
            hintEl.hidden = true;
            hintEl.textContent = '';
        }
    }
}

/**
 * @param {import('./dashboard-audience-resolve.js').DashboardPersonas|null|undefined} personas
 * @param {import('./dashboard-layout-mode.js').DashboardLayoutMode} [layoutMode]
 */
export function syncAuthMenuDashViewSection(personas, layoutMode) {
    const section = document.getElementById('ms365AuthDashViewSection');
    if (!section) return;

    const show = !!(personas && personas.loggedIn && personas.isIt);
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
        if (show && mode === 'preview-schueler') {
            hint.hidden = false;
            hint.textContent = 'Vorschau: so sieht das Portal für Schüler/innen aus.';
        } else if (show && mode === 'preview-lehrer') {
            hint.hidden = false;
            hint.textContent = 'Portal – Alltags-Apps. IT-Backend über „IT“.';
        } else if (show && mode === 'full') {
            hint.hidden = false;
            hint.textContent = 'IT-Backend – Setup, Listen, voller Katalog.';
        } else {
            hint.hidden = true;
            hint.textContent = '';
        }
    }

    const label = document.getElementById('ms365AuthDashViewLabel');
    if (label && show) {
        label.textContent = 'Bereich';
    }

    const trigger = document.getElementById('ms365AuthBadge');
    if (trigger) {
        trigger.classList.toggle('ms365-auth-menu__trigger--dash-preview', show && inPortalPreview);
        trigger.setAttribute('title', show ? layoutModeLabel(mode) : 'Konto');
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
