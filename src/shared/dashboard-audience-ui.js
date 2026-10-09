/**
 * Rollenbasierte Dashboard-Ansicht (index.html).
 */
import {
    isToolVisibleForView,
    viewLabel,
    DASHBOARD_QUICKSTART_BY_VIEW,
    DASHBOARD_TOOL_LINKS,
    DASHBOARD_TOOL_ICONS,
    DASHBOARD_TOOL_BLURBS,
    toolLabel
} from './dashboard-audience-catalog.js';
import { getEffectiveToolRule, loadDashboardToolAccessConfig } from './dashboard-audience-store.js';
import {
    getDashboardPersonas,
    clearDashboardPersonaCache
} from './dashboard-persona-session.js';
import { applyAuthMenuItAudience } from './dashboard-auth-menu-policy.js';
import {
    resolveLayoutModeFromPersonas,
    isAppsOnlyLayout,
    catalogViewForLayout,
    layoutModeLabel,
    writeItPreviewChoice,
    readItPreviewChoice
} from './dashboard-layout-mode.js';

/** @type {import('./dashboard-layout-mode.js').DashboardLayoutMode} */
let layoutMode = 'full';
/** @type {import('./dashboard-audience-resolve.js').DashboardPersonas|null} */
let personas = null;

function plannerContext() {
    if (!personas) return { schularbeitenRoles: [], freistellungRoles: [] };
    return {
        schularbeitenRoles: personas.schularbeitenRoles || [],
        freistellungRoles: personas.freistellungRoles || []
    };
}

function catalogView() {
    return catalogViewForLayout(layoutMode);
}

/** IT-Leiste / Vorschau-Umschalter nur für echte Schul-IT. */
function itDashboardChrome() {
    if (!personas || !personas.loggedIn) return false;
    return !!personas.isIt;
}

function isToolVisible(toolId) {
    const rule = getEffectiveToolRule(toolId, loadDashboardToolAccessConfig());
    return isToolVisibleForView(toolId, plannerContext(), catalogView(), rule);
}

export function isToolElementVisible(card) {
    if (layoutMode === 'full') return true;
    const id = String(card.getAttribute('data-tool-id') || '').trim();
    if (!id) return true;
    return isToolVisible(id);
}

function appOnlyStackEl() {
    return document.getElementById('dashAppOnlyStack');
}

function syncCatalogCardsHidden() {
    const catalog = document.getElementById('dash-catalog');
    if (!catalog || layoutMode === 'full') return;
    const searching = document.body.classList.contains('dashboard-searching');
    catalog.querySelectorAll('.choice[data-tool-id]').forEach((card) => {
        const audOk = isToolElementVisible(card);
        card.toggleAttribute('data-dash-audience-hidden', !audOk);
        if (!searching) card.hidden = !audOk;
    });
}

function renderQuickstart() {
    let wrap = document.getElementById('dashQuickApps');
    const stack = appOnlyStackEl();
    if (!isAppsOnlyLayout(layoutMode)) {
        if (wrap) wrap.hidden = true;
        if (stack) stack.hidden = true;
        return;
    }

    const view = catalogView();
    const list = DASHBOARD_QUICKSTART_BY_VIEW[view] || [];
    const visible = list.filter((id) => isToolVisible(id));
    if (!visible.length) {
        if (wrap) wrap.hidden = true;
        if (stack) stack.hidden = true;
        return;
    }

    if (!wrap) {
        wrap = document.createElement('section');
        wrap.id = 'dashQuickApps';
        wrap.className = 'dash-quick-apps';
        wrap.setAttribute('aria-label', 'Schnellzugriff');
        if (stack) {
            stack.insertBefore(wrap, stack.firstChild);
        } else {
            const content = document.querySelector('.content');
            if (content) content.insertBefore(wrap, content.firstChild);
        }
    } else if (stack && wrap.parentElement !== stack) {
        stack.insertBefore(wrap, stack.firstChild);
    }
    wrap.hidden = false;
    if (stack) stack.hidden = false;

    const cards = visible
        .map((id) => {
            const href = DASHBOARD_TOOL_LINKS[id] || `tools/${id}.html`;
            const icon = DASHBOARD_TOOL_ICONS[id] || 'bi-box-arrow-up-right';
            const blurb = DASHBOARD_TOOL_BLURBS[id] || '';
            return `<a class="dash-quick-apps__card" href="${href}">
                <div class="dash-quick-apps__card-head">
                  <h3><i class="bi ${icon}" aria-hidden="true"></i>${toolLabel(id)}</h3>
                  <i class="bi bi-chevron-right dash-quick-apps__chev" aria-hidden="true"></i>
                </div>
                ${blurb ? `<p class="dash-quick-apps__card-desc">${blurb}</p>` : ''}
                <span class="dash-quick-apps__card-action">Öffnen <i class="bi bi-box-arrow-up-right" aria-hidden="true"></i></span>
            </a>`;
        })
        .join('');

    wrap.innerHTML = `
        <h2 class="dash-quick-apps__title">Meine Anwendungen</h2>
        <div class="dash-quick-apps__grid">${cards}</div>`;
}

function applyDomFilter() {
    const appOnly = isAppsOnlyLayout(layoutMode);
    const full = layoutMode === 'full';

    document.querySelectorAll('[data-dash-section-audience]').forEach((el) => {
        el.hidden = appOnly;
    });

    ['dashCatalogSection', 'dashboard-tools', 'dashboard-tasks'].forEach((id) => {
        const el = document.getElementById(id);
        if (el) el.hidden = appOnly;
    });

    const footer = document.querySelector('.dashboard-local-footer');
    if (footer) footer.hidden = appOnly;

    const appStack = appOnlyStackEl();
    if (appStack && !appOnly) appStack.hidden = true;

    document.documentElement.setAttribute('data-dash-view', catalogView());
    document.documentElement.toggleAttribute('data-dash-filtering', !full);
    document.documentElement.toggleAttribute('data-dash-app-only', appOnly);

    syncCatalogCardsHidden();
    renderQuickstart();
    removeLegacyDashboardChrome();
    syncDashboardAuthMenuView();

    if (typeof window.ms365DashboardAudienceRefreshSearch === 'function') {
        window.ms365DashboardAudienceRefreshSearch();
    }
}

const DASH_AUTH_LAYOUT_IDS = ['full', 'preview-lehrer', 'preview-schueler'];

function removeLegacyDashboardChrome() {
    document.getElementById('dashPersonaBar')?.remove();
    document.getElementById('dashItPreviewBanner')?.remove();
}

function syncDashboardAuthMenuView() {
    const section = document.getElementById('ms365AuthDashViewSection');
    if (!section) return;

    const show = itDashboardChrome();
    section.hidden = !show;

    const hint = document.getElementById('ms365AuthDashViewHint');
    const inPreview =
        layoutMode === 'preview-lehrer' || layoutMode === 'preview-schueler';

    section.querySelectorAll('[data-dash-layout]').forEach((btn) => {
        const id = btn.getAttribute('data-dash-layout');
        const on = id === layoutMode;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-checked', on ? 'true' : 'false');
    });

    if (hint) {
        if (show && inPreview) {
            hint.hidden = false;
            hint.textContent =
                layoutMode === 'preview-schueler'
                    ? 'Vorschau: Dashboard wie Schülerinnen/Schüler (nur Schularbeiten-Planer und Freistellungen).'
                    : 'Vorschau: Dashboard wie Lehrkräfte (nur Schularbeiten-Planer und Freistellungen).';
        } else {
            hint.hidden = true;
            hint.textContent = '';
        }
    }

    const label = document.getElementById('ms365AuthDashViewLabel');
    if (label && show) {
        label.textContent = 'Ansicht (Rollen-Vorschau)';
    }

    const trigger = document.getElementById('ms365AuthBadge');
    if (trigger) {
        trigger.classList.toggle('ms365-auth-menu__trigger--dash-preview', show && inPreview);
        trigger.setAttribute(
            'title',
            show && inPreview ? layoutModeLabel(layoutMode) : 'Konto'
        );
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

    if (!section.dataset.bound) {
        section.dataset.bound = '1';
        section.querySelectorAll('[data-dash-layout]').forEach((btn) => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const m = btn.getAttribute('data-dash-layout');
                if (!m || !DASH_AUTH_LAYOUT_IDS.includes(m)) return;
                setLayoutMode(m);
            });
        });
    }
}

function setLayoutMode(mode) {
    if (!itDashboardChrome()) {
        applyDomFilter();
        return;
    }
    if (mode === 'full') writeItPreviewChoice('full');
    else if (mode === 'preview-lehrer') writeItPreviewChoice('preview-lehrer');
    else if (mode === 'preview-schueler') writeItPreviewChoice('preview-schueler');
    layoutMode = mode;
    applyDomFilter();
}

function applyPersonas(next, opts) {
    personas = next;
    const keepLayout = !!(opts && opts.keepLayout);

    if (!personas || !personas.loggedIn) {
        layoutMode = 'full';
        applyAuthMenuItAudience(false);
        applyDomFilter();
        return;
    }

    if (!keepLayout || !personas.isIt) {
        layoutMode = resolveLayoutModeFromPersonas(personas);
    } else if (layoutMode !== 'full' && !readItPreviewChoice()) {
        layoutMode = 'full';
    }

    applyAuthMenuItAudience(!!(personas && personas.isIt));
    applyDomFilter();
}

async function refreshPersonas(opts) {
    const next = await getDashboardPersonas(opts);
    applyPersonas(next, opts);
}

export function mountDashboardAudience(opts) {
    if (opts && typeof opts.onSearchRefresh === 'function') {
        window.ms365DashboardAudienceRefreshSearch = opts.onSearchRefresh;
    }

    window.ms365DashboardAudience = {
        isToolElementVisible,
        refresh: () => refreshPersonas({ force: true }),
        getLayoutMode: () => layoutMode,
        getPersonas: () => personas,
        setLayoutMode,
        debug: () => ({ personas, layoutMode, catalogView: catalogView() })
    };

    refreshPersonas({ force: true });

    window.addEventListener('ms365-dashboard-persona-ready', () => {
        refreshPersonas({ force: false, keepLayout: true });
    });

    window.addEventListener('ms365-auth-state-changed', () => {
        clearDashboardPersonaCache();
        refreshPersonas({ force: true, keepLayout: false });
    });

    window.addEventListener('ms365-dashboard-tool-access-changed', () => {
        applyDomFilter();
    });

    window.addEventListener('ms365-auth-widget-ready', () => {
        syncDashboardAuthMenuView();
    });
}
