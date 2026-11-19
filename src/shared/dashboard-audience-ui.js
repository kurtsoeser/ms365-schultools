/**
 * Rollenbasierte Dashboard-Ansicht (index.html).
 * Schulportal zuerst; IT-Bereich nur bewusst (layoutMode === 'full').
 */
import {
    isToolVisibleForView,
    DASHBOARD_QUICKSTART_BY_VIEW,
    DASHBOARD_PORTAL_SECTIONS_BY_VIEW,
    DASHBOARD_TOOL_LINKS,
    DASHBOARD_TOOL_ICONS,
    DASHBOARD_TOOL_BLURBS,
    toolLabel,
    viewLabel
} from './dashboard-audience-catalog.js';
import { getEffectiveToolRule, loadDashboardToolAccessConfig } from './dashboard-audience-store.js';
import {
    applyPortalSectionOrder,
    loadPortalDefaultsConfig,
    resolvePortalHomeToolIds
} from './dashboard-portal-defaults-store.js';
import { loadDashboardFavorites } from './dashboard-personal-tools.js';
import {
    getDashboardPersonas,
    clearDashboardPersonaCache
} from './dashboard-persona-session.js';
import {
    applyAuthMenuItAudience,
    syncAuthMenuAudienceUi
} from './dashboard-auth-menu-policy.js';
import {
    resolveLayoutModeFromPersonas,
    isAppsOnlyLayout,
    catalogViewForLayout,
    writeItPreviewChoice
} from './dashboard-layout-mode.js';
import {
    loadPortalQuickLinks,
    savePortalQuickLinks,
    activePortalQuickLinks,
    PORTAL_QUICK_LINK_DEFAULTS,
    PORTAL_QUICK_LINK_SCHOOL_IDS
} from './dashboard-portal-links.js';
import { greetingForHour, extractFirstName } from './dashboard-greeting-band.js';
import { buildDashboardSchoolCard } from './school-profile-logic.js';

/** @type {import('./dashboard-layout-mode.js').DashboardLayoutMode} */
let layoutMode = 'preview-lehrer';
/** @type {import('./dashboard-audience-resolve.js').DashboardPersonas|null} */
let personas = null;

const PORTAL_NAV_KEY = 'ms365-dashboard-portal-nav-v1';
/** @type {string} */
let portalNavId = 'home';

try {
    const saved = String(sessionStorage.getItem(PORTAL_NAV_KEY) || '').trim();
    if (saved) portalNavId = saved;
} catch {
    /* ignore */
}

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

/** IT-Leiste / Umschalter nur für echte Schul-IT. */
function itDashboardChrome() {
    if (!personas || !personas.loggedIn) return false;
    return !!personas.isIt;
}

function isToolVisible(toolId) {
    const rule = getEffectiveToolRule(toolId, loadDashboardToolAccessConfig());
    return isToolVisibleForView(toolId, plannerContext(), catalogView(), rule);
}

function isToolLicenseAllowed(toolId) {
    const id = String(toolId || '').trim();
    if (!id) return true;
    try {
        if (typeof window.ms365LicenseHasFeature === 'function') {
            return !!window.ms365LicenseHasFeature(id);
        }
        const core = window.ms365LicenseGateCore;
        if (core && typeof core.hasFeature === 'function') {
            return !!core.hasFeature(id);
        }
    } catch {
        /* ignore */
    }
    return true;
}

export function isToolElementVisible(card) {
    const id = String(card.getAttribute('data-tool-id') || '').trim();
    if (!id) return true;
    if (!isToolLicenseAllowed(id)) return false;
    if (layoutMode === 'full') return true;
    return isToolVisible(id);
}

function appOnlyStackEl() {
    return document.getElementById('dashAppOnlyStack');
}

function syncCatalogCardsHidden() {
    const catalog = document.getElementById('dash-catalog');
    if (!catalog) return;
    const searching = document.body.classList.contains('dashboard-searching');
    const full = layoutMode === 'full';
    catalog.querySelectorAll('.choice[data-tool-id]').forEach((card) => {
        const id = String(card.getAttribute('data-tool-id') || '').trim();
        const licOk = isToolLicenseAllowed(id);
        const audOk = full ? true : isToolVisible(id);
        card.toggleAttribute('data-dash-license-hidden', !licOk);
        card.toggleAttribute('data-dash-audience-hidden', !audOk);
        if (!searching) card.hidden = !(licOk && audOk);
    });
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function portalDisplayName() {
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            return String((info && (info.name || info.username)) || '').trim();
        }
    } catch {
        /* ignore */
    }
    return '';
}

function renderToolCard(id) {
    const href = DASHBOARD_TOOL_LINKS[id] || `tools/${id}.html`;
    const icon = DASHBOARD_TOOL_ICONS[id] || 'bi-box-arrow-up-right';
    const blurb = DASHBOARD_TOOL_BLURBS[id] || '';
    return `<a class="dash-quick-apps__card" href="${escapeHtml(href)}" data-tool-id="${escapeHtml(id)}">
                <div class="dash-quick-apps__card-head">
                  <h3><i class="bi ${escapeHtml(icon)}" aria-hidden="true"></i>${escapeHtml(toolLabel(id))}</h3>
                  <i class="bi bi-chevron-right dash-quick-apps__chev" aria-hidden="true"></i>
                </div>
                ${blurb ? `<p class="dash-quick-apps__card-desc">${escapeHtml(blurb)}</p>` : ''}
                <span class="dash-quick-apps__card-action">Öffnen <i class="bi bi-box-arrow-up-right" aria-hidden="true"></i></span>
            </a>`;
}

function renderPortalLinksStrip(links) {
    const active = activePortalQuickLinks(links);
    const chips = active
        .map((l) => {
            const target = l.external === false ? '' : ' target="_blank" rel="noopener noreferrer"';
            return `<a class="dash-portal-links__chip" href="${escapeHtml(l.url)}"${target} data-portal-link="${escapeHtml(l.id)}">
                <i class="bi ${escapeHtml(l.icon || 'bi-link-45deg')}" aria-hidden="true"></i>
                <span>${escapeHtml(l.label)}</span>
            </a>`;
        })
        .join('');
    const emptyHint = active.length
        ? ''
        : '<p class="dash-portal-links__empty">Noch keine Schul-Links – über „Links bearbeiten“ WebUntis, Sokrates oder Intranet ergänzen.</p>';
    return `<section class="dash-portal-links" aria-label="Schnellzugriff Links">
        <div class="dash-portal-links__head">
            <h2 class="dash-portal-links__title">Schnellzugriff</h2>
            <button type="button" class="dash-portal-links__edit btn btn-ghost" data-portal-links-edit>Links bearbeiten</button>
        </div>
        <div class="dash-portal-links__chips">${chips}</div>
        ${emptyHint}
    </section>`;
}

function openPortalLinksEditor() {
    const current = loadPortalQuickLinks();
    const schoolIds = new Set(PORTAL_QUICK_LINK_SCHOOL_IDS);
    const rows = PORTAL_QUICK_LINK_DEFAULTS.map((d) => {
        const cur = current.find((x) => x.id === d.id) || d;
        const editable = schoolIds.has(d.id) || !d.url;
        return `<label class="dash-portal-links-edit__row">
            <span>${escapeHtml(cur.label)}</span>
            <input type="url" data-portal-link-id="${escapeHtml(d.id)}" value="${escapeHtml(cur.url)}" placeholder="${editable ? 'https://…' : ''}" ${editable ? '' : 'readonly'} />
        </label>`;
    }).join('');

    let dialog = document.getElementById('dashPortalLinksDialog');
    if (!dialog) {
        dialog = document.createElement('dialog');
        dialog.id = 'dashPortalLinksDialog';
        dialog.className = 'dash-portal-links-dialog';
        document.body.appendChild(dialog);
    }
    dialog.innerHTML = `
        <form method="dialog" class="dash-portal-links-edit">
            <h3>Schnellzugriff-Links</h3>
            <p class="muted">Outlook, Teams und Kalender sind vorausgefüllt. Schul-Links (WebUntis, Sokrates, Intranet) legt die Schule fest – Speicherung im Browser bzw. in den Stammdaten der Schule.</p>
            <div class="dash-portal-links-edit__fields">${rows}</div>
            <div class="dash-portal-links-edit__actions">
                <button type="submit" value="cancel" class="btn btn-ghost">Abbrechen</button>
                <button type="submit" value="save" class="btn">Speichern</button>
            </div>
        </form>`;
    dialog.addEventListener(
        'close',
        () => {
            if (dialog.returnValue !== 'save') return;
            const next = PORTAL_QUICK_LINK_DEFAULTS.map((d) => {
                const input = dialog.querySelector(`[data-portal-link-id="${d.id}"]`);
                const cur = current.find((x) => x.id === d.id) || d;
                return {
                    ...cur,
                    url: input ? String(input.value || '').trim() : cur.url
                };
            });
            savePortalQuickLinks(next);
            renderQuickstart();
        },
        { once: true }
    );
    if (typeof dialog.showModal === 'function') dialog.showModal();
    else dialog.setAttribute('open', '');
}

/**
 * @param {'lehrer'|'schueler'} view
 */
function visiblePortalSections(view) {
    const sections = DASHBOARD_PORTAL_SECTIONS_BY_VIEW[view] || [];
    const filtered = sections
        .map((sec) => {
            const toolIds = (sec.toolIds || []).filter(
                (id) => isToolVisible(id) && isToolLicenseAllowed(id)
            );
            return { ...sec, toolIds };
        })
        .filter((sec) => sec.toolIds.length > 0);
    const defaults = loadPortalDefaultsConfig();
    const order =
        view === 'lehrer' || view === 'schueler' ? defaults[view].sectionOrder : [];
    return applyPortalSectionOrder(filtered, order);
}

function setPortalNavId(id) {
    portalNavId = String(id || 'home').trim() || 'home';
    try {
        sessionStorage.setItem(PORTAL_NAV_KEY, portalNavId);
    } catch {
        /* ignore */
    }
}

function renderItAreaCtaCard() {
    return `<button type="button" class="dash-quick-apps__card dash-portal-it-cta" data-open-it-shell>
                <div class="dash-quick-apps__card-head">
                  <h3><i class="bi bi-grid-3x3-gap-fill" aria-hidden="true"></i>IT-Bereich</h3>
                  <i class="bi bi-chevron-right dash-quick-apps__chev" aria-hidden="true"></i>
                </div>
                <p class="dash-quick-apps__card-desc">Vollständiger Werkzeugkatalog, Stammdaten und Organigramm – nur für Schul-IT.</p>
                <span class="dash-quick-apps__card-action">Öffnen <i class="bi bi-box-arrow-up-right" aria-hidden="true"></i></span>
            </button>`;
}

function readSchoolCardInput() {
    let schoolName = '';
    let domain = '';
    let schoolProfile = null;
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad();
            schoolName = String((s && s.schoolName) || '').trim();
            domain = String((s && s.domain) || '').trim();
        }
    } catch {
        /* ignore */
    }
    try {
        const api = window.ms365AppDataV2;
        if (api && typeof api.getContainer === 'function') {
            const c = api.getContainer();
            const core = c && c.core ? c.core : {};
            if (!schoolName) schoolName = String(core.schoolName || '').trim();
            if (!domain) domain = String(core.domain || '').trim();
        }
        if (api && typeof api.getSetup === 'function') {
            const setup = api.getSetup();
            schoolProfile = setup && setup.schoolProfile ? setup.schoolProfile : null;
        }
    } catch {
        /* ignore */
    }
    return { schoolName, domain, schoolProfile };
}

function formatPortalClock(now = new Date()) {
    try {
        const date = new Intl.DateTimeFormat('de-AT', {
            weekday: 'long',
            day: 'numeric',
            month: 'long',
            year: 'numeric'
        }).format(now);
        const time = new Intl.DateTimeFormat('de-AT', {
            hour: '2-digit',
            minute: '2-digit'
        }).format(now);
        return { date, time };
    } catch {
        return {
            date: now.toLocaleDateString('de-AT'),
            time: now.toLocaleTimeString('de-AT', { hour: '2-digit', minute: '2-digit' })
        };
    }
}

function renderPortalSchoolCard() {
    const card = buildDashboardSchoolCard(readSchoolCardInput());
    const clock = formatPortalClock();
    const name = card.schoolName || 'Schule noch nicht hinterlegt';
    const logoSrc = card.logoSrc || '';
    const logo = logoSrc
        ? `<img class="dash-portal-school__logo" src="${escapeHtml(logoSrc)}" alt="" width="56" height="56" referrerpolicy="no-referrer">`
        : `<div class="dash-portal-school__logo-fallback" aria-hidden="true"><i class="bi bi-building"></i></div>`;
    const metaParts = [];
    if (card.addressLine) {
        metaParts.push(
            `<p class="dash-portal-school__meta"><i class="bi bi-geo-alt" aria-hidden="true"></i>${escapeHtml(card.addressLine)}</p>`
        );
    }
    if (card.websiteHref) {
        metaParts.push(
            `<p class="dash-portal-school__meta"><i class="bi bi-globe2" aria-hidden="true"></i><a href="${escapeHtml(card.websiteHref)}" target="_blank" rel="noopener noreferrer">${escapeHtml(card.websiteLabel)}</a></p>`
        );
    } else if (card.domain) {
        metaParts.push(
            `<p class="dash-portal-school__meta"><i class="bi bi-at" aria-hidden="true"></i>${escapeHtml(card.domain)}</p>`
        );
    }
    if (card.profile.phone) {
        metaParts.push(
            `<p class="dash-portal-school__meta"><i class="bi bi-telephone" aria-hidden="true"></i>${escapeHtml(card.profile.phone)}</p>`
        );
    }
    if (!metaParts.length && !card.hasIdentity) {
        metaParts.push(
            `<p class="dash-portal-school__meta dash-portal-school__meta--muted">Name, Logo und Adresse unter „Portal vorbereiten“ oder in den Stammdaten pflegen.</p>`
        );
    }
    return `
        <aside class="dash-portal-school" aria-label="Schule und Uhrzeit">
            <div class="dash-portal-school__clock">
                <p class="dash-portal-school__time" data-portal-clock-time>${escapeHtml(clock.time)}</p>
                <p class="dash-portal-school__date" data-portal-clock-date>${escapeHtml(clock.date)}</p>
            </div>
            <div class="dash-portal-school__body">
                ${logo}
                <div class="dash-portal-school__text">
                    <p class="dash-portal-school__eyebrow">Meine Schule</p>
                    <p class="dash-portal-school__name">${escapeHtml(name)}</p>
                    ${metaParts.join('')}
                </div>
            </div>
        </aside>`;
}

function bindPortalSchoolClock(wrap) {
    if (!wrap) return;
    if (wrap._portalClockTimer) {
        clearInterval(wrap._portalClockTimer);
        wrap._portalClockTimer = null;
    }
    const tick = () => {
        const clock = formatPortalClock();
        const timeEl = wrap.querySelector('[data-portal-clock-time]');
        const dateEl = wrap.querySelector('[data-portal-clock-date]');
        if (timeEl) timeEl.textContent = clock.time;
        if (dateEl) dateEl.textContent = clock.date;
    };
    tick();
    wrap._portalClockTimer = setInterval(tick, 30000);
}

function renderPortalHome(view, sections, links) {
    const firstName = extractFirstName(portalDisplayName());
    const role = viewLabel(view);
    const greeting = greetingForHour();
    const welcomeLine = firstName
        ? `${greeting}, ${escapeHtml(firstName)}!`
        : `${greeting}!`;

    // Schul-Pins zuerst, dann persönliche Favoriten, dann Produkt-Fallback.
    const homePriority = [
        'schularbeiten-planer',
        'freistellung-planer',
        'lehrer-freistellung-planer',
        'schulaktivitaeten-planer',
        'termine-hub',
        'srdp-anmeldung',
        'elternbrief',
        'aufgaben',
        'projektwochen'
    ];
    const visibleIds = sections.flatMap((sec) => sec.toolIds || []);
    const defaults = loadPortalDefaultsConfig();
    const schoolPinned =
        view === 'lehrer' || view === 'schueler' ? defaults[view].homePinned || [] : [];
    const unique = resolvePortalHomeToolIds({
        visibleIds,
        schoolPinned,
        personalFavorites: loadDashboardFavorites(),
        fallbackPriority: homePriority,
        max: 6
    });
    const itCta = itDashboardChrome() ? renderItAreaCtaCard() : '';

    return `
        <header class="dash-portal-hero">
            <div class="dash-portal-hero__text">
                <p class="dash-portal-hero__eyebrow">${escapeHtml(role)}</p>
                <h1 class="dash-portal-hero__title">${welcomeLine}</h1>
                <p class="dash-portal-hero__lead">Hier findest du die wichtigsten Schultools für den Alltag.</p>
            </div>
            ${renderPortalSchoolCard()}
        </header>
        ${renderPortalLinksStrip(links)}
        <section class="dash-portal-section" aria-label="Meine Apps">
            <h2 class="dash-portal-section__title"><i class="bi bi-grid" aria-hidden="true"></i>Meine Apps</h2>
            <div class="dash-quick-apps__grid">
                ${unique.map(renderToolCard).join('')}
                ${itCta}
            </div>
        </section>`;
}

function renderPortalSectionPanel(sec) {
    return `
        <section class="dash-portal-section" data-portal-section="${escapeHtml(sec.id)}" aria-label="${escapeHtml(sec.label)}">
            <h2 class="dash-portal-section__title"><i class="bi ${escapeHtml(sec.icon || 'bi-grid')}" aria-hidden="true"></i>${escapeHtml(sec.label)}</h2>
            <div class="dash-quick-apps__grid">${sec.toolIds.map(renderToolCard).join('')}</div>
        </section>`;
}

function bindPortalShellEvents(wrap, view) {
    wrap.querySelectorAll('[data-portal-nav]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-portal-nav') || 'home';
            if (id === 'it-bereich') {
                setLayoutMode('full');
                return;
            }
            setPortalNavId(id);
            renderQuickstart();
        });
    });

    wrap.querySelectorAll('[data-open-it-shell]').forEach((btn) => {
        btn.addEventListener('click', () => setLayoutMode('full'));
    });

    const editBtn = wrap.querySelector('[data-portal-links-edit]');
    if (editBtn) {
        if (view === 'schueler') editBtn.hidden = true;
        else editBtn.addEventListener('click', () => openPortalLinksEditor());
    }
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
    let sections = visiblePortalSections(view);
    const flatFallback = DASHBOARD_QUICKSTART_BY_VIEW[view] || [];

    if (!sections.length) {
        const visible = flatFallback.filter((id) => isToolVisible(id) && isToolLicenseAllowed(id));
        if (!visible.length && !itDashboardChrome()) {
            if (wrap) wrap.hidden = true;
            if (stack) stack.hidden = true;
            return;
        }
        if (visible.length) {
            sections = [
                {
                    id: 'meine',
                    label: 'Meine Anwendungen',
                    icon: 'bi-grid',
                    toolIds: visible
                }
            ];
        }
    }

    const validNavIds = new Set(['home', ...sections.map((s) => s.id)]);
    if (itDashboardChrome()) validNavIds.add('it-bereich');
    if (!validNavIds.has(portalNavId)) portalNavId = 'home';

    if (!wrap) {
        wrap = document.createElement('section');
        wrap.id = 'dashQuickApps';
        wrap.className = 'dash-quick-apps dash-lehrer-portal';
        wrap.setAttribute('aria-label', 'Schul-Portal');
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
    wrap.classList.add('dash-lehrer-portal');
    if (stack) stack.hidden = false;

    const links = loadPortalQuickLinks();
    const activeSec = sections.find((s) => s.id === portalNavId);

    const navItems = [
        {
            id: 'home',
            label: 'Dashboard',
            icon: 'bi-house-door'
        },
        ...sections.map((s) => ({
            id: s.id,
            label: s.label,
            icon: s.icon || 'bi-grid'
        }))
    ];
    if (itDashboardChrome()) {
        navItems.push({
            id: 'it-bereich',
            label: 'IT-Bereich',
            icon: 'bi-grid-3x3-gap-fill',
            it: true
        });
    }

    const navHtml = navItems
        .map((item) => {
            const selected = item.id === portalNavId && item.id !== 'it-bereich';
            const cls = [
                'dashboard-tab',
                'dash-portal-shell__tab',
                item.it ? 'dash-portal-shell__tab--it' : '',
                selected ? '' : ''
            ]
                .filter(Boolean)
                .join(' ');
            return `<button type="button" class="${cls}" role="tab" data-portal-nav="${escapeHtml(item.id)}" aria-selected="${selected ? 'true' : 'false'}"${selected ? '' : ' tabindex="-1"'}>
                <i class="bi ${escapeHtml(item.icon)}" aria-hidden="true"></i>${escapeHtml(item.label)}
            </button>`;
        })
        .join('');

    let mainHtml;
    if (portalNavId === 'home' || !activeSec) {
        mainHtml = renderPortalHome(view, sections, links);
    } else {
        mainHtml = renderPortalSectionPanel(activeSec);
    }

    wrap.innerHTML = `
        <div class="dash-portal-shell" role="region" aria-label="Schulportal">
            <div class="dash-portal-shell__nav dashboard-tablist" role="tablist" aria-label="Portal-Bereiche">
                ${navHtml}
            </div>
            <div class="dash-portal-shell__main" role="tabpanel">
                ${mainHtml}
            </div>
        </div>`;

    bindPortalShellEvents(wrap, view);
    if (portalNavId === 'home' || !activeSec) {
        bindPortalSchoolClock(wrap);
    } else if (wrap._portalClockTimer) {
        clearInterval(wrap._portalClockTimer);
        wrap._portalClockTimer = null;
    }
}

/** IT-only Header-Chrome (auch wenn IT die Lehrer-/Schüler-Vorschau nutzt). */
const APPS_ONLY_HIDDEN_HEADER_IDS = [
    'dashHeaderNavMount',
    'dashHeaderSearchSlot',
    'dashHeaderNavMenuBtn',
    'dashMainSectionToggle',
    'ms365BackupHeader',
    'ms365HeaderSchulOrganigramm',
    'ms365HeaderSchulregister',
    'dashRegisterToolbar',
    'dashGreetingBand'
];

function applyAppsOnlyHeaderChrome(appOnly) {
    APPS_ONLY_HIDDEN_HEADER_IDS.forEach((id) => {
        const el = document.getElementById(id);
        if (!el) return;
        el.hidden = !!appOnly;
        el.setAttribute('aria-hidden', appOnly ? 'true' : 'false');
        if (appOnly) el.setAttribute('data-dash-section-audience', 'it');
    });
    document.querySelectorAll('#dashHeaderNavDrawer, #dashHeaderNavDrawerBackdrop').forEach((el) => {
        el.hidden = true;
    });
    document.body.classList.remove('dash-header-nav-drawer-open');
}

function removeItShellBanner() {
    document.getElementById('dashItShellBanner')?.remove();
}

function applyShellAttributes(appOnly, full) {
    const root = document.documentElement;
    const itShell = full && itDashboardChrome();
    root.setAttribute('data-dash-view', catalogView());
    root.toggleAttribute('data-dash-filtering', !full);
    root.toggleAttribute('data-dash-app-only', appOnly);
    root.toggleAttribute('data-dash-it-shell', itShell);
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

    applyShellAttributes(appOnly, full);
    applyAppsOnlyHeaderChrome(appOnly);
    removeItShellBanner();
    syncCatalogCardsHidden();
    renderQuickstart();
    removeLegacyDashboardChrome();
    syncDashboardAuthMenuView();

    if (typeof window.ms365DashboardAudienceRefreshSearch === 'function') {
        window.ms365DashboardAudienceRefreshSearch();
    }
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-dash-app-only-changed', { detail: { appOnly, layoutMode } })
        );
    } catch {
        /* ignore */
    }
}

function removeLegacyDashboardChrome() {
    document.getElementById('dashPersonaBar')?.remove();
    document.getElementById('dashItPreviewBanner')?.remove();
}

function syncDashboardAuthMenuView() {
    syncAuthMenuAudienceUi(personas, layoutMode);
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
    try {
        window.dispatchEvent(new CustomEvent('ms365-it-shell-changed', { detail: { mode } }));
    } catch {
        /* ignore */
    }
}

function applyPersonas(next) {
    personas = next;

    if (!personas || !personas.loggedIn) {
        layoutMode = 'preview-lehrer';
        applyAuthMenuItAudience(false);
        applyDomFilter();
        return;
    }

    // IT: Speicher (Portal zuerst / IT-Bereich / Schüler-Vorschau); sonst apps-*
    layoutMode = resolveLayoutModeFromPersonas(personas);

    applyAuthMenuItAudience(!!(personas && personas.isIt));
    applyDomFilter();
}

async function refreshPersonas(opts) {
    const next = await getDashboardPersonas(opts);
    applyPersonas(next);
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
        debug: () => ({ personas, layoutMode, catalogView: catalogView(), portalNavId })
    };

    // Portal-first bevor Personas resolved sind
    applyShellAttributes(true, false);
    applyAppsOnlyHeaderChrome(true);

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

    window.addEventListener('ms365-dashboard-portal-defaults-changed', () => {
        if (isAppsOnlyLayout(layoutMode)) renderQuickstart();
        else applyDomFilter();
    });

    window.addEventListener('ms365-license-changed', () => {
        // Designierter IT-Kontakt / Operator oft erst hier – Persona neu, nicht nur DOM-Filter
        clearDashboardPersonaCache();
        refreshPersonas({ force: true });
    });

    window.addEventListener('ms365-tenant-settings-changed', () => {
        clearDashboardPersonaCache();
        refreshPersonas({ force: true });
        if (isAppsOnlyLayout(layoutMode)) renderQuickstart();
    });

    window.addEventListener('ms365-auth-widget-ready', () => {
        syncDashboardAuthMenuView();
        applyAppsOnlyHeaderChrome(isAppsOnlyLayout(layoutMode));
    });
}
