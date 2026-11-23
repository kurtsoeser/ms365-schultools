/**
 * Schul-IT-Dashboard: eine Ansicht (Aufgaben-/Gruppen-Shell).
 * Der frühere Toggle Aufgaben / Playbooks / Katalog entfällt –
 * Playbooks hängen in den Gruppen, der Katalog bleibt für die Suche.
 */
const STORAGE_KEY = 'ms365-dash-main-section-v1';

/** @typedef {'tasks'|'catalog'|'playbooks'} DashMainSection */

/** @param {unknown} v */
function normalizeSection(v) {
    if (v === 'tasks' || v === 'catalog' || v === 'playbooks') return v;
    return null;
}

/** Immer Tasks – Katalog/Playbooks-Tabs sind entfernt. */
export function readMainSection() {
    return 'tasks';
}

/** @param {DashMainSection} [_section] */
export function writeMainSection(_section) {
    try {
        localStorage.setItem(STORAGE_KEY, 'tasks');
    } catch {
        /* ignore */
    }
    applyMainSectionDom('tasks');
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-dash-main-section-changed', { detail: { section: 'tasks' } })
        );
    } catch {
        /* ignore */
    }
    return 'tasks';
}

/** @param {DashMainSection} [_section] */
export function applyMainSectionDom(_section) {
    const root = document.documentElement;
    if (!root) return;
    root.setAttribute('data-dash-main-section', 'tasks');
    const heading = document.getElementById('dashMainHeading');
    if (heading) heading.textContent = 'IT-Bereiche';
}

function itChromeVisible() {
    try {
        if (document.documentElement.hasAttribute('data-dash-app-only')) return false;
        const api = window.ms365DashboardAudience;
        if (api && typeof api.getLayoutMode === 'function') {
            const mode = api.getLayoutMode();
            if (
                mode === 'apps-lehrer' ||
                mode === 'apps-schueler' ||
                mode === 'preview-lehrer' ||
                mode === 'preview-schueler'
            ) {
                return false;
            }
        }
        const personas = api && typeof api.getPersonas === 'function' ? api.getPersonas() : null;
        if (personas && personas.loggedIn && personas.isIt) return true;
    } catch {
        /* ignore */
    }
    const tasks = document.getElementById('dashboard-tasks');
    if (tasks && tasks.getAttribute('data-dash-section-audience') === 'it' && !tasks.hidden) return true;
    return false;
}

/** Entfernt Relikte des alten Ansichts-Toggles. */
function cleanupLegacyMainSectionUi() {
    const toggle = document.getElementById('dashMainSectionToggle');
    if (toggle) toggle.remove();
    const menuBtn = document.getElementById('dashHeaderNavMenuBtn');
    if (menuBtn) menuBtn.remove();
    const drawer = document.getElementById('dashHeaderNavDrawer');
    if (drawer) drawer.remove();
    const backdrop = document.getElementById('dashHeaderNavDrawerBackdrop');
    if (backdrop) backdrop.remove();
    document.body.classList.remove('dash-header-nav-drawer-open');
    const headerNav = document.getElementById('dashHeaderNavMount');
    if (headerNav) {
        headerNav.hidden = true;
        headerNav.setAttribute('aria-hidden', 'true');
        headerNav.innerHTML = '';
    }
    const contentToggleMount = document.querySelector('.dashboard-tasks-head-tools');
    if (contentToggleMount) contentToggleMount.hidden = true;
}

export function mountDashboardMainSection() {
    const onDashboard = document.body && document.body.classList.contains('page-dashboard');
    if (!onDashboard) return;

    cleanupLegacyMainSectionUi();

    function refresh() {
        cleanupLegacyMainSectionUi();
        applyMainSectionDom('tasks');
        const showChrome = itChromeVisible();
        const headRow = document.querySelector('.dashboard-tasks-head-row');
        if (headRow) headRow.hidden = !showChrome;
    }

    window.ms365DashMainSection = {
        read: readMainSection,
        set: writeMainSection
    };

    window.addEventListener('ms365-dashboard-persona-ready', refresh);
    window.addEventListener('ms365-auth-state-changed', refresh);
    window.addEventListener('ms365-dash-main-section-changed', refresh);
    window.addEventListener('ms365-dash-app-only-changed', refresh);

    refresh();
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', mountDashboardMainSection);
    } else {
        mountDashboardMainSection();
    }
}
