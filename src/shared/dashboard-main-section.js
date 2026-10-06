/**
 * Hauptansicht Schul-IT-Dashboard: Aufgaben, Playbooks, Werkzeugkatalog.
 */
const STORAGE_KEY = 'ms365-dash-main-section-v1';
const LEGACY_IT_FOCUS_KEY = 'ms365-dash-it-focus-v1';

/** @typedef {'tasks'|'catalog'|'playbooks'} DashMainSection */

const SECTION_HEADINGS = {
    tasks: 'Was möchten Sie tun?',
    catalog: 'Alle Werkzeuge',
    playbooks: 'Playbooks'
};

/** @param {unknown} v */
function normalizeSection(v) {
    if (v === 'tasks' || v === 'catalog' || v === 'playbooks') return v;
    return null;
}

export function readMainSection() {
    try {
        const v = normalizeSection(localStorage.getItem(STORAGE_KEY));
        if (v) return v;
        const legacy = localStorage.getItem(LEGACY_IT_FOCUS_KEY);
        if (legacy === '0') return 'catalog';
    } catch {
        /* ignore */
    }
    return 'tasks';
}

/** @param {DashMainSection} section */
export function writeMainSection(section) {
    const next = normalizeSection(section) || 'tasks';
    try {
        localStorage.setItem(STORAGE_KEY, next);
    } catch {
        /* ignore */
    }
    applyMainSectionDom(next);
    try {
        window.dispatchEvent(new CustomEvent('ms365-dash-main-section-changed', { detail: { section: next } }));
    } catch {
        /* ignore */
    }
    return next;
}

/** @param {DashMainSection} section */
export function applyMainSectionDom(section) {
    const root = document.documentElement;
    if (!root) return;
    const next = normalizeSection(section) || 'tasks';
    root.setAttribute('data-dash-main-section', next);
    const heading = document.getElementById('dashMainHeading');
    if (heading) {
        heading.textContent = SECTION_HEADINGS[next] || SECTION_HEADINGS.tasks;
    }
}

function itChromeVisible() {
    try {
        const api = window.ms365DashboardAudience;
        const personas = api && typeof api.getPersonas === 'function' ? api.getPersonas() : null;
        if (personas && personas.loggedIn && personas.isIt) return true;
    } catch {
        /* ignore */
    }
    const tasks = document.getElementById('dashboard-tasks');
    if (tasks && tasks.getAttribute('data-dash-section-audience') === 'it' && !tasks.hidden) return true;
    return false;
}

function syncToggleUi(wrap, section) {
    if (!wrap) return;
    wrap.querySelectorAll('[data-dash-main-section]').forEach(function (btn) {
        const id = btn.getAttribute('data-dash-main-section');
        const on = id === section;
        btn.setAttribute('aria-pressed', on ? 'true' : 'false');
        btn.classList.toggle('active', on);
    });
}

/** @param {HTMLElement} desktopToggle */
function mountMobileMainSectionMenu(desktopToggle) {
    const actions = document.getElementById('adminAppTopActions');
    if (!actions || document.getElementById('dashHeaderNavMenuBtn')) return;

    const menuBtn = document.createElement('button');
    menuBtn.type = 'button';
    menuBtn.id = 'dashHeaderNavMenuBtn';
    menuBtn.className = 'dash-header-nav-menu-btn';
    menuBtn.setAttribute('aria-label', 'Dashboard-Ansicht wählen');
    menuBtn.setAttribute('aria-haspopup', 'dialog');
    menuBtn.setAttribute('aria-expanded', 'false');
    menuBtn.setAttribute('aria-controls', 'dashHeaderNavDrawer');
    menuBtn.innerHTML = '<i class="bi bi-list" aria-hidden="true"></i>';
    actions.insertBefore(menuBtn, actions.firstChild);

    const backdrop = document.createElement('div');
    backdrop.className = 'dash-header-nav-drawer-backdrop';
    backdrop.id = 'dashHeaderNavDrawerBackdrop';
    backdrop.hidden = true;

    const drawer = document.createElement('div');
    drawer.className = 'dash-header-nav-drawer';
    drawer.id = 'dashHeaderNavDrawer';
    drawer.setAttribute('role', 'dialog');
    drawer.setAttribute('aria-label', 'Dashboard-Ansicht');
    drawer.hidden = true;
    drawer.innerHTML =
        '<div class="dash-header-nav-drawer__head">' +
        '<span>Ansicht</span>' +
        '<button type="button" class="dash-header-nav-drawer__close" id="dashHeaderNavDrawerClose" aria-label="Schließen">' +
        '<i class="bi bi-x-lg" aria-hidden="true"></i></button></div>' +
        '<div class="dash-header-nav-drawer__nav" id="dashHeaderNavDrawerNav"></div>';

    document.body.appendChild(backdrop);
    document.body.appendChild(drawer);

    const drawerNav = drawer.querySelector('#dashHeaderNavDrawerNav');
    if (drawerNav && desktopToggle) {
        desktopToggle.querySelectorAll('[data-dash-main-section]').forEach(function (srcBtn) {
            const clone = document.createElement('button');
            clone.type = 'button';
            clone.className = 'dash-header-nav-drawer__item';
            clone.setAttribute('data-dash-main-section', srcBtn.getAttribute('data-dash-main-section') || '');
            clone.textContent = srcBtn.textContent || '';
            drawerNav.appendChild(clone);
        });
    }

    function closeDrawer() {
        drawer.hidden = true;
        backdrop.hidden = true;
        menuBtn.setAttribute('aria-expanded', 'false');
        document.body.classList.remove('dash-header-nav-drawer-open');
    }

    function openDrawer() {
        drawer.hidden = false;
        backdrop.hidden = false;
        menuBtn.setAttribute('aria-expanded', 'true');
        document.body.classList.add('dash-header-nav-drawer-open');
        syncMobileMainSectionMenu(readMainSection());
    }

    menuBtn.addEventListener('click', function (e) {
        e.stopPropagation();
        if (drawer.hidden) openDrawer();
        else closeDrawer();
    });
    backdrop.addEventListener('click', closeDrawer);
    const closeBtn = drawer.querySelector('#dashHeaderNavDrawerClose');
    if (closeBtn) {
        closeBtn.addEventListener('click', function (e) {
            e.preventDefault();
            e.stopPropagation();
            closeDrawer();
        });
    }

    document.addEventListener('keydown', function (e) {
        if (e.key === 'Escape' && !drawer.hidden) closeDrawer();
    });

    closeDrawer();

    drawer.addEventListener('click', function (e) {
        const btn = e.target.closest('[data-dash-main-section]');
        if (!btn) return;
        const next = normalizeSection(btn.getAttribute('data-dash-main-section')) || 'tasks';
        writeMainSection(next);
        closeDrawer();
    });

    window.__ms365DashCloseNavDrawer = closeDrawer;
}

/** @param {DashMainSection} section */
function syncMobileMainSectionMenu(section) {
    const drawerNav = document.getElementById('dashHeaderNavDrawerNav');
    if (!drawerNav) return;
    drawerNav.querySelectorAll('[data-dash-main-section]').forEach(function (btn) {
        const id = btn.getAttribute('data-dash-main-section');
        const on = id === section;
        btn.classList.toggle('active', on);
        btn.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
}

export function mountDashboardMainSection() {
    const headerNav = document.getElementById('dashHeaderNavMount');
    const mountTools = headerNav || document.querySelector('.dashboard-tasks-head-tools');
    if (!mountTools || mountTools.querySelector('#dashMainSectionToggle')) return;

    const onDashboard = document.body && document.body.classList.contains('page-dashboard');
    if (!onDashboard) return;
    const toggle = document.createElement('div');
    toggle.id = 'dashMainSectionToggle';
    toggle.className = 'dash-main-section-toggle header-nav';
    toggle.setAttribute('role', 'navigation');
    toggle.setAttribute('data-dash-section-audience', 'it');
    toggle.setAttribute('aria-label', 'Dashboard-Ansicht');
    toggle.innerHTML =
        '<button type="button" class="dash-main-section-toggle__btn nav-tab" data-dash-main-section="tasks" aria-pressed="false" title="Aufgaben mit Status und Empfehlungen">Was möchten Sie tun?</button>' +
        '<button type="button" class="dash-main-section-toggle__btn nav-tab" data-dash-main-section="playbooks" aria-pressed="false" title="Geführte Checklisten">Playbooks</button>' +
        '<button type="button" class="dash-main-section-toggle__btn nav-tab" data-dash-main-section="catalog" aria-pressed="false" title="Vollständiger Werkzeugkatalog">Alle Werkzeuge</button>';
    mountTools.appendChild(toggle);

    const contentToggleMount = document.querySelector('.dashboard-tasks-head-tools');
    if (headerNav && contentToggleMount) contentToggleMount.hidden = true;

    mountMobileMainSectionMenu(toggle);

    function refresh() {
        const showChrome = itChromeVisible();
        toggle.hidden = !showChrome;
        if (headerNav) headerNav.hidden = !showChrome;
        const headRow = document.querySelector('.dashboard-tasks-head-row');
        if (headRow) headRow.hidden = !showChrome;
        const menuBtn = document.getElementById('dashHeaderNavMenuBtn');
        if (menuBtn) menuBtn.hidden = !showChrome;
        if (!showChrome) {
            applyMainSectionDom('tasks');
            return;
        }
        const section = readMainSection();
        applyMainSectionDom(section);
        syncToggleUi(toggle, section);
        syncMobileMainSectionMenu(section);
    }

    toggle.addEventListener('click', function (e) {
        const btn = e.target.closest('[data-dash-main-section]');
        if (!btn) return;
        const next = normalizeSection(btn.getAttribute('data-dash-main-section')) || 'tasks';
        writeMainSection(next);
        refresh();
        if (typeof window.__ms365DashCloseNavDrawer === 'function') {
            window.__ms365DashCloseNavDrawer();
        }
    });

    window.addEventListener('ms365-dashboard-persona-ready', refresh);
    window.addEventListener('ms365-auth-state-changed', refresh);
    window.addEventListener('ms365-dash-main-section-changed', refresh);

    window.ms365DashMainSection = {
        read: readMainSection,
        set: writeMainSection
    };

    refresh();
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', mountDashboardMainSection);
    } else {
        mountDashboardMainSection();
    }
}
