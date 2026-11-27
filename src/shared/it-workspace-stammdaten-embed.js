/**
 * Auf tenant.html: gleiches IT-Menü wie überall.
 * - Zweite Stammdaten-Rail ausblenden
 * - Register-Unterpunkte der Stammdaten-Gruppe in-page umschalten (kein Extra-Menü)
 */
import { IT_WORKSPACE_STAMMDATEN_TASK_ID } from './it-workspace-nav.js';
import { stammdatenRegisterNavTools } from './stammdaten-nav-catalog.js';
import { buildRegisterHash, parseRegisterHash } from './stammdaten-register-mode.js';

const EMBED_CLASS = 'it-workspace-stammdaten-embed';

function isTenantPage() {
    try {
        const leaf = String(location.pathname || '')
            .replace(/\\/g, '/')
            .split('/')
            .filter(Boolean)
            .pop();
        return /^tenant\.html$/i.test(leaf || '');
    } catch {
        return false;
    }
}

/**
 * @returns {{ id: string, label: string, icon: string, group: string }[]}
 */
export function collectTenantRegisterNavItems() {
    const tabs = document.getElementById('tenantMainTabs');
    if (tabs) {
        /** @type {{ id: string, label: string, icon: string, group: string }[]} */
        const out = [];
        tabs.querySelectorAll('.tab-btn[role="tab"]').forEach(function (btn) {
            if (btn.hidden || btn.getAttribute('aria-hidden') === 'true') return;
            const id = String(btn.id || '').trim();
            if (!id) return;
            const iconEl = btn.querySelector('i.bi');
            const label = String(btn.textContent || '')
                .replace(/\s+/g, ' ')
                .trim();
            if (!label) return;
            out.push({
                id: id,
                label: label,
                icon: iconEl && iconEl.className ? String(iconEl.className) : 'bi bi-dot',
                group: String(btn.getAttribute('data-rail-group') || '').trim()
            });
        });
        if (out.length) return out;
    }
    return stammdatenRegisterNavTools().map(function (t) {
        return {
            id: t.tabBtnId,
            label: t.label,
            icon: t.icon,
            group: t.group
        };
    });
}

function activateTenantTab(tabBtnId) {
    const id = String(tabBtnId || '').trim();
    if (!id) return;
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-register-tab-request', { detail: { tabBtnId: id } })
        );
    } catch {
        if (typeof window.ms365TenantTabActivate === 'function') {
            window.ms365TenantTabActivate(id);
        }
    }
    const hash = buildRegisterHash(id);
    if (hash) {
        try {
            history.replaceState(null, '', '#' + hash);
        } catch {
            location.hash = hash;
        }
    }
    syncStammdatenEmbedCurrent();
}

function findStammdatenBlock(nav) {
    if (!nav) return null;
    const byId = nav.querySelector(
        '.it-workspace-shell__block[data-task-id="' + IT_WORKSPACE_STAMMDATEN_TASK_ID + '"]'
    );
    if (byId) return byId;
    const links = nav.querySelectorAll('a.it-workspace-shell__subitem[href]');
    for (let i = 0; i < links.length; i++) {
        const href = String(links[i].getAttribute('href') || '');
        if (/tenant\.html/i.test(href)) {
            return links[i].closest('.it-workspace-shell__block');
        }
    }
    return null;
}

function tabIdFromHref(href) {
    const raw = String(href || '');
    const hash = raw.includes('#') ? raw.slice(raw.indexOf('#') + 1) : '';
    if (!hash) return '';
    const parsed = parseRegisterHash(hash);
    return parsed && parsed.tabBtnId ? parsed.tabBtnId : '';
}

function syncStammdatenEmbedCurrent() {
    const block = findStammdatenBlock(document.getElementById('itWorkspaceNav'));
    if (!block) return;
    const active =
        document.querySelector('#tenantMainTabs .tab-btn[role="tab"][aria-selected="true"]') ||
        document.querySelector('#tenantMainTabs .tab-btn[role="tab"].active');
    const activeId = active ? String(active.id || '') : '';
    block.querySelectorAll('a.it-workspace-shell__subitem[href]').forEach(function (a) {
        const id = a.getAttribute('data-tenant-tab') || tabIdFromHref(a.getAttribute('href'));
        if (id && id === activeId) a.setAttribute('aria-current', 'page');
        else a.removeAttribute('aria-current');
    });
    const parent = block.querySelector(':scope > a.it-workspace-shell__item');
    if (parent && activeId) parent.removeAttribute('aria-current');
}

/**
 * Bestehende Shell-Unterlinks (tenant.html#…) für In-Page-Tabs verdrahten –
 * kein zweites Menü, keine „Weitere Bereiche“-Optik.
 * @returns {boolean}
 */
export function mountItWorkspaceStammdatenEmbed() {
    if (!isTenantPage()) {
        unmountItWorkspaceStammdatenEmbed();
        return false;
    }
    const nav = document.getElementById('itWorkspaceNav');
    const shell = document.getElementById('itWorkspaceShell');
    if (!nav || !shell || !document.body.classList.contains('it-workspace-shell')) {
        return false;
    }

    const block = findStammdatenBlock(nav);
    if (!block) return false;

    const sub = block.querySelector(':scope > .it-workspace-shell__sub');
    if (!sub) return false;

    // Alte Embed-Artefakte entfernen
    const leftover = document.getElementById('itWorkspaceStammdatenMoreLabel');
    if (leftover) leftover.remove();
    block.classList.remove('it-workspace-shell__block--stammdaten', 'it-workspace-shell__block--other');
    nav.querySelectorAll('.it-workspace-shell__block--other').forEach(function (el) {
        el.classList.remove('it-workspace-shell__block--other');
    });

    sub.querySelectorAll('a.it-workspace-shell__subitem[href]').forEach(function (a) {
        const href = String(a.getAttribute('href') || '');
        if (!/tenant\.html/i.test(href) && !/^#/.test(href)) return;
        const tabId = a.getAttribute('data-tenant-tab') || tabIdFromHref(href);
        if (!tabId) return;
        a.setAttribute('data-tenant-tab', tabId);
        if (a.dataset.tenantTabWired === '1') return;
        a.dataset.tenantTabWired = '1';
        a.addEventListener('click', function (ev) {
            ev.preventDefault();
            activateTenantTab(tabId);
        });
    });

    document.body.classList.add(EMBED_CLASS);
    syncStammdatenEmbedCurrent();

    if (!window.__ms365StammdatenEmbedTabWired) {
        window.__ms365StammdatenEmbedTabWired = true;
        document.addEventListener(
            'click',
            function (ev) {
                const t = ev.target;
                if (!t || !t.closest) return;
                if (!document.body.classList.contains(EMBED_CLASS)) return;
                if (t.closest('#tenantMainTabs .tab-btn')) {
                    requestAnimationFrame(syncStammdatenEmbedCurrent);
                }
            },
            true
        );
        window.addEventListener('ms365-register-tab-request', function () {
            requestAnimationFrame(syncStammdatenEmbedCurrent);
        });
        window.addEventListener('hashchange', function () {
            requestAnimationFrame(syncStammdatenEmbedCurrent);
        });
    }

    return true;
}

export function unmountItWorkspaceStammdatenEmbed() {
    const more = document.getElementById('itWorkspaceStammdatenMoreLabel');
    if (more) more.remove();
    if (typeof document !== 'undefined') {
        document.querySelectorAll('.it-workspace-shell__block--other').forEach(function (el) {
            el.classList.remove('it-workspace-shell__block--other');
        });
        document.querySelectorAll('.it-workspace-shell__block--stammdaten').forEach(function (el) {
            el.classList.remove('it-workspace-shell__block--stammdaten');
        });
        if (document.body) document.body.classList.remove(EMBED_CLASS);
    }
}

export function bootItWorkspaceStammdatenEmbed() {
    if (!isTenantPage()) return false;
    const tryMount = function () {
        if (mountItWorkspaceStammdatenEmbed()) return true;
        return false;
    };
    if (tryMount()) return true;
    let n = 0;
    const t = setInterval(function () {
        n += 1;
        if (tryMount() || n > 20) clearInterval(t);
    }, 100);
    return false;
}

if (typeof window !== 'undefined') {
    window.ms365BootItWorkspaceStammdatenEmbed = bootItWorkspaceStammdatenEmbed;
}
