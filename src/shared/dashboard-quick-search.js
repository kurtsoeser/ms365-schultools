/**
 * Mittige IT-Schnellsuche mit kategorisierten Treffern (Admin-Center-Stil).
 */
import { TAB_LABEL_BY_PANEL } from './dashboard-catalog-layout.js';
import {
    catalogCardMatchesQuery,
    entryFromCatalogCard,
    groupQuickSearchHits,
    rankQuickSearch
} from './dashboard-quick-search-logic.js';

const RECENT_KEY = 'ms365-dash-quick-search-recent-v1';

/** @returns {string[]} */
function loadRecentIds() {
    try {
        const raw = JSON.parse(localStorage.getItem(RECENT_KEY) || '[]');
        return Array.isArray(raw) ? raw.map(String).slice(0, 8) : [];
    } catch {
        return [];
    }
}

/** @param {string} id */
function pushRecentId(id) {
    const next = [id]
        .concat(
            loadRecentIds().filter(function (x) {
                return x !== id;
            })
        )
        .slice(0, 8);
    try {
        localStorage.setItem(RECENT_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
}

/** @returns {import('./dashboard-quick-search-logic.js').QuickSearchEntry[]} */
function collectStaticEntries() {
    /** @type {import('./dashboard-quick-search-logic.js').QuickSearchEntry[]} */
    const out = [
        {
            id: 'einstellung-stammdaten',
            title: 'Stammdaten',
            blurb: 'Schulprofil, Klassen, Import und Verknüpfungen',
            keywords: 'stammdaten tenant einstellungen profil',
            href: 'tenant.html',
            kind: 'einstellung',
            icon: 'bi bi-database'
        },
        {
            id: 'einstellung-organigramm',
            title: 'Schul-Organigramm',
            blurb: 'Rollen und Zuständigkeiten der Schule',
            keywords: 'organigramm rolle zuständigkeit',
            href: 'tools/schul-organigramm.html',
            kind: 'einstellung',
            icon: 'bi bi-diagram-3'
        },
        {
            id: 'einstellung-portal-vorbereiten',
            title: 'Portal vorbereiten',
            blurb: 'Lehrer-Portal gestalten: Logo, Home-Pins, Werkzeug-Zugriff und freigeben',
            keywords:
                'portal vorbereiten dashboard gestalten freigeben pins home logo werkzeug zugriff lehrerportal go live',
            href: 'dashboard-werkzeug-zugriff.html',
            kind: 'einstellung',
            icon: 'bi bi-sliders2-vertical'
        },
        {
            id: 'einstellung-portal',
            title: 'Portal-Ansicht',
            blurb: 'Lehrer-/Schüler-Portal (Startseite)',
            keywords: 'portal lehrer schüler start home ansicht',
            href: 'index.html',
            kind: 'einstellung',
            icon: 'bi bi-house'
        }
    ];

    const areaMap = {
        dashTaskImportVerknuepfen: { title: 'Stammdaten', keywords: 'import blaupause stammdaten' },
        dashTaskGruppen: { title: 'Mitgliedschaften', keywords: 'sammelgruppe slg mitgliedschaft' },
        dashTaskUnterricht: { title: 'Klassen & Unterricht', keywords: 'klasse kursteam unterricht' },
        dashTaskPersonen: { title: 'Personen', keywords: 'person gast lifecycle' },
        dashTaskSchuljahr: { title: 'Schuljahresstart', keywords: 'schuljahr wechsel playbook' },
        dashTaskIntranet: { title: 'Intranet', keywords: 'intranet sharepoint hub' },
        dashTaskSchulApps: {
            title: 'Schul-Apps',
            keywords: 'planer freistellung schularbeiten portal vorbereiten'
        },
        dashTaskOrdnung: { title: 'Aufräumen', keywords: 'cleanup hygiene audit' }
    };
    Object.keys(areaMap).forEach(function (taskId) {
        const meta = areaMap[taskId];
        out.push({
            id: 'bereich-' + taskId,
            taskId: taskId,
            title: meta.title,
            blurb: 'Bereich in der linken Navigation',
            keywords: meta.keywords,
            href: 'index.html',
            kind: 'bereich',
            icon: 'bi bi-layout-sidebar'
        });
    });
    return out;
}

/** @returns {import('./dashboard-quick-search-logic.js').QuickSearchEntry[]} */
function collectCatalogEntries() {
    const catalog = document.getElementById('dash-catalog');
    if (!catalog) return [];
    return Array.from(catalog.querySelectorAll('.choice[data-tool-id]'))
        .map(entryFromCatalogCard)
        .filter(Boolean)
        .map(function (entry) {
            const cluster = entry.group || '';
            const panelLabel = TAB_LABEL_BY_PANEL[cluster] || '';
            if (panelLabel) entry.group = panelLabel;
            return entry;
        });
}

function collectIndex() {
    const byId = Object.create(null);
    collectStaticEntries().concat(collectCatalogEntries()).forEach(function (e) {
        if (!e || !e.id) return;
        byId[e.id] = e;
    });
    return Object.keys(byId).map(function (k) {
        return byId[k];
    });
}

/**
 * @param {import('./dashboard-quick-search-logic.js').QuickSearchEntry} entry
 */
function activateEntry(entry) {
    if (!entry) return;
    pushRecentId(entry.id);

    if (entry.kind === 'bereich' && entry.taskId) {
        try {
            if (window.ms365DashMainSection && typeof window.ms365DashMainSection.set === 'function') {
                window.ms365DashMainSection.set('tasks');
            }
        } catch {
            /* ignore */
        }
        if (window.ms365DashTasksSplit && typeof window.ms365DashTasksSplit.select === 'function') {
            window.ms365DashTasksSplit.select(entry.taskId, { pulse: true });
            return;
        }
    }

    if (entry.href) {
        window.location.href = entry.href;
    }
}

/**
 * @param {HTMLInputElement} input
 */
export function mountDashboardQuickSearch(input) {
    if (!input || input.dataset.quickSearchBound === '1') return null;
    input.dataset.quickSearchBound = '1';

    const wrap = input.closest('.dashboard-search-wrap') || input.parentElement;
    const root = input.closest('.dashboard-search') || wrap;
    if (!wrap || !root) return null;

    const panel = document.createElement('div');
    panel.className = 'dash-quick-search';
    panel.id = 'dashQuickSearchPanel';
    panel.hidden = true;
    panel.setAttribute('role', 'listbox');
    panel.setAttribute('aria-label', 'Suchergebnisse');
    wrap.appendChild(panel);

    let activeIndex = -1;
    /** @type {import('./dashboard-quick-search-logic.js').QuickSearchHit[]} */
    let flatHits = [];

    function closePanel() {
        panel.hidden = true;
        panel.textContent = '';
        activeIndex = -1;
        flatHits = [];
        input.setAttribute('aria-expanded', 'false');
    }

    function setActive(idx) {
        activeIndex = idx;
        panel.querySelectorAll('.dash-quick-search__item').forEach(function (el, i) {
            const on = i === idx;
            el.classList.toggle('is-active', on);
            if (on) el.setAttribute('aria-selected', 'true');
            else el.removeAttribute('aria-selected');
        });
    }

    /**
     * @param {import('./dashboard-quick-search-logic.js').QuickSearchHit} hit
     * @param {number} idx
     */
    function appendItem(parent, hit, idx) {
        const e = hit.entry;
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'dash-quick-search__item';
        btn.setAttribute('role', 'option');
        btn.setAttribute('data-hit-index', String(idx));

        const icon = document.createElement('span');
        icon.className = 'dash-quick-search__icon';
        icon.setAttribute('aria-hidden', 'true');
        icon.innerHTML = '<i class="' + String(e.icon || 'bi bi-search').replace(/"/g, '') + '"></i>';

        const body = document.createElement('span');
        body.className = 'dash-quick-search__body';
        const title = document.createElement('span');
        title.className = 'dash-quick-search__title';
        title.textContent = e.title;
        body.appendChild(title);

        const metaText = e.group || (e.blurb ? e.blurb.slice(0, 72) : '');
        if (metaText) {
            const meta = document.createElement('span');
            meta.className = 'dash-quick-search__meta';
            meta.textContent = metaText;
            body.appendChild(meta);
        }

        btn.appendChild(icon);
        btn.appendChild(body);
        parent.appendChild(btn);
    }

    function render(query) {
        const q = String(query || '').trim();
        panel.textContent = '';

        if (!q) {
            const recentIds = loadRecentIds();
            const index = collectIndex();
            const byId = Object.create(null);
            index.forEach(function (e) {
                byId[e.id] = e;
            });
            const recentHits = recentIds
                .map(function (id) {
                    return byId[id] ? { entry: byId[id], score: 1, match: /** @type {'exact'} */ ('exact') } : null;
                })
                .filter(Boolean);
            flatHits = recentHits;
            if (!recentHits.length) {
                closePanel();
                return;
            }
            const group = document.createElement('div');
            group.className = 'dash-quick-search__group';
            const label = document.createElement('div');
            label.className = 'dash-quick-search__group-label';
            label.textContent = 'Zuletzt gewählt';
            group.appendChild(label);
            recentHits.forEach(function (hit, i) {
                appendItem(group, hit, i);
            });
            panel.appendChild(group);
            panel.hidden = false;
            input.setAttribute('aria-expanded', 'true');
            setActive(0);
            return;
        }

        const hits = rankQuickSearch(collectIndex(), q, { limit: 14 });
        if (!hits.length) {
            flatHits = [];
            const empty = document.createElement('div');
            empty.className = 'dash-quick-search__empty';
            empty.textContent =
                'Kein Treffer für „' + q + '“. Tippen Sie z. B. Portal, Dashboard, SLG oder Intranet.';
            panel.appendChild(empty);
            panel.hidden = false;
            input.setAttribute('aria-expanded', 'true');
            activeIndex = -1;
            return;
        }

        // flatHits muss der sichtbaren Reihenfolge entsprechen (Gruppierung ≠ Ranking-Liste),
        // sonst öffnet Klick/Enter den falschen Eintrag.
        const groups = groupQuickSearchHits(hits);
        flatHits = [];
        groups.forEach(function (g) {
            const group = document.createElement('div');
            group.className = 'dash-quick-search__group';
            const label = document.createElement('div');
            label.className = 'dash-quick-search__group-label';
            label.textContent = g.label;
            group.appendChild(label);
            g.hits.forEach(function (hit) {
                appendItem(group, hit, flatHits.length);
                flatHits.push(hit);
            });
            panel.appendChild(group);
        });
        panel.hidden = false;
        input.setAttribute('aria-expanded', 'true');
        setActive(0);
    }

    function openActive() {
        if (activeIndex < 0 || !flatHits[activeIndex]) return;
        const entry = flatHits[activeIndex].entry;
        closePanel();
        activateEntry(entry);
    }

    input.setAttribute('role', 'combobox');
    input.setAttribute('aria-autocomplete', 'list');
    input.setAttribute('aria-expanded', 'false');
    input.setAttribute('aria-controls', 'dashQuickSearchPanel');
    input.placeholder = 'Suchen: Tool, Playbook, Einstellung …';

    input.addEventListener('input', function () {
        render(input.value);
    });
    input.addEventListener('focus', function () {
        render(input.value);
    });
    input.addEventListener('keydown', function (e) {
        if (panel.hidden && (e.key === 'ArrowDown' || e.key === 'Enter')) {
            render(input.value);
        }
        if (panel.hidden) return;
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            setActive(Math.min(flatHits.length - 1, activeIndex + 1));
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            setActive(Math.max(0, activeIndex - 1));
        } else if (e.key === 'Enter') {
            if (activeIndex >= 0 && flatHits[activeIndex]) {
                e.preventDefault();
                openActive();
            }
        } else if (e.key === 'Escape') {
            e.preventDefault();
            if (String(input.value || '').trim()) {
                input.value = '';
                input.dispatchEvent(new Event('input', { bubbles: true }));
            }
            closePanel();
            input.blur();
        }
    });

    panel.addEventListener('mousedown', function (e) {
        const btn = e.target.closest('.dash-quick-search__item');
        if (!btn) return;
        e.preventDefault();
        const i = Number(btn.getAttribute('data-hit-index'));
        if (!Number.isFinite(i) || !flatHits[i]) return;
        activeIndex = i;
        openActive();
    });

    document.addEventListener('click', function (e) {
        if (root.contains(e.target)) return;
        closePanel();
    });

    return {
        render: render,
        close: closePanel,
        catalogCardMatchesQuery: catalogCardMatchesQuery
    };
}

export function bootDashboardQuickSearch() {
    const input = document.getElementById('dashSearch');
    if (!input) return null;
    return mountDashboardQuickSearch(/** @type {HTMLInputElement} */ (input));
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bootDashboardQuickSearch);
    } else {
        bootDashboardQuickSearch();
    }
}

window.ms365DashQuickSearch = {
    catalogCardMatchesQuery: catalogCardMatchesQuery,
    boot: bootDashboardQuickSearch
};
