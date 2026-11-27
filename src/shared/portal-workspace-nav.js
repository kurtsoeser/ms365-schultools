/**
 * Portal-Menü für Lehrkräfte und Schüler/innen.
 * Dieselbe Gruppenlogik wie auf der Startseite, auch auf der geöffneten Werkzeugseite.
 */
import {
    DASHBOARD_PORTAL_SECTIONS_BY_VIEW,
    DASHBOARD_TOOL_ICONS,
    DASHBOARD_TOOL_LINKS,
    isToolVisibleForView,
    toolLabel
} from './dashboard-audience-catalog.js';
import { getEffectiveToolRule } from './dashboard-audience-store.js';
import { readItShellChoice } from './dashboard-layout-mode.js';
import { normalizeWorkspaceHref, workspaceHrefsMatch } from './it-workspace-nav.js';
import { normalizeHttpsUrl, readSchoolIntranetOpenUrl } from './school-intranet-open.js';

/** Welcher Portal-Bereich auf der Startseite offen ist. */
export const PORTAL_SECTION_NAV_KEY = 'ms365-dashboard-portal-nav-v1';
/** Sichtbare Gruppen + Werkzeuge, wie sie das Portal zuletzt gezeigt hat. */
export const PORTAL_WORKSPACE_NAV_KEY = 'ms365-portal-workspace-nav-v1';

/**
 * @param {string} icon
 * @returns {string}
 */
export function normalizeIconClass(icon) {
    const raw = String(icon || '').trim();
    if (!raw) return 'bi bi-circle';
    if (/^bi bi-[a-z0-9-]+$/i.test(raw)) return raw;
    if (/^bi-[a-z0-9-]+$/i.test(raw)) return 'bi ' + raw;
    return 'bi bi-circle';
}

/**
 * Lehrkräfte und Schüler/innen öffnen beim Intranet die hinterlegte SharePoint-Seite.
 * Der Einrichten-Assistent bleibt der Schul-IT vorbehalten.
 * @param {string} toolId
 * @param {{ intranetUrl?: string, fallbackHref?: string }} [opts]
 * @returns {string}
 */
export function portalToolHref(toolId, opts) {
    const o = opts || {};
    if (toolId === 'sharepoint-intranet-hub') {
        if (Object.prototype.hasOwnProperty.call(o, 'intranetUrl')) {
            return normalizeHttpsUrl(o.intranetUrl);
        }
        return readSchoolIntranetOpenUrl();
    }
    const fallback = String(o.fallbackHref || '').trim();
    if (fallback) return fallback;
    return DASHBOARD_TOOL_LINKS[toolId] || 'tools/' + toolId + '.html';
}

/**
 * Relative Werkzeugseiten und die externe Intranet-Adresse.
 * @param {string} href
 * @returns {string}
 */
export function sanitizePortalNavHref(href) {
    const raw = String(href || '').trim();
    if (!raw) return '';
    if (/^https:\/\//i.test(raw)) return normalizeHttpsUrl(raw);
    if (/^(?:tools\/)?[a-z0-9-]+\.html(?:\?[^#]*)?$/i.test(raw)) return raw;
    return '';
}

/**
 * @param {Array<{ id?: string, label?: string, icon?: string, toolIds?: string[] }>} sections
 * @param {{ intranetUrl?: string }} [opts]
 * @returns {{ id: string, label: string, icon: string, tools: { id: string, href: string, label: string, icon: string }[] }[]}
 */
export function portalTreeFromSections(sections, opts) {
    const hrefOpts =
        opts && Object.prototype.hasOwnProperty.call(opts, 'intranetUrl')
            ? { intranetUrl: opts.intranetUrl }
            : undefined;
    return (sections || [])
        .map(function (sec) {
            const tools = (sec.toolIds || [])
                .map(function (id) {
                    const toolId = String(id || '').trim();
                    if (!toolId) return null;
                    return {
                        id: toolId,
                        href: portalToolHref(toolId, hrefOpts),
                        label: toolLabel(toolId),
                        icon: normalizeIconClass(DASHBOARD_TOOL_ICONS[toolId] || 'bi-box-arrow-up-right')
                    };
                })
                .filter(Boolean);
            return {
                id: String(sec.id || '').trim(),
                label: String(sec.label || '').trim(),
                icon: normalizeIconClass(sec.icon || 'bi-grid'),
                tools: tools
            };
        })
        .filter(function (sec) {
            return sec.id && sec.label && sec.tools.length;
        });
}

/**
 * Katalog-Fallback, wenn die Startseite den Baum noch nicht gespeichert hat.
 * @param {'lehrer'|'schueler'} view
 */
/**
 * Gespeicherte Leiste an die Katalog-Gruppen anlegen.
 * Ein Werkzeug, das die Gruppe wechselt, erscheint sofort in der neuen Sektion,
 * auch wenn die Sitzung noch den alten Baum hat. Nicht gespeicherte Werkzeuge
 * bleiben ausgeblendet.
 * @param {any[]} tree
 * @param {'lehrer'|'schueler'} view
 */
export function reconcilePortalTreeWithCatalog(tree, view) {
    const key = view === 'schueler' ? 'schueler' : 'lehrer';
    const catalog = DASHBOARD_PORTAL_SECTIONS_BY_VIEW[key] || [];
    const byId = new Map();
    (Array.isArray(tree) ? tree : []).forEach(function (sec) {
        ((sec && sec.tools) || []).forEach(function (tool) {
            const id = String((tool && tool.id) || '').trim();
            if (id && !byId.has(id)) byId.set(id, tool);
        });
    });
    return catalog
        .map(function (sec) {
            const tools = (sec.toolIds || [])
                .map(function (id) {
                    return byId.get(String(id || '').trim()) || null;
                })
                .filter(Boolean);
            return {
                id: String(sec.id || '').trim(),
                label: String(sec.label || '').trim(),
                icon: normalizeIconClass(sec.icon || 'bi-grid'),
                tools: tools
            };
        })
        .filter(function (sec) {
            return sec.id && sec.label && sec.tools.length;
        });
}

export function buildDefaultPortalTree(view) {
    const key = view === 'schueler' ? 'schueler' : 'lehrer';
    const sections = DASHBOARD_PORTAL_SECTIONS_BY_VIEW[key] || [];
    const visible = sections
        .map(function (sec) {
            return {
                id: sec.id,
                label: sec.label,
                icon: sec.icon,
                toolIds: (sec.toolIds || []).filter(function (id) {
                    return isToolVisibleForView(id, {}, key);
                })
            };
        })
        .filter(function (sec) {
            return sec.toolIds.length;
        });
    return portalTreeFromSections(visible);
}

/**
 * Entfernt Werkzeuge, die für diese Portal-Ansicht nicht mehr freigegeben sind
 * (z. B. alter Session-Baum, nachdem ein Setup-Werkzeug zur Schul-IT gezogen wurde).
 * @param {any[]} tree
 * @param {'lehrer'|'schueler'} view
 */
export function filterPortalTreeToVisible(tree, view) {
    const key = view === 'schueler' ? 'schueler' : 'lehrer';
    return (Array.isArray(tree) ? tree : [])
        .map(function (sec) {
            const tools = ((sec && sec.tools) || []).filter(function (tool) {
                const id = String((tool && tool.id) || '').trim();
                if (!id) return false;
                return isToolVisibleForView(id, {}, key, getEffectiveToolRule(id));
            });
            return Object.assign({}, sec, { tools: tools });
        })
        .filter(function (sec) {
            return sec && Array.isArray(sec.tools) && sec.tools.length;
        });
}

/**
 * @param {any[]} tree
 * @param {'lehrer'|'schueler'} view
 */
export function persistPortalWorkspaceNav(tree, view) {
    try {
        const cleanView = view === 'schueler' ? 'schueler' : 'lehrer';
        sessionStorage.setItem(
            PORTAL_WORKSPACE_NAV_KEY,
            JSON.stringify({
                version: 1,
                view: cleanView,
                tree: Array.isArray(tree) ? tree : []
            })
        );
    } catch {
        /* ignore */
    }
}

/**
 * @returns {{ version?: number, view?: string, tree: any[] }|null}
 */
export function loadPortalWorkspaceNav() {
    try {
        const raw = sessionStorage.getItem(PORTAL_WORKSPACE_NAV_KEY);
        if (!raw) return null;
        const data = JSON.parse(raw);
        if (!data || !Array.isArray(data.tree)) return null;
        if (data.view !== 'lehrer' && data.view !== 'schueler') return null;
        return data;
    } catch {
        return null;
    }
}

/**
 * @param {string} id
 */
export function rememberPortalSection(id) {
    const next = String(id || 'home').trim() || 'home';
    try {
        sessionStorage.setItem(PORTAL_SECTION_NAV_KEY, next);
    } catch {
        /* ignore */
    }
}

/**
 * @param {any[]} tree
 * @param {string} href
 * @param {boolean} loose
 * @returns {{ section: any, tool: any }|null}
 */
function findInTree(tree, href, loose) {
    const opts = loose ? { looseFile: true } : undefined;
    for (let i = 0; i < tree.length; i++) {
        const section = tree[i];
        const tools = (section && section.tools) || [];
        for (let j = 0; j < tools.length; j++) {
            if (workspaceHrefsMatch(tools[j].href, href, opts)) {
                return { section: section, tool: tools[j] };
            }
        }
    }
    return null;
}

/**
 * @param {any[]} tree
 * @param {string} href
 * @returns {{ section: any, tool: any }|null}
 */
export function resolvePortalLocation(tree, href) {
    const list = Array.isArray(tree) ? tree : [];
    const current = normalizeWorkspaceHref(href);
    if (!current) return null;
    return findInTree(list, current, false) || findInTree(list, current, true);
}

/**
 * Portal-Leiste: zuerst die aktive Ansicht, danach die andere.
 * So bekommt ein Schüler-Formular seine Leiste, auch wenn die Vorschau gerade „Lehrkraft“ ist.
 * @param {'lehrer'|'schueler'|''} personaView
 * @returns {Array<'lehrer'|'schueler'>}
 */
export function portalShellViewOrder(personaView) {
    if (personaView === 'schueler') return ['schueler', 'lehrer'];
    if (personaView === 'lehrer') return ['lehrer', 'schueler'];
    return ['lehrer', 'schueler'];
}

/**
 * @param {{ loggedIn?: boolean, isIt?: boolean, isLehrer?: boolean, isSchueler?: boolean }|null|undefined} personas
 * @returns {'lehrer'|'schueler'|''}
 */
export function resolvePortalWorkspaceView(personas) {
    if (!personas || !personas.loggedIn) return '';
    if (personas.isIt) {
        const choice = readItShellChoice();
        if (choice === 'full') return '';
        if (choice === 'schueler') return 'schueler';
        return 'lehrer';
    }
    if (personas.isSchueler && !personas.isLehrer) return 'schueler';
    return 'lehrer';
}
