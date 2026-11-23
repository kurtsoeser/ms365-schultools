/**
 * IT-Arbeitsfläche: Nav-Baum (Bereiche + Untertools) für Dashboard-Leiste und Tool-Seiten.
 */
import { readItShellChoice } from './dashboard-layout-mode.js';
import { mergeStammdatenWorkspaceTools } from './stammdaten-nav-catalog.js';

export const IT_WORKSPACE_FLAG_KEY = 'ms365-it-workspace-v1';
export const IT_WORKSPACE_NAV_KEY = 'ms365-it-workspace-nav-v1';
export const IT_WORKSPACE_TASK_KEY = 'ms365-it-workspace-task-v1';
/** Stammdaten ist eigener Top-Level-Bereich – tenant.html gehört nur hierher. */
export const IT_WORKSPACE_STAMMDATEN_TASK_ID = 'dashTaskImportVerknuepfen';

/**
 * Fallback-Baum, wenn die Session noch keinen Nav-Stand vom Dashboard hat
 * (z. B. direkter Sprung zu Stammdaten / Portal vorbereiten).
 * @returns {{ id: string, label: string, icon: string, overviewHash: string, tools: { href: string, label: string, icon: string, expert?: boolean }[] }[]}
 */
export function defaultItWorkspaceTree() {
    return [
        {
            id: 'dashTaskImportVerknuepfen',
            label: 'Stammdaten',
            icon: 'bi bi-database',
            overviewHash: 'import',
            tools: mergeStammdatenWorkspaceTools()
        },
        {
            id: 'dashTaskGruppen',
            label: 'Mitgliedschaften',
            icon: 'bi bi-people',
            overviewHash: 'mitgliedschaften',
            tools: [
                { href: 'tools/lehrer-sammelgruppe.html', label: 'Lehrkräfte-Sammelgruppe', icon: 'bi bi-people' },
                { href: 'tools/schueler-sammelgruppe.html', label: 'Schüler-Sammelgruppe', icon: 'bi bi-mortarboard' }
            ]
        },
        {
            id: 'dashTaskUnterricht',
            label: 'Klassen & Unterricht',
            icon: 'bi bi-mortarboard',
            overviewHash: 'klassen',
            tools: [
                { href: 'tools/jahrgangsgruppen.html', label: 'Klassengruppen', icon: 'bi bi-collection' },
                { href: 'tools/kursteams.html', label: 'Unterrichtsteams', icon: 'bi bi-easel' }
            ]
        },
        {
            id: 'dashTaskPersonen',
            label: 'Personen & Gäste',
            icon: 'bi bi-person-badge',
            overviewHash: 'personen',
            tools: [{ href: 'tools/gaeste-verwalten.html', label: 'Gäste verwalten', icon: 'bi bi-person-plus' }]
        },
        {
            id: 'dashTaskSchuljahr',
            label: 'Schuljahresstart',
            icon: 'bi bi-calendar2-range',
            overviewHash: 'schuljahresstart',
            tools: [
                // Primärer Jahres-Einstieg = Schuljahresstart-Playbook (Org-Assistent nur unterstützend).
                {
                    href: 'tools/playbook-schuljahresstart.html',
                    label: 'Schuljahresstart',
                    icon: 'bi bi-rocket-takeoff'
                }
            ]
        },
        {
            id: 'dashTaskIntranet',
            label: 'Intranet',
            icon: 'bi bi-house-door',
            overviewHash: 'intranet',
            tools: [
                { href: 'tools/sharepoint-intranet-hub.html', label: 'Schul-Intranet', icon: 'bi bi-house-door' }
            ]
        },
        {
            id: 'dashTaskSchulApps',
            label: 'Schul-Apps',
            icon: 'bi bi-window-stack',
            overviewHash: 'schulapps',
            tools: [
                {
                    href: 'dashboard-werkzeug-zugriff.html',
                    label: 'Portal vorbereiten',
                    icon: 'bi bi-sliders2-vertical'
                },
                { href: 'tools/schularbeiten-planer.html', label: 'Schularbeiten-Planer', icon: 'bi bi-journal-check' },
                { href: 'tools/freistellung-planer.html', label: 'Freistellungen', icon: 'bi bi-calendar2-check' }
            ]
        },
        {
            id: 'dashTaskOrdnung',
            label: 'Aufräumen',
            icon: 'bi bi-ui-checks-grid',
            overviewHash: 'aufraeumen',
            tools: [{ href: 'tools/datenhygiene.html', label: 'Datenhygiene', icon: 'bi bi-shield-check' }]
        }
    ];
}

/**
 * Prüft auf dem Dashboard (`index.html`), ob der echte IT-Bereich aktiv ist
 * (`data-dash-it-shell` + gespeicherte IT-Shell-Wahl `full`).
 * @returns {boolean}
 */
export function isDashboardItShellActive() {
    if (typeof document === 'undefined' || !document.documentElement) return false;
    if (!document.documentElement.hasAttribute('data-dash-it-shell')) return false;
    try {
        return readItShellChoice() === 'full';
    } catch {
        return false;
    }
}

/**
 * Stellt sicher, dass Nav-Baum + Flag für IT-Seiten vorhanden sind.
 * @returns {{ version: number, tree: any[] }|null}
 */
export function ensureItWorkspaceNavReady() {
    let stored = loadItWorkspaceNav();
    if (!stored || !stored.tree || !stored.tree.length) {
        const tree = defaultItWorkspaceTree();
        persistItWorkspaceNav(tree, '');
        stored = { version: 1, tree: tree };
    } else {
        try {
            sessionStorage.setItem(IT_WORKSPACE_FLAG_KEY, '1');
        } catch {
            /* ignore */
        }
    }
    return stored;
}

/**
 * @param {string} href
 * @returns {string}
 */
const ROOT_WORKSPACE_PAGES_RE = /^(tenant|index|dashboard-werkzeug-zugriff)\.html$/i;

export function normalizeWorkspaceHref(href) {
    const raw = String(href || '').trim();
    if (!raw || raw === '#') return '';
    try {
        const base =
            typeof location !== 'undefined' && location.href
                ? location.href
                : 'https://ms365.schule/';
        const u = new URL(raw, base);
        let path = String(u.pathname || '').replace(/\\/g, '/');
        const leaf = path.split('/').filter(Boolean).pop() || '';
        // Bekannte Root-Seiten nie unter tools/ führen (auch bei Relativ-Auflösung aus /tools/).
        if (ROOT_WORKSPACE_PAGES_RE.test(leaf)) {
            path = leaf;
        } else {
            const toolsIdx = path.toLowerCase().lastIndexOf('/tools/');
            if (toolsIdx >= 0) path = path.slice(toolsIdx + 1);
            else if (/\.html$/i.test(leaf)) path = 'tools/' + leaf;
            else path = leaf;
        }
        return path + (u.search || '') + (u.hash || '');
    } catch {
        return raw.replace(/^\.\.\//, '').replace(/^\.\//, '');
    }
}

/**
 * @param {string} a
 * @param {string} b
 * @param {{ looseFile?: boolean }} [opts]
 */
export function workspaceHrefsMatch(a, b, opts) {
    const left = normalizeWorkspaceHref(a);
    const right = normalizeWorkspaceHref(b);
    if (!left || !right) return false;
    if (left === right) return true;
    const stripHash = function (s) {
        return s.replace(/#.*$/, '');
    };
    const hashOf = function (s) {
        const i = String(s).indexOf('#');
        return i >= 0 ? String(s).slice(i + 1).toLowerCase() : '';
    };
    if (stripHash(left) === stripHash(right)) {
        const hL = hashOf(left);
        const hR = hashOf(right);
        // Beide mit Hash (z. B. #faecher vs #klassen): Exact nur bei gleichem Tab.
        if (hL && hR) {
            if (hL === hR) return true;
            return !!(opts && opts.looseFile);
        }
        // Nur eine Seite gehasht: Exact nein, looseFile ja.
        if (hL || hR) return !!(opts && opts.looseFile);
        return true;
    }
    if (opts && opts.looseFile) {
        const fileOf = function (s) {
            return stripHash(s).split('?')[0].toLowerCase();
        };
        return fileOf(left) === fileOf(right);
    }
    return false;
}

/**
 * @param {HTMLElement} task
 * @returns {{ href: string, label: string, icon: string, expert: boolean }[]}
 */
export function extractToolsFromTask(task) {
    if (!task) return [];
    const seenHref = Object.create(null);
    const seenFileLabel = Object.create(null);
    const out = [];
    task.querySelectorAll('a.dash-task-row[href]').forEach(function (a) {
        const href = String(a.getAttribute('href') || '').trim();
        const key = normalizeWorkspaceHref(href);
        if (!key || seenHref[key]) return;
        const titleEl = a.querySelector('.dash-task-row__title');
        let label = titleEl
            ? String(titleEl.textContent || '')
                  .replace(/\s+/g, ' ')
                  .trim()
            : '';
        if (!label) {
            const labelEl = a.querySelector('.dash-task-row__label');
            if (labelEl) {
                const clone = labelEl.cloneNode(true);
                clone.querySelectorAll('i, .bi, .dash-task-row__desc, .dash-task-row__action, .dash-task-row__status').forEach(
                    function (n) {
                        n.remove();
                    }
                );
                label = String(clone.textContent || '')
                    .replace(/\s+/g, ' ')
                    .trim();
            }
        }
        if (!label) {
            label = String(a.getAttribute('aria-label') || a.title || key)
                .replace(/\s+/g, ' ')
                .trim();
        }
        const finalLabel = label || key;
        const baseFile = key.replace(/#.*$/, '').split('?')[0].toLowerCase();
        const fileLabelKey = baseFile + '|' + finalLabel.toLowerCase();
        if (seenFileLabel[fileLabelKey]) return;

        seenHref[key] = true;
        seenFileLabel[fileLabelKey] = true;
        const iconEl = a.querySelector('.dash-task-row__label > i, .dash-task-row__head i, i.bi');
        const expert = !!(a.closest && a.closest('details.dash-task-expert'));
        out.push({
            href: href,
            label: finalLabel,
            icon: iconEl && iconEl.className ? String(iconEl.className) : 'bi bi-box-arrow-up-right',
            expert: expert
        });
    });
    return out;
}

/**
 * tenant.html nur unter Stammdaten führen – nicht unter Schuljahresstart o. ä.
 * @param {string} href
 * @returns {boolean}
 */
export function isTenantWorkspaceHref(href) {
    const key = normalizeWorkspaceHref(href);
    if (!key) return false;
    const file = key.replace(/#.*$/, '').split('?')[0].toLowerCase();
    return file === 'tenant.html';
}

/**
 * Stammdaten-Node: Register-Bereiche als Primär-Untermenü erzwingen.
 * @param {any} node
 * @returns {any}
 */
export function ensureStammdatenNodeTools(node) {
    if (!node || typeof node !== 'object') return node;
    if (String(node.id || '') !== IT_WORKSPACE_STAMMDATEN_TASK_ID) return node;
    const extracted = Array.isArray(node.tools) ? node.tools : [];
    return Object.assign({}, node, {
        label: node.label || 'Stammdaten',
        icon: node.icon || 'bi bi-database',
        tools: mergeStammdatenWorkspaceTools(extracted)
    });
}

/**
 * Stammdaten bleibt eigene Gruppe: tenant.html-Links aus anderen Bereichen entfernen
 * und Stammdaten-Untermenü auf Register-Bereiche normieren.
 * @param {any[]} tree
 * @returns {any[]}
 */
export function sanitizeItWorkspaceTree(tree) {
    if (!Array.isArray(tree)) return [];
    return tree.map(function (node) {
        if (!node || typeof node !== 'object') return node;
        if (String(node.id || '') === IT_WORKSPACE_STAMMDATEN_TASK_ID) {
            return ensureStammdatenNodeTools(node);
        }
        const tools = Array.isArray(node.tools) ? node.tools : [];
        const nextTools = tools.filter(function (t) {
            return !isTenantWorkspaceHref(t && t.href);
        });
        if (nextTools.length === tools.length) return node;
        return Object.assign({}, node, { tools: nextTools });
    });
}

/**
 * @param {{ id: string, label: string, icon?: string, overviewHash?: string, tools: { href: string, label: string, icon: string }[] }[]} tree
 */
export function persistItWorkspaceNav(tree, activeTaskId) {
    try {
        const clean = sanitizeItWorkspaceTree(tree || []);
        sessionStorage.setItem(IT_WORKSPACE_FLAG_KEY, '1');
        sessionStorage.setItem(IT_WORKSPACE_NAV_KEY, JSON.stringify({ version: 1, tree: clean }));
        if (activeTaskId) sessionStorage.setItem(IT_WORKSPACE_TASK_KEY, String(activeTaskId));
    } catch {
        /* ignore */
    }
}

/** Deaktiviert die bleibende IT-Leiste (z. B. beim Wechsel auf Portal/Schüler-Ansicht). */
export function deactivateItWorkspace() {
    try {
        sessionStorage.removeItem(IT_WORKSPACE_FLAG_KEY);
        sessionStorage.removeItem(IT_WORKSPACE_TASK_KEY);
    } catch {
        /* ignore */
    }
}

/** Löscht den gesamten gespeicherten IT-Arbeitsflächen-Zustand (z. B. für Nicht-IT-Konten). */
export function clearItWorkspaceNav() {
    try {
        sessionStorage.removeItem(IT_WORKSPACE_FLAG_KEY);
        sessionStorage.removeItem(IT_WORKSPACE_NAV_KEY);
        sessionStorage.removeItem(IT_WORKSPACE_TASK_KEY);
    } catch {
        /* ignore */
    }
}

/** @returns {boolean} */
export function isItWorkspaceActive() {
    try {
        if (sessionStorage.getItem(IT_WORKSPACE_FLAG_KEY) === '1') return true;
    } catch {
        /* ignore */
    }
    if (typeof document === 'undefined') return false;
    return document.body.classList.contains('it-workspace-shell');
}

/**
 * @returns {{ version: number, tree: any[] }|null}
 */
export function loadItWorkspaceNav() {
    try {
        const raw = sessionStorage.getItem(IT_WORKSPACE_NAV_KEY);
        if (!raw) return null;
        const data = JSON.parse(raw);
        if (!data || !Array.isArray(data.tree)) return null;
        data.tree = sanitizeItWorkspaceTree(data.tree);
        return data;
    } catch {
        return null;
    }
}

/** @returns {string} */
export function loadItWorkspaceTaskId() {
    try {
        return String(sessionStorage.getItem(IT_WORKSPACE_TASK_KEY) || '');
    } catch {
        return '';
    }
}

/**
 * @param {{ btn: HTMLElement, task: HTMLElement }[]} navButtons
 * @param {(task: HTMLElement) => string} labelFn
 * @param {(task: HTMLElement) => string} iconFn
 */
export function buildNavTreeFromDashboard(navButtons, labelFn, iconFn) {
    return sanitizeItWorkspaceTree(
        (navButtons || []).map(function (item) {
            const task = item.task;
            const extracted = extractToolsFromTask(task);
            const isStammdaten = String(task.id || '') === IT_WORKSPACE_STAMMDATEN_TASK_ID;
            return {
                id: task.id,
                label: labelFn(task),
                icon: iconFn(task),
                overviewHash: '',
                tools: isStammdaten ? mergeStammdatenWorkspaceTools(extracted) : extracted
            };
        })
    );
}

/**
 * Relative hrefs von Tool-Seiten aus korrigieren (tools/… → ../tools/… bzw. tenant).
 * @param {string} href
 * @param {boolean} fromToolsFolder
 */
export function resolveWorkspaceHrefForPage(href, fromToolsFolder) {
    const raw = String(href || '').trim();
    if (!raw) return raw;
    if (/^(https?:|mailto:|javascript:)/i.test(raw)) return raw;
    if (!fromToolsFolder) return raw;
    if (raw.startsWith('../') || raw.startsWith('/')) return raw;
    if (/^tools\//i.test(raw)) return '../' + raw;
    if (/^tenant\.html/i.test(raw) || /^index\.html/i.test(raw) || /^dashboard-werkzeug-zugriff\.html/i.test(raw)) {
        return '../' + raw;
    }
    return raw;
}
