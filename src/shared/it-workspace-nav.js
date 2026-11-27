/**
 * IT-Arbeitsfläche: Nav-Baum (Bereiche + Untertools) für Dashboard-Leiste und Tool-Seiten.
 */
import { readItShellChoice } from './dashboard-layout-mode.js';
import { itNavCatalog } from './it-nav-catalog.js';
import { mergeStammdatenWorkspaceTools } from './stammdaten-nav-catalog.js';

export const IT_WORKSPACE_FLAG_KEY = 'ms365-it-workspace-v1';
export const IT_WORKSPACE_NAV_KEY = 'ms365-it-workspace-nav-v1';
export const IT_WORKSPACE_TASK_KEY = 'ms365-it-workspace-task-v1';
/** Legacy-Key (früher sessionStorage); Embed gilt nur noch bei echtem iframe. */
export const IT_WORKSPACE_EMBED_KEY = 'ms365-it-embed-v1';
/** Stammdaten ist eigener Top-Level-Bereich – tenant.html gehört nur hierher. */
export const IT_WORKSPACE_STAMMDATEN_TASK_ID = 'dashTaskImportVerknuepfen';

/** True nur im iframe der IT-Shell – nicht top-level (Refresh behält normale Shell). */
export function isItWorkspaceEmbedMode() {
    try {
        return typeof window !== 'undefined' && window.self !== window.top;
    } catch {
        return true;
    }
}

/** @param {boolean} [_on] Legacy no-op – Erkennung läuft über window.self !== top. */
export function setItWorkspaceEmbedMode(_on) {
    try {
        sessionStorage.removeItem(IT_WORKSPACE_EMBED_KEY);
    } catch {
        /* ignore */
    }
}

/**
 * Kanonischer IT-Menübaum. Kommt aus it-nav-catalog.js, nie aus dem HTML.
 * @returns {{ id: string, label: string, icon: string, overviewHash: string, tools: { href: string, label: string, icon: string, expert?: boolean }[] }[]}
 */
export function defaultItWorkspaceTree() {
    return itNavCatalog();
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
 * Schreibt den Katalog in die Session. Eine ältere HTML-Kopie wird ersetzt,
 * damit Unterrichtsteams von überall dieselben Einträge hat.
 * @returns {{ version: number, tree: any[] }|null}
 */
export function ensureItWorkspaceNavReady() {
    let active = '';
    try {
        active = sessionStorage.getItem(IT_WORKSPACE_TASK_KEY) || '';
    } catch {
        active = '';
    }
    const tree = defaultItWorkspaceTree();
    persistItWorkspaceNav(tree, active);
    return { version: 1, tree: tree };
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
/**
 * Legacy-Redirect-Stubs auf die kanonische SLG-URL legen (Hash/Query sonst unverändert).
 * @param {string} href
 * @returns {string}
 */
export function canonicalizeWorkspaceHref(href) {
    const n = normalizeWorkspaceHref(href);
    if (!n) return '';
    const stripHash = n.replace(/#.*$/, '');
    const lower = stripHash.toLowerCase();
    if (lower === 'tools/schueler-sammelgruppe.html') {
        return 'tools/schueler-lehrer-gruppen.html?kind=schueler';
    }
    if (lower === 'tools/lehrer-sammelgruppe.html') {
        return 'tools/schueler-lehrer-gruppen.html?kind=lehrer';
    }
    return n;
}

export function workspaceHrefsMatch(a, b, opts) {
    const left = normalizeWorkspaceHref(a);
    const right = normalizeWorkspaceHref(b);
    if (!left || !right) return false;
    if (left === right) return true;
    const leftCanon = canonicalizeWorkspaceHref(left);
    const rightCanon = canonicalizeWorkspaceHref(right);
    // Nur wenn ein Alias greift (Stub ↔ kanonisch) – Hashes an tenant.html nicht kollabieren.
    if ((leftCanon !== left || rightCanon !== right) && leftCanon === rightCanon) return true;
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
            const canon = canonicalizeWorkspaceHref(s);
            return stripHash(canon).split('?')[0].toLowerCase();
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
 * Merkt die IT-Arbeitsfläche und den aktiven Bereich.
 * Den Menübaum speichert das nicht mehr: gezeichnet wird immer der Katalog.
 * Eine ältere Session-Kopie wird dabei entfernt.
 * @param {unknown} [_tree]
 * @param {string} [activeTaskId]
 */
export function persistItWorkspaceNav(_tree, activeTaskId) {
    try {
        sessionStorage.setItem(IT_WORKSPACE_FLAG_KEY, '1');
        sessionStorage.removeItem(IT_WORKSPACE_NAV_KEY);
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
 * Immer der Katalog. sessionStorage wird nicht mehr gelesen;
 * eine liegen gebliebene Kopie wird verworfen.
 * @returns {{ version: number, tree: any[] }}
 */
export function loadItWorkspaceNav() {
    try {
        sessionStorage.removeItem(IT_WORKSPACE_NAV_KEY);
    } catch {
        /* ignore */
    }
    return { version: 1, tree: sanitizeItWorkspaceTree(defaultItWorkspaceTree()) };
}

/** @returns {string} */
export function loadItWorkspaceTaskId() {
    try {
        return String(sessionStorage.getItem(IT_WORKSPACE_TASK_KEY) || '');
    } catch {
        return '';
    }
}

/** Menü kommt aus dem Katalog, nicht aus den Werkzeugzeilen der Startseite. */
export function buildNavTreeFromDashboard() {
    return defaultItWorkspaceTree();
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
