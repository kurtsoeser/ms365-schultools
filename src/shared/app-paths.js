/**
 * Korrekte hrefs für Tool-Seiten unabhängig vom aktuellen Pfad (Root vs. /tools/…).
 * @file
 */

/**
 * @param {string} toolFile z. B. "datenhygiene.html" oder "tools/datenhygiene.html"
 * @returns {string}
 */
export function resolveToolsHref(toolFile) {
    const file = String(toolFile || '')
        .replace(/^[./]+/, '')
        .replace(/^tools\//, '');
    if (!file) return './';
    try {
        const p = String(window.location.pathname || '/').split('?')[0].split('#')[0];
        const lower = p.toLowerCase();
        const toolsNeedle = '/tools/';
        const idx = lower.indexOf(toolsNeedle);
        if (idx !== -1) {
            return p.slice(0, idx + toolsNeedle.length) + file;
        }
        const dir = p.endsWith('/') ? p : p.slice(0, p.lastIndexOf('/') + 1);
        return dir + 'tools/' + file;
    } catch {
        return 'tools/' + file;
    }
}

/**
 * @param {ParentNode|Document} [root]
 */
export function fixToolsAnchors(root) {
    const host = root && typeof root.querySelectorAll === 'function' ? root : document;
    host.querySelectorAll('a[href^="tools/"]').forEach(function (a) {
        const rel = String(a.getAttribute('href') || '').replace(/^tools\//, '');
        if (!rel) return;
        a.setAttribute('href', resolveToolsHref(rel));
    });
}

const api = { resolveToolsHref: resolveToolsHref, fixToolsAnchors: fixToolsAnchors };

if (typeof window !== 'undefined') {
    window.ms365AppPaths = api;
}

export default api;
