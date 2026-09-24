/**
 * Minimal-DOM-Helpers für Tool-UIs.
 */

/** Kurzform für `document.getElementById`. */
export function getEl(id) {
    if (typeof document === 'undefined') return null;
    return document.getElementById(id);
}

/**
 * Toast-Helper. Preferiert die zentrale `ms365ShowToast`-API (Icons + Varianten).
 * Fallback: Element `#toast` oder `ms365ToastOrAlert`.
 *
 * @param {string} msg
 * @param {object} [opts]
 * @param {number} [opts.durationMs=3500]
 * @param {string} [opts.elementId='toast']
 * @param {'info'|'success'|'error'|'warning'} [opts.kind]
 * @param {string} [opts.title]
 */
export function showToast(msg, opts = {}) {
    if (typeof window !== 'undefined' && typeof window.ms365ShowToast === 'function') {
        window.ms365ShowToast(msg, opts);
        return;
    }
    if (typeof document === 'undefined') return;
    const { durationMs = 3500, elementId = 'toast' } = opts;
    const el = document.getElementById(elementId);
    if (!el) {
        if (typeof window !== 'undefined' && typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(msg, opts);
        }
        return;
    }
    el.textContent = String(msg ?? '');
    el.classList.add('show');
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => el.classList.remove('show'), durationMs);
}
