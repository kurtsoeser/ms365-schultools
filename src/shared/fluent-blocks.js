/**
 * Welle 2: gemeinsame Interaktions-Bausteine ohne eigene Fachlogik.
 * Stellt die leise Meldung mit optionalem „Rückgängig“ bereit.
 */

let activeToastTimer = null;

export function dismissFluentToast(doc) {
    const d = doc || (typeof document !== 'undefined' ? document : null);
    if (!d) return;
    if (activeToastTimer) {
        clearTimeout(activeToastTimer);
        activeToastTimer = null;
    }
    const existing = d.getElementById('ms365FluentToast');
    if (existing) existing.remove();
}

/**
 * @param {{
 *   message: string,
 *   actionLabel?: string,
 *   onAction?: () => void,
 *   durationMs?: number,
 *   doc?: Document
 * }} opts
 */
export function showFluentToast(opts) {
    const o = opts || {};
    const d = o.doc || (typeof document !== 'undefined' ? document : null);
    const text = String(o.message || '').trim();
    if (!d || !d.body || !text) return null;

    dismissFluentToast(d);

    const toast = d.createElement('div');
    toast.id = 'ms365FluentToast';
    toast.className = 'fb-toast';
    toast.setAttribute('role', 'status');
    toast.setAttribute('aria-live', 'polite');

    const span = d.createElement('span');
    span.className = 'fb-toast__text';
    span.textContent = text;
    toast.appendChild(span);

    if (o.actionLabel && typeof o.onAction === 'function') {
        const btn = d.createElement('button');
        btn.type = 'button';
        btn.className = 'fb-toast__action';
        btn.textContent = String(o.actionLabel);
        btn.addEventListener('click', function () {
            dismissFluentToast(d);
            o.onAction();
        });
        toast.appendChild(btn);
    }

    d.body.appendChild(toast);

    const ms = Number.isFinite(o.durationMs) ? Math.max(1500, Number(o.durationMs)) : 5500;
    activeToastTimer = setTimeout(function () {
        dismissFluentToast(d);
    }, ms);

    return toast;
}
