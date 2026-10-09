/**
 * Stammdaten: Entra-/Listen-Status als Badge an den Haupt-Tabs.
 */

const TAB_STATUS_CLASS = ['tab-btn--st-ok', 'tab-btn--st-warn', 'tab-btn--st-error', 'tab-btn--st-muted'];

/**
 * @param {{ total: number, matched: number, notFound: number, unchecked: number }} stat
 * @returns {{ kind: string, icon: string }}
 */
export function personMatchTabStatus(stat) {
    const total = stat && stat.total ? stat.total : 0;
    const matched = stat && stat.matched ? stat.matched : 0;
    const notFound = stat && stat.notFound ? stat.notFound : 0;
    const unchecked = stat && stat.unchecked ? stat.unchecked : 0;
    if (total === 0) return { kind: 'muted', icon: 'bi-dash' };
    if (matched === total) return { kind: 'ok', icon: 'bi-check-circle-fill' };
    if (unchecked > 0) return { kind: 'warn', icon: 'bi-question-circle' };
    if (notFound > 0) return { kind: 'error', icon: 'bi-x-circle' };
    return { kind: 'ok', icon: 'bi-check-circle-fill' };
}

/**
 * @returns {{ kind: string, icon: string }}
 */
export function classMatchTabStatus(clTotal, clMatched, clNotFound, clUnchecked) {
    if (!clTotal) return { kind: 'muted', icon: 'bi-dash' };
    if (clMatched === clTotal) return { kind: 'ok', icon: 'bi-check-circle-fill' };
    if (clUnchecked > 0) return { kind: 'warn', icon: 'bi-question-circle' };
    if (clNotFound > 0) return { kind: 'error', icon: 'bi-x-circle' };
    return { kind: 'ok', icon: 'bi-check-circle-fill' };
}

/**
 * @returns {{ kind: string, icon: string }}
 */
export function catalogLinkTabStatus(total, linked) {
    if (!total) return { kind: 'muted', icon: 'bi-dash' };
    if (linked === total) return { kind: 'ok', icon: 'bi-check-circle-fill' };
    if (linked > 0) return { kind: 'warn', icon: 'bi-info-circle' };
    return { kind: 'muted', icon: 'bi-info-circle' };
}

/**
 * @param {Record<string, { kind: string, icon: string, title?: string }>} map
 */
export function applyTabRegisterStatus(map) {
    if (!map || typeof map !== 'object') return;
    Object.keys(map).forEach(function (tabId) {
        const btn = document.getElementById(tabId);
        if (!btn) return;
        const st = map[tabId];
        if (!st) return;
        TAB_STATUS_CLASS.forEach(function (c) {
            btn.classList.remove(c);
        });
        btn.classList.add('tab-btn--st-' + (st.kind || 'muted'));
        let badge = btn.querySelector('.tab-btn__st');
        if (!badge) {
            badge = document.createElement('span');
            badge.className = 'tab-btn__st';
            badge.setAttribute('aria-hidden', 'true');
            btn.appendChild(badge);
        }
        const kind = st.kind || 'muted';
        const icon = st.icon || 'bi-dash';
        const isEmpty = kind === 'muted' && icon === 'bi-dash';
        badge.className = 'tab-btn__st tab-btn__st--' + kind + (isEmpty ? ' tab-btn__st--empty' : '');
        badge.innerHTML = '';
        if (st.title) {
            const base = btn.getAttribute('data-tab-title-base');
            const label = base || btn.textContent.replace(/\s+/g, ' ').trim();
            if (!base) btn.setAttribute('data-tab-title-base', label);
            btn.title = st.title;
            btn.setAttribute('aria-label', label + ' – ' + st.title);
        }
    });
}

const REASON_LABELS = {
    autosave: 'Auto-gespeichert',
    'manual-save': 'Gespeichert',
    writeback: 'Übernommen',
    render: 'Geladen'
};

let lastSaved = { reason: 'render', at: Date.now() };

/**
 * @param {string} [reason]
 */
export function updateRegisterLastSavedMeta(reason) {
    lastSaved = {
        reason: String(reason || 'render').trim() || 'render',
        at: Date.now()
    };
}

/**
 * @returns {string} z. B. „ — Geladen um 10:58“
 */
export function formatRegisterLastSavedSuffix() {
    const time = new Date(lastSaved.at).toLocaleTimeString('de-AT', { hour: '2-digit', minute: '2-digit' });
    const prefix = REASON_LABELS[lastSaved.reason] || 'Aktualisiert';
    return ' — ' + prefix + ' um ' + time;
}
