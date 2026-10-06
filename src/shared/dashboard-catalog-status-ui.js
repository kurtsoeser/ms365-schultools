/**
 * Status-Zeile auf Werkzeug-Kacheln im Katalog (Alle Werkzeuge).
 */
import { resolveDashboardToolStatus } from './dashboard-tool-status-logic.js';

/**
 * @param {HTMLElement} card
 * @param {{ tone?: string, primary?: string, secondary?: string }|null} info
 */
function statusToneClass(tone) {
    const t = String(tone || '').trim();
    if (t === 'ok') return 'success';
    if (t === 'warn' || t === 'mismatch') return 'warning';
    if (t === 'error') return 'error';
    return 'warning';
}

export function paintCatalogToolStatus(card, info) {
    if (!card) return;
    const existing = card.querySelector('.card-status, .dash-choice-status');
    if (!info || !info.primary) {
        if (existing) {
            existing.hidden = true;
            existing.textContent = '';
        }
        card.removeAttribute('data-status-tone');
        card.classList.remove('recommended');
        return;
    }
    let wrap = existing;
    if (!wrap) {
        wrap = document.createElement('div');
        wrap.className = 'card-status';
        const actions = card.querySelector('.card-actions');
        if (actions) card.insertBefore(wrap, actions);
        else card.appendChild(wrap);
    }
    const tone = info.tone || 'pending';
    card.setAttribute('data-status-tone', tone);
    wrap.hidden = false;
    wrap.className = 'card-status ' + statusToneClass(tone);
    wrap.setAttribute('data-tone', tone);

    const parts = [info.primary];
    if (info.secondary) parts.push(info.secondary);
    wrap.innerHTML =
        '<span class="status-dot-sm" aria-hidden="true"></span>' + parts.join(' · ');

    if (tone === 'warn' || tone === 'mismatch' || tone === 'unmatched') {
        card.classList.add('recommended');
        let hint = card.querySelector('.card-recommended-hint');
        if (!hint && !card.classList.contains('coming-soon')) {
            hint = document.createElement('div');
            hint.className = 'card-recommended-hint';
            hint.textContent = '⚠️ Behebt aktuelle Abweichung';
            const actions = card.querySelector('.card-actions');
            if (actions) card.insertBefore(hint, actions);
            else card.appendChild(hint);
        }
    }
}

/**
 * @param {{
 *   container: object|null,
 *   settings: object|null,
 *   hygieneById: Record<string, string>|null,
 *   show: boolean,
 *   root?: ParentNode
 * }} opts
 */
export function applyDashboardCatalogToolStatuses(opts) {
    const root = opts.root || document.getElementById('dash-catalog');
    if (!root) return;
    const hygieneApi =
        typeof window !== 'undefined' ? window.ms365MembershipHygiene : null;
    const ctx = {
        container: opts.container,
        settings: opts.settings,
        hygieneById: opts.hygieneById,
        hygieneApi: hygieneApi,
        show: !!opts.show
    };

    root.querySelectorAll('.choice[data-tool-id]').forEach(function (card) {
        const toolId = card.getAttribute('data-tool-id');
        const info = resolveDashboardToolStatus(toolId, ctx);
        paintCatalogToolStatus(card, info);
    });
}

if (typeof window !== 'undefined') {
    window.ms365DashboardCatalogStatus = {
        apply: applyDashboardCatalogToolStatuses
    };
}
