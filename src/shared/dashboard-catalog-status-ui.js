/**
 * Status-Zeile auf Werkzeug-Kacheln im Katalog (Alle Werkzeuge).
 */
import { resolveDashboardToolStatus } from './dashboard-tool-status-logic.js?v=kc-dash-6';
import { applySchulappsUsableAppearance } from './dashboard-tool-card.js';

/**
 * @param {string} tone
 */
function statusToneClass(tone) {
    const t = String(tone || '').trim();
    if (t === 'ok') return 'success';
    if (t === 'warn' || t === 'mismatch') return 'warning';
    if (t === 'error') return 'error';
    return 'warning';
}

/**
 * @param {string} tone
 */
function statusRank(tone) {
    const t = String(tone || '').trim();
    if (t === 'warn' || t === 'mismatch' || t === 'unmatched' || t === 'pending' || t === 'error') {
        return 0;
    }
    if (!t) return 1;
    if (t === 'ok') return 2;
    return 1;
}

/**
 * Offene / gelbe Karten in Einrichten-Grids nach oben, ruhige OK-Karten nach unten.
 * @param {ParentNode} root
 */
export function prioritizeOpenCatalogCards(root) {
    if (!root) return;
    root.querySelectorAll('[data-cluster-grid]').forEach(function (grid) {
        const cards = Array.from(grid.querySelectorAll(':scope > .choice[data-tool-id]'));
        if (cards.length < 2) return;
        cards
            .slice()
            .sort(function (a, b) {
                const ra = statusRank(a.getAttribute('data-status-tone'));
                const rb = statusRank(b.getAttribute('data-status-tone'));
                if (ra !== rb) return ra - rb;
                return cards.indexOf(a) - cards.indexOf(b);
            })
            .forEach(function (card) {
                grid.appendChild(card);
            });
    });
}

/**
 * @param {HTMLElement} card
 * @param {{ tone?: string, primary?: string, secondary?: string, variant?: string }|null} info
 */
export function paintCatalogToolStatus(card, info) {
    if (!card) return;
    const existing = card.querySelector('.card-status, .dash-choice-status');
    if (!info || !info.primary) {
        if (existing) {
            existing.hidden = true;
            existing.textContent = '';
            existing.removeAttribute('data-variant');
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
    const variant = String(info.variant || '').trim();
    card.setAttribute('data-status-tone', tone);
    wrap.hidden = false;
    wrap.className =
        'card-status ' +
        statusToneClass(tone) +
        (variant === 'subtle' ? ' card-status--subtle' : '') +
        (variant === 'emphasis' ? ' card-status--emphasis' : '');
    wrap.setAttribute('data-tone', tone);
    if (variant) wrap.setAttribute('data-variant', variant);
    else wrap.removeAttribute('data-variant');

    if (variant === 'subtle' && tone === 'ok') {
        wrap.innerHTML =
            '<i class="bi bi-check2 card-status__check" aria-hidden="true"></i>' +
            '<span class="card-status__label">' +
            String(info.primary) +
            '</span>';
    } else {
        const parts = [info.primary];
        if (info.secondary) parts.push(info.secondary);
        wrap.innerHTML =
            '<span class="status-dot-sm" aria-hidden="true"></span>' + parts.join(' · ');
    }

    if (tone === 'warn' || tone === 'mismatch' || tone === 'unmatched' || tone === 'pending') {
        card.classList.add('recommended');
        let hint = card.querySelector('.card-recommended-hint');
        if (!hint && !card.classList.contains('coming-soon') && tone !== 'pending') {
            hint = document.createElement('div');
            hint.className = 'card-recommended-hint';
            hint.textContent = '⚠️ Behebt aktuelle Abweichung';
            const actions = card.querySelector('.card-actions');
            if (actions) card.insertBefore(hint, actions);
            else card.appendChild(hint);
        }
    } else {
        card.classList.remove('recommended');
        const hint = card.querySelector('.card-recommended-hint');
        if (hint) hint.remove();
    }
}

function intranetConnected(container) {
    const setup = container && container.setup ? container.setup : {};
    const url =
        (setup.intranetSiteUrl && String(setup.intranetSiteUrl).trim()) ||
        (setup.schoolIntranetSiteUrl && String(setup.schoolIntranetSiteUrl).trim()) ||
        '';
    return !!url;
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
    const intranetOk = intranetConnected(opts.container);

    root.querySelectorAll('.choice[data-tool-id]').forEach(function (card) {
        const toolId = card.getAttribute('data-tool-id');
        const info = resolveDashboardToolStatus(toolId, ctx);
        paintCatalogToolStatus(card, info);
        applySchulappsUsableAppearance(card, { intranetOk: intranetOk });
    });

    if (opts.show) {
        prioritizeOpenCatalogCards(root);
    }
}

if (typeof window !== 'undefined') {
    window.ms365DashboardCatalogStatus = {
        apply: applyDashboardCatalogToolStatuses,
        paint: paintCatalogToolStatus,
        prioritizeOpen: prioritizeOpenCatalogCards
    };
}
