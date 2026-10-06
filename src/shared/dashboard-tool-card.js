/**
 * Werkzeugkatalog: einheitliche .tool-card-Struktur (UI-Spec Abschnitt 4).
 */
import { toolLabel } from './dashboard-audience-catalog.js';
import { specToolDescription } from './dashboard-tool-spec-copy.js';

/** Wichtige Einstiegs-Werkzeuge → dunkle Primärkachel */
const CATALOG_PRIMARY_TOOLS = new Set([
    'slg-schueler',
    'slg-lehrer',
    'verwaltung',
    'jahrgang',
    'kursteams',
    'kursteam-einzeln',
    'personen-verwaltung',
    'schueler-lifecycle',
    'playbook-schuljahresstart',
    'organisations-assistent',
    'sharepoint-intranet-hub',
    'schularbeiten-planer',
    'freistellung-planer',
    'datenhygiene',
    'cleanup-playbook',
    'leere-gruppen-report'
]);

/**
 * @param {string} toolId
 * @param {string} [linkText]
 */
export function catalogToolActionLabel(toolId, linkText) {
    const id = String(toolId || '').trim();
    if (id === 'personen-verwaltung') return 'Suchen';
    if (id === 'namenskonvention-audit') return 'Prüfen';
    if (id === 'bildungsportal-stammdaten') return 'Roadmap';
    if (id.startsWith('playbook-')) return 'Playbook';
    const t = String(linkText || '').trim();
    if (/suchen/i.test(t)) return 'Suchen';
    if (/playbook/i.test(t)) return 'Playbook';
    if (/prüfen|pruefen/i.test(t)) return 'Prüfen';
    if (/roadmap/i.test(t)) return 'Roadmap';
    if (/abgleich/i.test(t)) return 'Abgleichen';
    if (/start/i.test(t)) return 'Start';
    return 'Öffnen';
}

/**
 * @param {HTMLElement|null} el
 */
function plainHeadingText(el) {
    if (!el) return '';
    const clone = el.cloneNode(true);
    clone.querySelectorAll('.context-tip, .card-info, .dashboard-fav-btn, .card-star').forEach(function (n) {
        n.remove();
    });
    return String(clone.textContent || '').replace(/\s+/g, ' ').trim();
}

/** @param {string} text @param {number} [max] */
export function truncateCatalogDescription(text, max) {
    const limit = max || 110;
    const s = String(text || '').replace(/\s+/g, ' ').trim();
    if (!s) return '';
    if (s.length <= limit) return s;
    const cut = s.slice(0, limit - 1);
    const lastSpace = cut.lastIndexOf(' ');
    const base = lastSpace > 40 ? cut.slice(0, lastSpace) : cut;
    return base.trim() + '…';
}

/**
 * @param {string} toolId
 * @param {string} fallback
 */
export function catalogCardDescription(toolId, fallback) {
    const id = String(toolId || '').trim();
    const fromSpec = specToolDescription(id);
    if (fromSpec) return truncateCatalogDescription(fromSpec, 120);
    try {
        const tips = window.ms365ContextTips && window.ms365ContextTips.TIPS;
        const tip = tips && id ? tips[id] : '';
        if (tip) {
            const first = String(tip).split(/(?<=[.!?])\s+/)[0] || tip;
            return truncateCatalogDescription(first, 120);
        }
    } catch {
        /* ignore */
    }
    return truncateCatalogDescription(fallback, 110);
}

/**
 * @param {HTMLElement} card
 */
function readLegacyCardParts(card) {
    const toolId = String(card.getAttribute('data-tool-id') || '').trim();
    const h2 = card.querySelector(':scope > h2, .card-title');
    const labeled = toolLabel(toolId);
    const title = labeled || plainHeadingText(h2) || toolId;
    const iconEl = card.querySelector('.card-icon i.bi, :scope > h2 i.bi, .card-title i.bi');
    const iconClass = iconEl ? iconEl.className : 'bi bi-box-arrow-up-right';
    const descEl = card.querySelector(':scope > p, .card-description');
    const rawDesc = descEl ? String(descEl.textContent || '').trim() : '';
    const description = catalogCardDescription(toolId, rawDesc);
    const link =
        card.querySelector('a.btn[href], a[href].btn, .card-actions a[href]') ||
        card.querySelector('a[href]');
    const href = link ? link.getAttribute('href') || '#' : '#';
    const linkText = link ? String(link.textContent || '').trim() : '';
    const comingSoon = card.classList.contains('tool-card--coming-soon');
    return { title, iconClass, description, href, linkText, comingSoon };
}

/**
 * @param {HTMLElement} card
 */
export function enhanceCatalogToolCard(card) {
    if (!card || card.dataset.toolCardEnhanced === '1') return;
    const toolId = String(card.getAttribute('data-tool-id') || '').trim();
    if (!toolId) return;

    const legacy = readLegacyCardParts(card);
    const existingStar = card.querySelector('.card-star, .dashboard-fav-btn');
    const existingStatus = card.querySelector('.card-status, .dash-choice-status');
    const statusHtml = existingStatus ? existingStatus.outerHTML : '';

    const isPrimary = CATALOG_PRIMARY_TOOLS.has(toolId);
    const tone = card.getAttribute('data-status-tone') || '';
    const recommended = tone === 'warn' || tone === 'mismatch' || tone === 'unmatched';
    const action = catalogToolActionLabel(toolId, legacy.linkText);

    card.classList.add('tool-card', isPrimary ? 'primary' : 'secondary');
    if (recommended) card.classList.add('recommended');
    if (legacy.comingSoon) card.classList.add('coming-soon');

    const starOuter = existingStar ? existingStar.outerHTML : '';
    const starBtn =
        starOuter ||
        '<button type="button" class="card-star dashboard-fav-btn" aria-pressed="false" aria-label="Zu Favoriten hinzufügen">☆</button>';

    let actionsHtml = '';
    if (!legacy.comingSoon) {
        actionsHtml =
            '<div class="card-actions">' +
            '<a class="card-btn primary" href="' +
            legacy.href +
            '">' +
            action +
            '</a></div>';
    } else {
        const badge = card.querySelector('.card-coming-soon-badge');
        actionsHtml = badge
            ? badge.outerHTML
            : '<div class="card-coming-soon-badge">In Entwicklung</div>';
    }

    let recommendedHint = '';
    if (recommended && !legacy.comingSoon) {
        recommendedHint =
            '<div class="card-recommended-hint" aria-hidden="true">⚠️ Behebt aktuelle Abweichung</div>';
    }

    const descBlock = legacy.description
        ? '<p class="card-description">' + legacy.description + '</p>'
        : '';

    card.innerHTML =
        starBtn +
        '<div class="card-header">' +
        '<span class="card-icon" aria-hidden="true"><i class="' +
        legacy.iconClass +
        '"></i></span>' +
        '<h3 class="card-title">' +
        legacy.title +
        '</h3>' +
        '</div>' +
        descBlock +
        (statusHtml || '<div class="card-status" hidden></div>') +
        recommendedHint +
        actionsHtml;

    const statusEl = card.querySelector('.card-status, .dash-choice-status');
    if (statusEl) {
        statusEl.classList.add('card-status');
        statusEl.classList.remove('dash-choice-status');
    }

    const star = card.querySelector('.card-star, .dashboard-fav-btn');
    if (star) {
        star.classList.add('card-star');
        if (!star.textContent || star.querySelector('.bi')) {
            /* bootstrap-icons variant from injectFavoriteButtons */
        } else if (!star.querySelector('.bi')) {
            star.textContent = '☆';
        }
    }

    card.dataset.toolCardEnhanced = '1';
}

/**
 * @param {ParentNode} root
 */
export function enhanceCatalogToolCards(root) {
    const catalog = root && root.querySelector ? root : document.getElementById('dash-catalog');
    if (!catalog) return;
    catalog.querySelectorAll('.choice[data-tool-id]').forEach(function (card) {
        enhanceCatalogToolCard(card);
    });
    try {
        if (window.ms365ContextTips && typeof window.ms365ContextTips.mount === 'function') {
            window.ms365ContextTips.mount(catalog);
        }
    } catch {
        /* ignore */
    }
}

/**
 * Karten nach Status-Update neu einlesen (Titel/Beschreibung bleiben kurz).
 * @param {ParentNode} root
 */
export function refreshCatalogToolCardCopy(root) {
    const catalog = root && root.querySelector ? root : document.getElementById('dash-catalog');
    if (!catalog) return;
    catalog.querySelectorAll('.choice.tool-card[data-tool-id]').forEach(function (card) {
        const toolId = String(card.getAttribute('data-tool-id') || '').trim();
        const titleEl = card.querySelector('.card-title');
        const descEl = card.querySelector('.card-description');
        if (titleEl) titleEl.textContent = toolLabel(toolId) || titleEl.textContent;
        if (descEl) {
            const next = catalogCardDescription(toolId, descEl.textContent);
            if (next) descEl.textContent = next;
            else descEl.remove();
        }
    });
}
