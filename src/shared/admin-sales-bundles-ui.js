/**
 * Admin-UI: Verkaufsbündel-Karten (Referenz + Schnellwahl Schildetail).
 */
import {
    SALES_ADDONS,
    SALES_BUNDLES,
    matchSalesBundle
} from './sales-bundles-catalog.js';
import {
    LICENSE_PACKAGES,
    packageLabel,
    packagePriceHint,
    sumPackageListPrices
} from './license-feature-catalog.js';

function escapeHtml(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * @param {HTMLElement} host
 * @param {{
 *   mode?: 'reference'|'apply',
 *   activePackageIds?: string[],
 *   onApply?: (bundleId: string) => void
 * }} [opts]
 */
export function renderSalesBundlesPanel(host, opts) {
    if (!host) return;
    const mode = opts && opts.mode === 'apply' ? 'apply' : 'reference';
    const active = (opts && opts.activePackageIds) || [];
    const matched = matchSalesBundle(active);
    const onApply = opts && typeof opts.onApply === 'function' ? opts.onApply : null;

    const cards = SALES_BUNDLES.map(function (b) {
        const listSum = sumPackageListPrices(b.packageIds);
        const save = Math.max(0, listSum - (b.priceEuro || 0));
        const pkgs = (b.packageIds || [])
            .map(function (pid) {
                const hint = packagePriceHint(pid);
                return (
                    '<span class="admin-sales-card__pkg" title="' +
                    escapeHtml(hint) +
                    '">' +
                    escapeHtml(packageLabel(pid) || pid) +
                    (hint ? ' · ' + escapeHtml(hint.replace(' / Jahr', '')) : '') +
                    '</span>'
                );
            })
            .join('');
        const highlights = (b.highlights || [])
            .map(function (h) {
                return '<li>' + escapeHtml(h) + '</li>';
            })
            .join('');
        const isMatch = matched.bundle && matched.bundle.id === b.id;
        const applyBtn =
            mode === 'apply' && onApply
                ? '<button type="button" class="btn btn-sm alt" data-sales-bundle-apply="' +
                  escapeHtml(b.id) +
                  '">' +
                  (isMatch ? 'Erneut setzen' : 'Bündel setzen') +
                  '</button>'
                : '';

        return (
            '<article class="admin-sales-card' +
            (b.featured ? ' is-featured' : '') +
            (isMatch ? ' is-match' : '') +
            '" data-sales-bundle="' +
            escapeHtml(b.id) +
            '">' +
            '<div class="admin-sales-card__top">' +
            '<div>' +
            '<p class="admin-sales-card__eyebrow">' +
            escapeHtml(b.audience) +
            (isMatch ? ' · aktuell' : '') +
            '</p>' +
            '<h3 class="admin-sales-card__title">' +
            escapeHtml(b.label) +
            '</h3>' +
            '<p class="admin-sales-card__price">' +
            escapeHtml(b.priceLabel) +
            '</p>' +
            (listSum
                ? '<p class="admin-sales-card__compare">Einzelpakete ' +
                  escapeHtml(String(listSum)) +
                  ' € · Bundle spart ' +
                  escapeHtml(String(save)) +
                  ' €</p>'
                : '') +
            '</div>' +
            (b.featured
                ? '<span class="admin-sales-card__badge">Empfohlen</span>'
                : '') +
            '</div>' +
            '<p class="admin-sales-card__tagline">' +
            escapeHtml(b.tagline) +
            '</p>' +
            '<ul class="admin-sales-card__list">' +
            highlights +
            '</ul>' +
            '<div class="admin-sales-card__pkgs">' +
            pkgs +
            '</div>' +
            applyBtn +
            '</article>'
        );
    }).join('');

    const addons = SALES_ADDONS.map(function (a) {
        return (
            '<li><strong>' +
            escapeHtml(a.label) +
            '</strong> · ' +
            escapeHtml(a.priceLabel) +
            ' — ' +
            escapeHtml(a.tagline) +
            '</li>'
        );
    }).join('');

    const moduleRows = LICENSE_PACKAGES.map(function (p) {
        return (
            '<tr>' +
            '<td><strong>' +
            escapeHtml(p.label) +
            '</strong><br><code>' +
            escapeHtml(p.id) +
            '</code></td>' +
            '<td>' +
            escapeHtml(p.priceHint || '–') +
            '</td>' +
            '<td>' +
            escapeHtml(String((p.toolIds || []).length)) +
            '</td>' +
            '</tr>'
        );
    }).join('');

    host.innerHTML =
        '<div class="admin-sales-panel">' +
        '<div class="admin-sales-panel__head">' +
        '<div>' +
        '<h3 class="admin-sales-panel__title">Verkaufsbündel</h3>' +
        '<p class="admin-sales-panel__lead">' +
        (mode === 'apply'
            ? 'Bundle setzen oder darunter einzelne Pakete nachjustieren.'
            : 'Bundle-Preise und Listenpreise je Paket.') +
        '</p>' +
        '</div>' +
        '</div>' +
        '<div class="admin-sales-grid">' +
        cards +
        '</div>' +
        '<details class="admin-sales-fold">' +
        '<summary>Add-ons</summary>' +
        '<ul class="admin-sales-addons__list">' +
        addons +
        '</ul>' +
        '</details>' +
        '<details class="admin-sales-fold">' +
        '<summary>Listenpreise je Paket</summary>' +
        '<table class="admin-sales-modules__table">' +
        '<thead><tr><th>Paket</th><th>Preis</th><th>Tools</th></tr></thead>' +
        '<tbody>' +
        moduleRows +
        '</tbody></table>' +
        '</details>' +
        '</div>';

    if (mode === 'apply' && onApply) {
        host.querySelectorAll('[data-sales-bundle-apply]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const id = btn.getAttribute('data-sales-bundle-apply');
                if (id) onApply(id);
            });
        });
    }
}
