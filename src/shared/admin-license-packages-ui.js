/**
 * Gemeinsame Modul-/Paket-Karten für Admin (Schildetail + Modul-Katalog).
 */
import {
    LICENSE_PACKAGES,
    packageIcon,
    packageLabel,
    resolveFeatureEntitlements,
    toolCoveredByPackages
} from './license-feature-catalog.js';
import { toolLabel } from './dashboard-audience-catalog.js';

function escapeHtml(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * @param {import('./license-feature-catalog.js').LicensePackage} pkg
 * @param {{
 *   interactive?: boolean,
 *   enabled?: boolean,
 *   expanded?: boolean,
 *   onToggle?: (id: string, on: boolean) => void
 * }} [opts]
 */
export function buildPackageCard(pkg, opts) {
    const interactive = !!(opts && opts.interactive);
    const enabled = !!(opts && opts.enabled);
    const expanded = !!(opts && opts.expanded);
    const card = document.createElement('article');
    card.className =
        'admin-mod-card' +
        (enabled ? ' is-on' : '') +
        (interactive ? ' is-interactive' : ' is-readonly');
    card.dataset.packageId = pkg.id;

    const reqHtml =
        Array.isArray(pkg.requires) && pkg.requires.length
            ? '<p class="admin-mod-card__req"><i class="bi bi-link-45deg" aria-hidden="true"></i>Benötigt ' +
              escapeHtml(pkg.requires.map(packageLabel).join(', ')) +
              '</p>'
            : '';

    const toolsHtml = (pkg.toolIds || [])
        .map(function (tid) {
            return (
                '<li class="admin-mod-card__tool">' +
                '<span class="admin-mod-card__tool-label">' +
                escapeHtml(toolLabel(tid)) +
                '</span>' +
                '<code>' +
                escapeHtml(tid) +
                '</code></li>'
            );
        })
        .join('');

    const switchHtml = interactive
        ? '<button type="button" class="admin-mod-switch" role="switch" aria-checked="' +
          (enabled ? 'true' : 'false') +
          '" data-package-switch="' +
          escapeHtml(pkg.id) +
          '" aria-label="' +
          escapeHtml(pkg.label) +
          ' freischalten">' +
          '<span class="admin-mod-switch__track" aria-hidden="true"><span class="admin-mod-switch__thumb"></span></span>' +
          '</button>'
        : '<span class="admin-mod-card__badge">' +
          escapeHtml(String((pkg.toolIds || []).length)) +
          ' Tools</span>';

    card.innerHTML =
        '<div class="admin-mod-card__top">' +
        '<div class="admin-mod-card__icon" aria-hidden="true"><i class="bi ' +
        escapeHtml(packageIcon(pkg.id)) +
        '"></i></div>' +
        '<div class="admin-mod-card__head">' +
        '<h3 class="admin-mod-card__title">' +
        escapeHtml(pkg.label) +
        '</h3>' +
        '<p class="admin-mod-card__desc">' +
        escapeHtml(pkg.description) +
        '</p>' +
        reqHtml +
        '</div>' +
        switchHtml +
        '</div>' +
        '<div class="admin-mod-card__meta">' +
        '<span class="admin-mod-card__id"><code>' +
        escapeHtml(pkg.id) +
        '</code></span>' +
        (pkg.priceHint
            ? '<span class="admin-mod-card__price" title="Listenpreis Einzelmodul / Schuljahr">' +
              escapeHtml(pkg.priceHint) +
              '</span>'
            : '') +
        '<span class="admin-mod-card__count">' +
        escapeHtml(String((pkg.toolIds || []).length)) +
        ' Werkzeuge</span>' +
        '</div>' +
        '<details class="admin-mod-card__details"' +
        (expanded ? ' open' : '') +
        '>' +
        '<summary>Werkzeuge im Paket</summary>' +
        '<ul class="admin-mod-card__tools">' +
        toolsHtml +
        '</ul></details>';

    if (interactive && typeof opts.onToggle === 'function') {
        const btn = card.querySelector('[data-package-switch]');
        if (btn) {
            btn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                const next = btn.getAttribute('aria-checked') !== 'true';
                opts.onToggle(pkg.id, next);
            });
        }
        card.addEventListener('click', function (ev) {
            if (ev.target && ev.target.closest && ev.target.closest('details, summary, a, button')) {
                return;
            }
            const btn2 = card.querySelector('[data-package-switch]');
            if (!btn2) return;
            const next = btn2.getAttribute('aria-checked') !== 'true';
            opts.onToggle(pkg.id, next);
        });
    }

    return card;
}

/**
 * @param {HTMLElement} host
 * @param {{
 *   interactive?: boolean,
 *   enabledIds?: string[]|Record<string, boolean>,
 *   onToggle?: (id: string, on: boolean) => void,
 *   expandAll?: boolean
 * }} [opts]
 */
export function renderPackageGrid(host, opts) {
    if (!host) return;
    const interactive = !!(opts && opts.interactive);
    const expandAll = !!(opts && opts.expandAll);
    /** @type {Record<string, boolean>} */
    const enabled = {};
    if (opts && Array.isArray(opts.enabledIds)) {
        opts.enabledIds.forEach(function (id) {
            enabled[id] = true;
        });
    } else if (opts && opts.enabledIds && typeof opts.enabledIds === 'object') {
        Object.keys(opts.enabledIds).forEach(function (id) {
            if (opts.enabledIds[id]) enabled[id] = true;
        });
    }

    host.replaceChildren();
    host.classList.add('admin-mod-grid');
    LICENSE_PACKAGES.forEach(function (pkg) {
        host.appendChild(
            buildPackageCard(pkg, {
                interactive: interactive,
                enabled: !!enabled[pkg.id],
                expanded: expandAll,
                onToggle: opts && opts.onToggle
            })
        );
    });
}

/**
 * @param {number} packageCount
 * @param {number} toolCount
 * @param {{ packageCaption?: string, toolCaption?: string, catalogCaption?: string, hideCatalog?: boolean }} [opts]
 */
export function packageStatsHtml(packageCount, toolCount, opts) {
    const packageCaption = (opts && opts.packageCaption) || 'Pakete aktiv';
    const toolCaption = (opts && opts.toolCaption) || 'Werkzeuge';
    const catalogCaption = (opts && opts.catalogCaption) || 'im Katalog';
    const hideCatalog = !!(opts && opts.hideCatalog);
    return (
        '<div class="admin-mod-stats" role="status">' +
        '<div class="admin-mod-stats__item"><strong>' +
        escapeHtml(String(packageCount)) +
        '</strong><span>' +
        escapeHtml(packageCaption) +
        '</span></div>' +
        '<div class="admin-mod-stats__item"><strong>' +
        escapeHtml(String(toolCount)) +
        '</strong><span>' +
        escapeHtml(toolCaption) +
        '</span></div>' +
        (hideCatalog
            ? ''
            : '<div class="admin-mod-stats__item"><strong>' +
              escapeHtml(String(LICENSE_PACKAGES.length)) +
              '</strong><span>' +
              escapeHtml(catalogCaption) +
              '</span></div>') +
        '</div>'
    );
}

/**
 * @param {{ packages?: string[], toolsInclude?: string[], toolsExclude?: string[] }} draft
 * @param {string} toolId
 * @returns {'package'|'include'|'exclude'|'off'}
 */
export function toolEntitlementKind(draft, toolId) {
    const id = String(toolId || '');
    const exclude = Array.isArray(draft && draft.toolsExclude) ? draft.toolsExclude : [];
    const include = Array.isArray(draft && draft.toolsInclude) ? draft.toolsInclude : [];
    if (exclude.indexOf(id) >= 0) return 'exclude';
    if (toolCoveredByPackages(draft, id)) return 'package';
    if (include.indexOf(id) >= 0) return 'include';
    return 'off';
}

/**
 * Werkzeugliste gruppiert nach Paketen, mit Paket-Master- und Einzelschaltern.
 * @param {HTMLElement} host
 * @param {{
 *   draft: { packages: string[], toolsInclude: string[], toolsExclude: string[] },
 *   search?: string,
 *   onToggle?: (toolId: string, on: boolean) => void,
 *   onPackageToggle?: (packageId: string, on: boolean) => void
 * }} opts
 */
export function renderToolEntitlementGroups(host, opts) {
    if (!host || !opts || !opts.draft) return;
    const draft = opts.draft;
    const q = String(opts.search || '')
        .trim()
        .toLowerCase();
    const resolved = resolveFeatureEntitlements(draft, { legacyFullAccess: false });
    const enabledSet = {};
    (resolved.features || []).forEach(function (id) {
        enabledSet[id] = true;
    });
    const pkgOn = {};
    (draft.packages || []).forEach(function (id) {
        pkgOn[id] = true;
    });

    host.replaceChildren();

    LICENSE_PACKAGES.forEach(function (pkg) {
        const tools = (pkg.toolIds || [])
            .slice()
            .filter(function (tid) {
                if (!q) return true;
                const hay = (toolLabel(tid) + ' ' + tid).toLowerCase();
                return hay.indexOf(q) >= 0;
            })
            .sort(function (a, b) {
                return toolLabel(a).localeCompare(toolLabel(b), 'de');
            });
        if (!tools.length) return;

        const onCount = tools.filter(function (tid) {
            return !!enabledSet[tid];
        }).length;
        const packageEnabled = !!pkgOn[pkg.id];

        const section = document.createElement('section');
        section.className =
            'admin-lic-tools__group' + (packageEnabled ? ' is-pkg-on' : ' is-pkg-off');
        section.dataset.packageId = pkg.id;

        const reqHtml =
            Array.isArray(pkg.requires) && pkg.requires.length
                ? '<span class="admin-lic-tools__req"><i class="bi bi-link-45deg" aria-hidden="true"></i>Benötigt ' +
                  escapeHtml(pkg.requires.map(packageLabel).join(', ')) +
                  '</span>'
                : '';

        section.innerHTML =
            '<header class="admin-lic-tools__group-head">' +
            '<div class="admin-lic-tools__group-title">' +
            '<span class="admin-lic-tools__group-icon" aria-hidden="true"><i class="bi ' +
            escapeHtml(packageIcon(pkg.id)) +
            '"></i></span>' +
            '<div>' +
            '<strong>' +
            escapeHtml(pkg.label) +
            '</strong>' +
            '<small>' +
            escapeHtml(String(onCount)) +
            ' / ' +
            escapeHtml(String(tools.length)) +
            ' aktiv' +
            (packageEnabled ? '' : ' · Paket aus') +
            '</small>' +
            '</div></div>' +
            '<div class="admin-lic-tools__group-actions">' +
            reqHtml +
            '<button type="button" class="admin-mod-switch" role="switch" aria-checked="' +
            (packageEnabled ? 'true' : 'false') +
            '" data-package-tools-switch="' +
            escapeHtml(pkg.id) +
            '" aria-label="Alle Werkzeuge in ' +
            escapeHtml(pkg.label) +
            ' freischalten" title="Alle Werkzeuge im Paket">' +
            '<span class="admin-mod-switch__track" aria-hidden="true"><span class="admin-mod-switch__thumb"></span></span>' +
            '</button>' +
            '</div></header>';

        const pkgBtn = section.querySelector('[data-package-tools-switch]');
        if (pkgBtn && typeof opts.onPackageToggle === 'function') {
            pkgBtn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                opts.onPackageToggle(pkg.id, pkgBtn.getAttribute('aria-checked') !== 'true');
            });
        }

        const list = document.createElement('ul');
        list.className = 'admin-lic-tools__rows';

        tools.forEach(function (tid) {
            const kind = toolEntitlementKind(draft, tid);
            const on = !!enabledSet[tid];
            const li = document.createElement('li');
            li.className =
                'admin-lic-tools__row' +
                (on ? ' is-on' : '') +
                (kind === 'include' ? ' is-override-on' : '') +
                (kind === 'exclude' ? ' is-override-off' : '');

            let badge = '';
            if (kind === 'include') badge = 'einzeln an';
            else if (kind === 'exclude') badge = 'einzeln aus';

            li.innerHTML =
                '<div class="admin-lic-tools__row-text">' +
                '<span class="admin-lic-tools__row-label">' +
                escapeHtml(toolLabel(tid)) +
                '</span>' +
                '<span class="admin-lic-tools__row-meta">' +
                '<code>' +
                escapeHtml(tid) +
                '</code>' +
                (badge
                    ? '<span class="admin-lic-tools__badge">' + escapeHtml(badge) + '</span>'
                    : '') +
                '</span></div>' +
                '<button type="button" class="admin-mod-switch" role="switch" aria-checked="' +
                (on ? 'true' : 'false') +
                '" data-tool-switch="' +
                escapeHtml(tid) +
                '" aria-label="' +
                escapeHtml(toolLabel(tid)) +
                ' freischalten">' +
                '<span class="admin-mod-switch__track" aria-hidden="true"><span class="admin-mod-switch__thumb"></span></span>' +
                '</button>';

            const btn = li.querySelector('[data-tool-switch]');
            if (btn && typeof opts.onToggle === 'function') {
                btn.addEventListener('click', function (ev) {
                    ev.preventDefault();
                    ev.stopPropagation();
                    opts.onToggle(tid, btn.getAttribute('aria-checked') !== 'true');
                });
            }
            list.appendChild(li);
        });

        section.appendChild(list);
        host.appendChild(section);
    });

    if (!host.childElementCount) {
        const empty = document.createElement('p');
        empty.className = 'admin-lic-tools__empty';
        empty.textContent = q ? 'Keine Werkzeuge für diese Suche.' : 'Keine Werkzeuge im Katalog.';
        host.appendChild(empty);
    }
}
