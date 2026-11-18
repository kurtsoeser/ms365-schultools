/**
 * IT-UI „Portal gestalten“: Schul-Pins, Sidebar-Reihenfolge, Live-Vorschau.
 */
import {
    DASHBOARD_LEHRER_SCHUELER_APPS,
    DASHBOARD_PORTAL_SECTIONS_BY_VIEW,
    DASHBOARD_TOOL_BLURBS,
    DASHBOARD_TOOL_ICONS,
    isToolVisibleForView,
    toolLabel,
    viewLabel
} from './dashboard-audience-catalog.js';
import { getEffectiveToolRule, loadDashboardToolAccessConfig } from './dashboard-audience-store.js';
import {
    loadPortalDefaultsConfig,
    resetPortalDefaultsConfig,
    resolvePortalHomeToolIds,
    setPortalHomePinned,
    setPortalSectionOrder
} from './dashboard-portal-defaults-store.js';

const HOME_FALLBACK = [
    'schularbeiten-planer',
    'freistellung-planer',
    'lehrer-freistellung-planer',
    'schulaktivitaeten-planer',
    'termine-hub',
    'srdp-anmeldung',
    'elternbrief',
    'aufgaben',
    'projektwochen'
];

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * @param {'lehrer'|'schueler'} view
 * @returns {string[]}
 */
function portalEligibleIdsForView(view) {
    const access = loadDashboardToolAccessConfig();
    return DASHBOARD_LEHRER_SCHUELER_APPS.filter((id) => {
        const rule = getEffectiveToolRule(id, access);
        return isToolVisibleForView(id, {}, view, rule);
    }).sort((a, b) => toolLabel(a).localeCompare(toolLabel(b), 'de'));
}

/**
 * @param {string[]} pinned
 * @param {string[]} eligible
 */
function orderedPinned(pinned, eligible) {
    const elig = new Set(eligible);
    const out = [];
    const seen = new Set();
    for (const id of pinned || []) {
        if (!elig.has(id) || seen.has(id)) continue;
        seen.add(id);
        out.push(id);
    }
    return out;
}

/**
 * @param {string} id
 * @param {{ pinned?: boolean, rank?: number }} [opts]
 */
function renderPreviewCard(id, opts) {
    const icon = DASHBOARD_TOOL_ICONS[id] || 'bi-app';
    const blurb = DASHBOARD_TOOL_BLURBS[id] || '';
    const pinned = !!(opts && opts.pinned);
    const rank = opts && opts.rank ? opts.rank : 0;
    return `
        <article class="dash-pg-preview-card${pinned ? ' is-pinned' : ''}" data-tool-id="${escapeHtml(id)}">
            <div class="dash-pg-preview-card__top">
                <span class="dash-pg-preview-card__icon" aria-hidden="true"><i class="bi ${escapeHtml(icon)}"></i></span>
                ${pinned ? `<span class="dash-pg-preview-card__badge">Pin ${rank}</span>` : ''}
            </div>
            <h5 class="dash-pg-preview-card__title">${escapeHtml(toolLabel(id))}</h5>
            ${blurb ? `<p class="dash-pg-preview-card__desc">${escapeHtml(blurb)}</p>` : ''}
        </article>`;
}

/**
 * @param {HTMLElement} root
 * @param {'lehrer'|'schueler'} view
 */
function renderPersonaEditor(root, view) {
    const cfg = loadPortalDefaultsConfig();
    const eligible = portalEligibleIdsForView(view);
    const pinned = orderedPinned(cfg[view].homePinned, eligible);
    const pinnedSet = new Set(pinned);
    const sectionDefs = DASHBOARD_PORTAL_SECTIONS_BY_VIEW[view] || [];
    const sectionOrder =
        cfg[view].sectionOrder && cfg[view].sectionOrder.length
            ? [
                  ...cfg[view].sectionOrder.filter((id) => sectionDefs.some((s) => s.id === id)),
                  ...sectionDefs.map((s) => s.id).filter((id) => !cfg[view].sectionOrder.includes(id))
              ]
            : sectionDefs.map((s) => s.id);

    const previewIds = resolvePortalHomeToolIds({
        visibleIds: eligible,
        schoolPinned: pinned,
        personalFavorites: [],
        fallbackPriority: HOME_FALLBACK,
        max: 6
    });

    const navItems = [
        { id: 'home', label: 'Dashboard', icon: 'bi-house-door' },
        ...sectionOrder.map((id) => {
            const sec = sectionDefs.find((s) => s.id === id);
            return {
                id,
                label: (sec && sec.label) || id,
                icon: (sec && sec.icon) || 'bi-grid'
            };
        })
    ];

    const previewCards = previewIds
        .map((id, i) =>
            renderPreviewCard(id, {
                pinned: pinnedSet.has(id),
                rank: pinnedSet.has(id) ? pinned.indexOf(id) + 1 : 0
            })
        )
        .join('');

    const pinControls = pinned
        .map(
            (id, idx) => `
        <li class="dash-pg-pin-row" data-pin-id="${escapeHtml(id)}">
            <span class="dash-pg-pin-row__rank">${idx + 1}</span>
            <span class="dash-pg-pin-row__icon" aria-hidden="true"><i class="bi ${escapeHtml(DASHBOARD_TOOL_ICONS[id] || 'bi-app')}"></i></span>
            <span class="dash-pg-pin-row__label">${escapeHtml(toolLabel(id))}</span>
            <span class="dash-pg-pin-row__actions">
                <button type="button" class="btn btn-sm" data-pin-up="${escapeHtml(id)}" title="Nach oben" ${idx === 0 ? 'disabled' : ''}><i class="bi bi-arrow-up"></i></button>
                <button type="button" class="btn btn-sm" data-pin-down="${escapeHtml(id)}" title="Nach unten" ${idx === pinned.length - 1 ? 'disabled' : ''}><i class="bi bi-arrow-down"></i></button>
                <button type="button" class="btn btn-sm" data-pin-remove="${escapeHtml(id)}" title="Pin entfernen"><i class="bi bi-x-lg"></i></button>
            </span>
        </li>`
        )
        .join('');

    const pickRows = eligible
        .map((id) => {
            const on = pinnedSet.has(id);
            const disabled = !on && pinned.length >= 6;
            return `
            <label class="dash-pg-pick${on ? ' is-on' : ''}${disabled ? ' is-disabled' : ''}">
                <input type="checkbox" data-pin-toggle="${escapeHtml(id)}" ${on ? 'checked' : ''} ${disabled ? 'disabled' : ''} />
                <span class="dash-pg-pick__icon" aria-hidden="true"><i class="bi ${escapeHtml(DASHBOARD_TOOL_ICONS[id] || 'bi-app')}"></i></span>
                <span class="dash-pg-pick__text">
                    <span class="dash-pg-pick__title">${escapeHtml(toolLabel(id))}</span>
                    ${DASHBOARD_TOOL_BLURBS[id] ? `<span class="dash-pg-pick__hint">${escapeHtml(DASHBOARD_TOOL_BLURBS[id])}</span>` : ''}
                </span>
            </label>`;
        })
        .join('');

    const sectionRows = sectionOrder
        .map((id, idx) => {
            const sec = sectionDefs.find((s) => s.id === id);
            const label = sec ? sec.label : id;
            const icon = (sec && sec.icon) || 'bi-grid';
            return `
        <li class="dash-pg-sec-row" data-sec-id="${escapeHtml(id)}">
            <span class="dash-pg-sec-row__label"><i class="bi ${escapeHtml(icon)}" aria-hidden="true"></i>${escapeHtml(label)}</span>
            <span class="dash-pg-pin-row__actions">
                <button type="button" class="btn btn-sm" data-sec-up="${escapeHtml(id)}" ${idx === 0 ? 'disabled' : ''}><i class="bi bi-arrow-up"></i></button>
                <button type="button" class="btn btn-sm" data-sec-down="${escapeHtml(id)}" ${idx === sectionOrder.length - 1 ? 'disabled' : ''}><i class="bi bi-arrow-down"></i></button>
            </span>
        </li>`;
        })
        .join('');

    root.innerHTML = `
        <div class="dash-pg" data-portal-defaults-view="${escapeHtml(view)}">
            <div class="dash-pg__layout">
                <aside class="dash-pg__controls">
                    <div class="dash-pg__panel">
                        <h4 class="dash-pg__panel-title">Home anheften <span class="dash-pg__count">${pinned.length}/6</span></h4>
                        <p class="dash-pg__panel-lead">Max. sechs Apps erscheinen oben auf dem Portal-Home. Reihenfolge = Anzeigereihenfolge.</p>
                        <ol class="dash-pg-pin-list">${pinControls || '<li class="dash-pg__empty">Noch keine Pins – Vorschau zeigt die Standard-Apps.</li>'}</ol>
                    </div>
                    <div class="dash-pg__panel">
                        <h4 class="dash-pg__panel-title">Sidebar-Reihenfolge</h4>
                        <ol class="dash-pg-sec-list">${sectionRows}</ol>
                    </div>
                    <div class="dash-pg__panel dash-pg__panel--pick">
                        <h4 class="dash-pg__panel-title">Apps für ${escapeHtml(viewLabel(view))} (${eligible.length})</h4>
                        <p class="dash-pg__panel-lead">Anhaken = auf dem Home anheften. Sichtbarkeit steuert die Matrix darunter.</p>
                        <div class="dash-pg-pick-grid">${pickRows || '<p class="dash-pg__empty">Keine portal-fähigen Apps sichtbar.</p>'}</div>
                    </div>
                </aside>
                <div class="dash-pg__preview-wrap">
                    <div class="dash-pg__preview-head">
                        <h4 class="dash-pg__panel-title">Vorschau Portal · ${escapeHtml(viewLabel(view))}</h4>
                        <p class="dash-pg__panel-lead">So sieht das Schulportal näherungsweise aus (ohne persönliche Favoriten).</p>
                    </div>
                    <div class="dash-pg-shell" aria-label="Portal-Vorschau">
                        <nav class="dash-pg-shell__nav" aria-label="Sidebar-Vorschau">
                            ${navItems
                                .map(
                                    (item, i) => `
                                <div class="dash-pg-shell__nav-item${i === 0 ? ' is-active' : ''}">
                                    <i class="bi ${escapeHtml(item.icon)}" aria-hidden="true"></i>
                                    <span>${escapeHtml(item.label)}</span>
                                </div>`
                                )
                                .join('')}
                        </nav>
                        <div class="dash-pg-shell__main">
                            <header class="dash-pg-shell__hero">
                                <p class="dash-pg-shell__eyebrow">${escapeHtml(viewLabel(view))}</p>
                                <p class="dash-pg-shell__hello">Guten Tag!</p>
                                <p class="dash-pg-shell__lead">Hier findest du die wichtigsten Schultools für den Alltag.</p>
                            </header>
                            <section class="dash-pg-shell__apps">
                                <h5 class="dash-pg-shell__apps-title">Meine Apps</h5>
                                <div class="dash-pg-shell__grid">
                                    ${previewCards || '<p class="dash-pg__empty">Keine Apps in der Vorschau.</p>'}
                                </div>
                            </section>
                        </div>
                    </div>
                </div>
            </div>
        </div>`;

    bindPersonaEditor(root, view, pinned, sectionOrder);
}

/**
 * @param {HTMLElement} root
 * @param {'lehrer'|'schueler'} view
 * @param {string[]} pinned
 * @param {string[]} sectionOrder
 */
function bindPersonaEditor(root, view, pinned, sectionOrder) {
    const refresh = () => renderPersonaEditor(root, view);

    root.querySelectorAll('[data-pin-toggle]').forEach((el) => {
        el.addEventListener('change', () => {
            const id = el.getAttribute('data-pin-toggle');
            if (!id) return;
            const on = /** @type {HTMLInputElement} */ (el).checked;
            let next = pinned.slice();
            if (on) {
                if (next.length >= 6) return;
                if (!next.includes(id)) next.push(id);
            } else {
                next = next.filter((x) => x !== id);
            }
            setPortalHomePinned(view, next);
            refresh();
        });
    });

    root.querySelectorAll('[data-pin-remove]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-pin-remove');
            setPortalHomePinned(
                view,
                pinned.filter((x) => x !== id)
            );
            refresh();
        });
    });

    root.querySelectorAll('[data-pin-up]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-pin-up');
            const i = pinned.indexOf(id || '');
            if (i <= 0) return;
            const next = pinned.slice();
            [next[i - 1], next[i]] = [next[i], next[i - 1]];
            setPortalHomePinned(view, next);
            refresh();
        });
    });

    root.querySelectorAll('[data-pin-down]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-pin-down');
            const i = pinned.indexOf(id || '');
            if (i < 0 || i >= pinned.length - 1) return;
            const next = pinned.slice();
            [next[i], next[i + 1]] = [next[i + 1], next[i]];
            setPortalHomePinned(view, next);
            refresh();
        });
    });

    root.querySelectorAll('[data-sec-up]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-sec-up');
            const i = sectionOrder.indexOf(id || '');
            if (i <= 0) return;
            const next = sectionOrder.slice();
            [next[i - 1], next[i]] = [next[i], next[i - 1]];
            setPortalSectionOrder(view, next);
            refresh();
        });
    });

    root.querySelectorAll('[data-sec-down]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-sec-down');
            const i = sectionOrder.indexOf(id || '');
            if (i < 0 || i >= sectionOrder.length - 1) return;
            const next = sectionOrder.slice();
            [next[i], next[i + 1]] = [next[i + 1], next[i]];
            setPortalSectionOrder(view, next);
            refresh();
        });
    });
}

function renderPortalDefaultsAdmin(root) {
    root.innerHTML = `
        <section class="dash-pg-section" aria-labelledby="dashPortalDefaultsTitle">
            <div class="dash-pg-section__head">
                <div>
                    <h3 id="dashPortalDefaultsTitle">Portal gestalten</h3>
                    <p class="dash-pg-section__lead">
                        Bereite das Schulportal für Lehrkräfte und Schüler vor: Home-Apps anheften, Sidebar sortieren –
                        die Vorschau rechts aktualisiert sich sofort. Persönliche Favoriten der User kommen danach dazu.
                    </p>
                </div>
                <button type="button" class="btn btn-sm" id="dashPortalDefaultsReset">Pins zurücksetzen</button>
            </div>
            <div class="dash-portal-defaults__tabs" role="tablist" aria-label="Portal-Persona">
                <button type="button" class="btn btn-sm is-active" role="tab" aria-selected="true" data-portal-defaults-tab="lehrer"><i class="bi bi-person-badge" aria-hidden="true"></i> Lehrkraft</button>
                <button type="button" class="btn btn-sm" role="tab" aria-selected="false" data-portal-defaults-tab="schueler"><i class="bi bi-mortarboard" aria-hidden="true"></i> Schüler/in</button>
            </div>
            <div data-portal-defaults-editor></div>
            <p class="dash-pg-section__note">
                Speicherung: <code>ms365-dashboard-portal-defaults-v1</code> (Schul-Config, Backup / SharePoint).
            </p>
        </section>`;

    const editor = root.querySelector('[data-portal-defaults-editor]');
    /** @type {'lehrer'|'schueler'} */
    let active = 'lehrer';

    const show = () => {
        if (editor) renderPersonaEditor(/** @type {HTMLElement} */ (editor), active);
        root.querySelectorAll('[data-portal-defaults-tab]').forEach((btn) => {
            const id = btn.getAttribute('data-portal-defaults-tab');
            const on = id === active;
            btn.classList.toggle('is-active', on);
            btn.setAttribute('aria-selected', on ? 'true' : 'false');
        });
    };

    root.querySelectorAll('[data-portal-defaults-tab]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-portal-defaults-tab');
            if (id === 'lehrer' || id === 'schueler') {
                active = id;
                show();
            }
        });
    });

    root.querySelector('#dashPortalDefaultsReset')?.addEventListener('click', () => {
        resetPortalDefaultsConfig();
        show();
    });

    show();
    window.addEventListener('ms365-dashboard-tool-access-changed', () => show());
}

/**
 * @param {string} [mountSelector]
 */
export function mountDashboardPortalDefaultsAdmin(mountSelector) {
    const sel = mountSelector || '[data-ms365-dashboard-portal-defaults-mount]';
    const root = document.querySelector(sel);
    if (!root) return;
    renderPortalDefaultsAdmin(/** @type {HTMLElement} */ (root));
}
