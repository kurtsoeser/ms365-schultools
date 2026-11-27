/**
 * Kacheln „Für dich“ auf der Portal-Startseite.
 * Im Alltag fest. „Anordnen“ öffnet das Raster, „Fertig“ speichert es für dieses Konto.
 */
import { DASHBOARD_TOOL_ICONS } from './dashboard-audience-catalog.js';
import {
    PORTAL_CANVAS_COLS,
    PORTAL_CANVAS_GAP_PX,
    PORTAL_CANVAS_ROW_PX,
    addCanvasTile,
    clampCanvasTile,
    defaultCanvasTiles,
    hasCanvasLayout,
    mergeCanvasTiles,
    placeCanvasTile,
    readCanvasLayout,
    removeCanvasTile,
    reorderCanvasTiles,
    sortCanvasTiles,
    writeCanvasLayout
} from './portal-canvas-layout.js';
import {
    PORTAL_WIDGET_DEFAULT_IDS,
    canonicalizeWidgetId,
    portalWidgetById,
    portalWidgetSizes
} from './portal-widget-catalog.js';

let generation = 0;

export function readPortalAccountEmail() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetUserPrincipalName === 'function') {
            return String(window.ms365AuthGetUserPrincipalName() || '')
                .trim()
                .toLowerCase();
        }
    } catch {
        /* ignore */
    }
    return '';
}

/**
 * @param {string} view
 */
function readCanvasIdentity(view) {
    let tenantId = '';
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            tenantId = info && info.tenantId ? String(info.tenantId).trim() : '';
        }
    } catch {
        tenantId = '';
    }
    return {
        email: readPortalAccountEmail(),
        tenantId,
        persona: view === 'schueler' ? 'schueler' : 'lehrer'
    };
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * @param {string} toolId
 */
function iconClass(toolId) {
    const raw = DASHBOARD_TOOL_ICONS[toolId] || 'bi-grid';
    return /^bi-[a-z0-9-]+$/.test(raw) ? raw : 'bi-grid';
}

/**
 * @param {object} card
 */
function summaryText(card) {
    if (card && card.summary) return String(card.summary);
    const n = card && Array.isArray(card.lines) ? card.lines.length : 0;
    if (!n) return '';
    return n === 1 ? '1 Eintrag' : n + ' Einträge';
}

/**
 * @param {object} card
 */
function cardKey(card) {
    return String((card && (card.widgetId || card.toolId)) || '').trim();
}

/**
 * @param {object[]} tiles
 */
function canonicalTiles(tiles) {
    return (tiles || []).map((tile) => clampCanvasTile({ ...tile, id: canonicalizeWidgetId(tile.id) }));
}

/**
 * @param {object[]} cards
 * @param {{ email?: string, tenantId?: string, persona?: string }} identity
 * @param {object[]|null} draft
 */
function placedTiles(cards, identity, draft) {
    const available = cards.map(cardKey).filter(Boolean);
    const sizes = portalWidgetSizes();
    if (draft) return mergeCanvasTiles(available, canonicalTiles(draft), { strict: true });
    if (!hasCanvasLayout(localStorage, identity)) {
        const defaults = PORTAL_WIDGET_DEFAULT_IDS.filter((id) => available.includes(id));
        return defaultCanvasTiles(defaults.length ? defaults : available, sizes);
    }
    return mergeCanvasTiles(available, canonicalTiles(safeReadLayout(identity)), { strict: true });
}

/**
 * @param {object} tile
 */
function tileStyle(tile) {
    const t = clampCanvasTile(tile);
    return 'grid-column:' + (t.x + 1) + ' / span ' + t.w + ';grid-row:' + (t.y + 1) + ' / span ' + t.h + ';';
}

/**
 * @param {{ email?: string, tenantId?: string, persona?: string }} identity
 */
function safeReadLayout(identity) {
    try {
        return readCanvasLayout(localStorage, identity);
    } catch {
        return [];
    }
}

/**
 * @param {{ email?: string, tenantId?: string, persona?: string }} identity
 * @param {object[]} tiles
 */
function safeWriteLayout(identity, tiles) {
    try {
        writeCanvasLayout(localStorage, identity, tiles);
    } catch {
        /* Dieser Browser blockiert den Speicher. Die Anordnung gilt dann nur bis zum Neuladen. */
    }
}

/**
 * @param {object[]} cards
 * @param {{ tiles?: object[], arranging?: boolean, picker?: boolean }} [opts]
 */
export function renderPortalBriefingHtml(cards, opts) {
    const list = Array.isArray(cards) ? cards.filter(Boolean) : [];
    if (!list.length) return '';
    const o = opts || {};
    const arranging = !!o.arranging;
    const byId = new Map(list.map((card) => [cardKey(card), card]));
    const tiles = sortCanvasTiles(
        Array.isArray(o.tiles)
            ? o.tiles
            : defaultCanvasTiles(list.map(cardKey))
    );
    const grid = tiles
        .map((tile) => {
            const rawId = String(tile.id || '');
            const card = byId.get(rawId) || byId.get(canonicalizeWidgetId(rawId));
            if (!card) return '';
            const href = escapeHtml(card.href || '#');
            const presentation = card.presentation || 'list';
            const lineTag = arranging ? 'div' : 'a';
            const lineHref = arranging ? '' : ` href="${href}"`;
            let body = '';
            if (presentation === 'stat') {
                body =
                    `<div class="dash-portal-briefing__stat">` +
                    `<strong>${escapeHtml(card.stat || '0')}</strong>` +
                    `<span>${escapeHtml(card.statLabel || '')}</span>` +
                    `</div>`;
            } else if (presentation === 'week') {
                const days = Array.isArray(card.days) ? card.days : [];
                body = days.length
                    ? `<div class="dash-portal-briefing__week">` +
                      days
                          .map((day) => {
                              const items = Array.isArray(day.items) ? day.items : [];
                              return (
                                  `<div class="dash-portal-briefing__day${day.today ? ' is-today' : ''}">` +
                                  `<span class="dash-portal-briefing__day-label">${escapeHtml(day.label)}</span>` +
                                  (items.length
                                      ? items
                                            .map(
                                                (item) =>
                                                    `<span class="dash-portal-briefing__day-item">${escapeHtml(item)}</span>`
                                            )
                                            .join('')
                                      : `<span class="dash-portal-briefing__day-item is-empty">–</span>`) +
                                  `</div>`
                              );
                          })
                          .join('') +
                      `</div>`
                    : `<p class="dash-portal-briefing__empty">${escapeHtml(card.emptyText || '')}</p>`;
            } else {
                const lines = Array.isArray(card.lines) ? card.lines : [];
                body = lines.length
                    ? `<div class="dash-portal-briefing__lines">` +
                      lines
                          .map((line) => {
                              const kicker = line.kicker
                                  ? `<span class="dash-portal-briefing__kicker">${escapeHtml(line.kicker)}</span>`
                                  : '';
                              return (
                                  `<${lineTag} class="dash-portal-briefing__line${kicker ? ' dash-portal-briefing__line--split' : ''}"${lineHref}>` +
                                  kicker +
                                  `<span class="dash-portal-briefing__copy">` +
                                  `<span class="dash-portal-briefing__label">${escapeHtml(line.label)}</span>` +
                                  (line.meta
                                      ? `<span class="dash-portal-briefing__meta">${escapeHtml(line.meta)}</span>`
                                      : '') +
                                  `</span></${lineTag}>`
                              );
                          })
                          .join('') +
                      `</div>`
                    : `<p class="dash-portal-briefing__empty">${escapeHtml(card.emptyText || '')}</p>`;
            }
            const attention = card.attention ? ' data-attention="true"' : '';
            const summary = presentation === 'stat' ? '' : summaryText(card);
            const headInner =
                `<span class="dash-portal-briefing__icon" aria-hidden="true"><i class="bi ${iconClass(card.toolId)}"></i></span>` +
                `<span class="dash-portal-briefing__heading">` +
                `<strong>${escapeHtml(card.title)}</strong>` +
                (summary ? `<small>${escapeHtml(summary)}</small>` : '') +
                `</span>`;
            const head = arranging
                ? `<button type="button" class="dash-portal-briefing__head dash-portal-briefing__head--drag" data-drag>${headInner}</button>`
                : `<a class="dash-portal-briefing__head" href="${href}">${headInner}</a>`;
            const resize = arranging
                ? `<button type="button" class="dash-portal-briefing__resize" data-resize aria-label="Größe ändern"></button>`
                : '';
            const remove = arranging
                ? `<button type="button" class="dash-portal-briefing__remove" data-canvas-remove aria-label="Kachel entfernen">×</button>`
                : '';
            const tab = arranging ? ' tabindex="0"' : '';
            return (
                `<article class="dash-portal-briefing__card dash-portal-briefing__card--${escapeHtml(presentation)}" data-briefing="${escapeHtml(cardKey(card))}"${attention}${tab} style="${tileStyle(tile)}">` +
                remove +
                head +
                body +
                resize +
                `</article>`
            );
        })
        .join('');
    const placed = new Set(tiles.map((tile) => canonicalizeWidgetId(tile.id)));
    const actions = arranging
        ? `<div class="dash-portal-briefing__actions">` +
          `<button type="button" class="dash-portal-briefing__action" data-canvas-add><i class="bi bi-plus-lg" aria-hidden="true"></i> Kachel</button>` +
          `<button type="button" class="dash-portal-briefing__action" data-canvas-reset>Zurücksetzen</button>` +
          `<button type="button" class="dash-portal-briefing__action dash-portal-briefing__action--primary" data-canvas-toggle><i class="bi bi-check2" aria-hidden="true"></i> Fertig</button>` +
          `</div>`
        : `<button type="button" class="dash-portal-briefing__action" data-canvas-toggle><i class="bi bi-arrows-move" aria-hidden="true"></i> Anordnen</button>`;
    const groups = [];
    if (arranging && o.picker) {
        list.forEach((card) => {
            const spec = portalWidgetById(cardKey(card));
            const group = (spec && spec.group) || 'Weitere';
            if (!groups.includes(group)) groups.push(group);
        });
    }
    const picker =
        arranging && o.picker
            ? `<div class="dash-portal-briefing__picker">` +
              groups
                  .map((group) => {
                      const items = list
                          .filter((card) => ((portalWidgetById(cardKey(card)) || {}).group || 'Weitere') === group)
                          .map((card) => {
                              const id = cardKey(card);
                              const spec = portalWidgetById(id);
                              const on = placed.has(id);
                              return (
                                  `<button type="button" class="dash-portal-briefing__pick" data-canvas-pick="${escapeHtml(id)}"${on ? ' disabled' : ''}>` +
                                  `<strong>${escapeHtml(card.title)}</strong>` +
                                  `<small>${escapeHtml((spec && spec.blurb) || (on ? 'Liegt auf dem Raster' : ''))}</small>` +
                                  `</button>`
                              );
                          })
                          .join('');
                      return `<p class="dash-portal-briefing__pick-group">${escapeHtml(group)}</p>` + items;
                  })
                  .join('') +
              `</div>`
            : '';
    const hint = arranging
        ? `<p class="dash-portal-briefing__hint">Ziehen zum Verschieben, Ecke für die Breite und Höhe. Kachel hinzufügen oder entfernen. Fertig speichert die Anordnung für dieses Konto auf diesem Gerät.</p>`
        : '';
    const emptyBoard = tiles.length
        ? ''
        : `<p class="dash-portal-briefing__hint">Noch keine Kachel auf dem Raster. Mit „Kachel“ eine hinzufügen.</p>`;
    return (
        `<div class="dash-portal-briefing__bar">` +
        `<h2 class="dash-portal-section__title">Für dich</h2>` +
        actions +
        `</div>` +
        hint +
        picker +
        emptyBoard +
        `<div class="dash-portal-briefing__grid${arranging ? ' is-arranging' : ''}">${grid}</div>`
    );
}

/**
 * @param {HTMLElement} grid
 */
function gridMetrics(grid) {
    const rect = grid.getBoundingClientRect();
    const gap = PORTAL_CANVAS_GAP_PX;
    const cols = PORTAL_CANVAS_COLS;
    const col = (rect.width - gap * (cols - 1)) / cols;
    return { col, row: PORTAL_CANVAS_ROW_PX, gap, stacked: getComputedStyle(grid).display !== 'grid' };
}

/**
 * @param {HTMLElement} host
 * @param {{ cards: object[], arranging: boolean, draft: object[]|null, view: string }} session
 */
function paintBriefing(host, session) {
    const identity = readCanvasIdentity(session.view);
    const tiles = placedTiles(session.cards, identity, session.arranging ? session.draft : null);
    if (session.arranging) session.draft = tiles;
    const html = renderPortalBriefingHtml(session.cards, {
        tiles,
        arranging: session.arranging,
        picker: session.picker
    });
    host.innerHTML = html;
    host.hidden = !html;
}

/**
 * @param {HTMLElement} host
 * @param {{ cards: object[], arranging: boolean, draft: object[]|null, view: string }} session
 * @param {() => void} paint
 */
function bindCanvas(host, session, paint) {
    if (host._canvasAbort) host._canvasAbort.abort();
    const ac = new AbortController();
    host._canvasAbort = ac;
    const signal = ac.signal;

    host.addEventListener(
        'click',
        (ev) => {
            const el = ev.target instanceof Element ? ev.target : null;
            if (!el || !host.contains(el)) return;
            const toggle = el.closest('[data-canvas-toggle]');
            const add = el.closest('[data-canvas-add]');
            const reset = el.closest('[data-canvas-reset]');
            const remove = el.closest('[data-canvas-remove]');
            const pick = el.closest('[data-canvas-pick]');
            if (!toggle && !add && !reset && !remove && !pick) return;
            ev.preventDefault();
            const identity = readCanvasIdentity(session.view);
            if (toggle) {
                if (session.arranging) {
                    safeWriteLayout(identity, session.draft || []);
                    session.arranging = false;
                    session.picker = false;
                    session.draft = null;
                } else {
                    session.draft = placedTiles(session.cards, identity, null);
                    session.arranging = true;
                }
            } else if (add) {
                session.picker = !session.picker;
            } else if (reset) {
                const available = session.cards.map(cardKey).filter(Boolean);
                const defaults = PORTAL_WIDGET_DEFAULT_IDS.filter((id) => available.includes(id));
                session.draft = defaultCanvasTiles(defaults.length ? defaults : available, portalWidgetSizes());
            } else if (remove) {
                const card = remove.closest('[data-briefing]');
                const id = card ? card.getAttribute('data-briefing') || '' : '';
                session.draft = removeCanvasTile(session.draft || [], id);
            } else if (pick && !pick.hasAttribute('disabled')) {
                const id = pick.getAttribute('data-canvas-pick') || '';
                const spec = portalWidgetById(id);
                session.draft = addCanvasTile(session.draft || [], id, spec ? { w: spec.w, h: spec.h } : {});
            }
            paint();
        },
        { signal }
    );

    host.addEventListener(
        'keydown',
        (ev) => {
            if (!session.arranging || !session.draft) return;
            const card = ev.target instanceof Element ? ev.target.closest('[data-briefing]') : null;
            if (!card || !host.contains(card)) return;
            const id = card.getAttribute('data-briefing') || '';
            const tile = session.draft.find((t) => t.id === id);
            if (!tile) return;
            const grow = ev.shiftKey;
            let patch = null;
            if (ev.key === 'ArrowLeft') patch = grow ? { w: tile.w - 1 } : { x: tile.x - 1 };
            else if (ev.key === 'ArrowRight') patch = grow ? { w: tile.w + 1 } : { x: tile.x + 1 };
            else if (ev.key === 'ArrowUp') patch = grow ? { h: tile.h - 1 } : { y: tile.y - 1 };
            else if (ev.key === 'ArrowDown') patch = grow ? { h: tile.h + 1 } : { y: tile.y + 1 };
            if (!patch) return;
            ev.preventDefault();
            session.draft = placeCanvasTile(session.draft, id, patch);
            paint();
            const next = host.querySelector('[data-briefing="' + CSS.escape(id) + '"]');
            if (next instanceof HTMLElement) next.focus();
        },
        { signal }
    );

    host.addEventListener(
        'pointerdown',
        (ev) => {
            if (!session.arranging || !session.draft || ev.button !== 0) return;
            const target = ev.target instanceof Element ? ev.target : null;
            if (!target) return;
            const resize = target.closest('[data-resize]');
            const drag = target.closest('[data-drag]');
            if (!resize && !drag) return;
            const card = target.closest('[data-briefing]');
            if (!(card instanceof HTMLElement) || !host.contains(card)) return;
            const id = card.getAttribute('data-briefing') || '';
            const tile = session.draft.find((t) => t.id === id);
            const grid = host.querySelector('.dash-portal-briefing__grid');
            if (!tile || !(grid instanceof HTMLElement)) return;
            ev.preventDefault();
            const mode = resize ? 'resize' : 'move';
            const metrics = gridMetrics(grid);
            if (metrics.stacked && mode === 'resize') return;
            const start = { x: ev.clientX, y: ev.clientY, tile: { ...tile }, live: null };
            card.classList.add('is-dragging');
            try {
                card.setPointerCapture(ev.pointerId);
            } catch {
                /* Zeiger bleibt auch ohne Capture am Element. */
            }

            const move = (e) => {
                if (metrics.stacked) return;
                const stepX = metrics.col + metrics.gap;
                const stepY = metrics.row + metrics.gap;
                const dx = Math.round((e.clientX - start.x) / stepX);
                const dy = Math.round((e.clientY - start.y) / stepY);
                const patch = mode === 'move'
                    ? { x: start.tile.x + dx, y: start.tile.y + dy }
                    : { w: start.tile.w + dx, h: start.tile.h + dy };
                const live = clampCanvasTile({ ...start.tile, ...patch, id });
                start.live = live;
                card.style.gridColumn = live.x + 1 + ' / span ' + live.w;
                card.style.gridRow = live.y + 1 + ' / span ' + live.h;
            };
            const up = (e) => {
                card.removeEventListener('pointermove', move);
                card.removeEventListener('pointerup', up);
                card.removeEventListener('pointercancel', up);
                if (metrics.stacked && mode === 'move') {
                    const cards = [...grid.querySelectorAll('[data-briefing]')];
                    const over = cards.find((el) => {
                        const r = el.getBoundingClientRect();
                        return e.clientY >= r.top && e.clientY <= r.bottom;
                    });
                    const to = over ? cards.indexOf(over) : cards.length - 1;
                    session.draft = reorderCanvasTiles(session.draft, id, to);
                } else if (start.live) {
                    session.draft = placeCanvasTile(session.draft, id, start.live);
                }
                paint();
            };
            card.addEventListener('pointermove', move);
            card.addEventListener('pointerup', up);
            card.addEventListener('pointercancel', up);
        },
        { signal }
    );
}

/**
 * @param {HTMLElement|null} host
 * @param {{ view?: string, isVisible?: (id: string) => boolean }} [ctx]
 */
export function mountPortalBriefing(host, ctx) {
    if (!host) return;
    const gen = ++generation;
    const view = ctx && ctx.view === 'schueler' ? 'schueler' : 'lehrer';
    const isVisible = ctx && typeof ctx.isVisible === 'function' ? ctx.isVisible : () => false;
    const session = { cards: [], arranging: false, draft: null, picker: false, view };
    const paint = () => {
        if (gen !== generation) return;
        paintBriefing(host, session);
    };
    bindCanvas(host, session, paint);
    host.hidden = true;
    host.innerHTML = '';
    import('./portal-briefing-load.js')
        .then((mod) =>
            mod.loadPortalBriefings(
                {
                    view,
                    email: readPortalAccountEmail(),
                    isVisible
                },
                (cards) => {
                    if (gen !== generation) return;
                    session.cards = Array.isArray(cards) ? cards.filter(Boolean) : [];
                    if (session.arranging) {
                        session.draft = placedTiles(
                            session.cards,
                            readCanvasIdentity(session.view),
                            session.draft || []
                        );
                    }
                    paint();
                }
            )
        )
        .catch(() => {
            if (gen !== generation) return;
            host.hidden = true;
            host.innerHTML = '';
        });
}
