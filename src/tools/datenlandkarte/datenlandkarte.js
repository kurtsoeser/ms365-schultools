/**
 * Datenlandkarte – Listen/Quellen als Blöcke mit Verknüpfungen.
 */
import { DATEN_BLOECKE, DATEN_LINKS } from './datenlandkarte-catalog.js';
import { collectDatenMetrics, mergeDatenMetrics, spoLoadingMetrics } from './datenlandkarte-metrics.js';
import { fetchSharePointListMetrics } from './datenlandkarte-spo-metrics.js';
import { layoutDatenlandkarte, recomputeLayoutFromBlocks } from './datenlandkarte-layout.js';
import {
    loadDatenlandkarteLayout,
    saveDatenlandkarteLayout,
    clearDatenlandkarteLayout,
    hasCustomDatenlandkarteLayout
} from './datenlandkarte-layout-store.js';
import { syncDatenlandkarteCanvasDom } from './datenlandkarte-layout-dom.js';
import { renderDatenlandkarteApp, applyCanvasTransform } from './datenlandkarte-ui.js';
import { setDatenlandkarteHrefContext } from './datenlandkarte-register-bridge.js';

const state = {
    metrics: {},
    layout: layoutDatenlandkarte(DATEN_BLOECKE, DATEN_LINKS, loadDatenlandkarteLayout()),
    savedLayout: loadDatenlandkarteLayout(),
    selectedBlockId: '',
    selectedLinkId: '',
    viewTransform: { x: 40, y: 20, k: 1 },
    spoStatus: 'idle',
    spoErrors: [],
    embedded: false
};

let root = null;
let spoFetchGen = 0;
/** @type {{ x: number, y: number, tx: number, ty: number, pointerId: number } | null} */
let panSession = null;
/** @type {{ id: string, pointerId: number, startClientX: number, startClientY: number, startX: number, startY: number, moved: boolean } | null} */
let blockDrag = null;

function loadSetup() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
            return window.ms365AppDataV2.getSetup() || {};
        }
    } catch {
        /* ignore */
    }
    return {};
}

function rebuildLayoutFromStateBlocks() {
    state.layout = recomputeLayoutFromBlocks(state.layout.blocks, DATEN_LINKS);
    state.layout.customized = hasCustomDatenlandkarteLayout(state.savedLayout);
}

function refreshLocalData() {
    state.metrics = collectDatenMetrics();
    state.savedLayout = loadDatenlandkarteLayout();
    state.layout = layoutDatenlandkarte(DATEN_BLOECKE, DATEN_LINKS, state.savedLayout);
}

async function refreshSharePointMetrics() {
    const gen = ++spoFetchGen;
    state.spoStatus = 'loading';
    state.spoErrors = [];
    state.metrics = mergeDatenMetrics(collectDatenMetrics(), spoLoadingMetrics());
    paint();

    const setup = loadSetup();
    let result;
    try {
        result = await fetchSharePointListMetrics(setup);
    } catch (e) {
        if (gen !== spoFetchGen) return;
        state.spoStatus = 'error';
        state.spoErrors = [e && e.message ? e.message : String(e)];
        state.metrics = collectDatenMetrics();
        paint();
        return;
    }

    if (gen !== spoFetchGen) return;
    state.metrics = mergeDatenMetrics(collectDatenMetrics(), result.metrics || {});
    state.spoErrors = result.errors || [];
    state.spoStatus = state.spoErrors.length ? 'partial' : 'ok';
    paint();
}

function refreshAll() {
    refreshLocalData();
    paint();
    void refreshSharePointMetrics();
}

function paint() {
    if (!root) return;
    renderDatenlandkarteApp(state, root);
    applyCanvasTransform(state);
}

function selectBlock(id) {
    state.selectedBlockId = id;
    state.selectedLinkId = '';
    paint();
}

function bindUiOnce() {
    if (!root || root.dataset.dlUiBound === '1') return;
    root.dataset.dlUiBound = '1';

    root.addEventListener('click', (ev) => {
        if (ev.target.closest('#dlBtnResetLayout')) {
            clearDatenlandkarteLayout();
            state.savedLayout = {};
            refreshLocalData();
            paint();
            requestAnimationFrame(() => fitView());
            return;
        }
        const block = ev.target.closest('[data-dl-block]');
        if (block) {
            if (ev.target.closest('.dl-block__link')) return;
            if (ev.target.closest('.dl-block__head')) return;
            ev.preventDefault();
            selectBlock(block.getAttribute('data-dl-block') || '');
            return;
        }
        const linkBtn = ev.target.closest('[data-dl-link]');
        if (linkBtn && linkBtn.classList.contains('dl-link-label')) {
            state.selectedLinkId = linkBtn.getAttribute('data-dl-link') || '';
            paint();
            return;
        }
        if (ev.target.closest('#dlBtnReload')) {
            refreshAll();
            return;
        }
        if (ev.target.closest('#dlBtnClearSel')) {
            state.selectedBlockId = '';
            state.selectedLinkId = '';
            paint();
            return;
        }
        if (ev.target.closest('#dlBtnFit')) {
            fitView();
        }
    });

    root.addEventListener(
        'wheel',
        (ev) => {
            const host = ev.target.closest('#dlCanvasHost');
            if (!host) return;
            ev.preventDefault();
            const delta = ev.deltaY < 0 ? 1.08 : 1 / 1.08;
            const t = state.viewTransform;
            const rect = host.getBoundingClientRect();
            const mx = ev.clientX - rect.left;
            const my = ev.clientY - rect.top;
            const nx = mx - (mx - t.x) * delta;
            const ny = my - (my - t.y) * delta;
            t.k = Math.min(2.5, Math.max(0.35, t.k * delta));
            t.x = nx;
            t.y = ny;
            applyCanvasTransform(state);
        },
        { passive: false }
    );

    root.addEventListener('pointerdown', (ev) => {
        const head = ev.target.closest('.dl-block__head');
        const blockEl = head && head.closest('.dl-block');
        if (blockEl) {
            const id = blockEl.getAttribute('data-dl-block') || '';
            blockDrag = {
                id,
                pointerId: ev.pointerId,
                startClientX: ev.clientX,
                startClientY: ev.clientY,
                startX: parseFloat(blockEl.style.left) || 0,
                startY: parseFloat(blockEl.style.top) || 0,
                moved: false
            };
            blockEl.classList.add('dl-block--dragging');
            blockEl.setPointerCapture(ev.pointerId);
            ev.preventDefault();
            return;
        }

        const host = ev.target.closest('#dlCanvasHost');
        if (!host) return;
        if (ev.target.closest('.dl-block')) return;
        panSession = {
            x: ev.clientX,
            y: ev.clientY,
            tx: state.viewTransform.x,
            ty: state.viewTransform.y,
            pointerId: ev.pointerId
        };
        host.classList.add('dl-panning');
        host.setPointerCapture(ev.pointerId);
    });

    root.addEventListener('pointermove', (ev) => {
        if (blockDrag && ev.pointerId === blockDrag.pointerId) {
            const k = state.viewTransform.k || 1;
            const dx = (ev.clientX - blockDrag.startClientX) / k;
            const dy = (ev.clientY - blockDrag.startClientY) / k;
            if (Math.abs(dx) > 3 || Math.abs(dy) > 3) blockDrag.moved = true;
            const nx = Math.max(0, blockDrag.startX + dx);
            const ny = Math.max(0, blockDrag.startY + dy);
            const blockEl = root.querySelector(`[data-dl-block="${CSS.escape(blockDrag.id)}"]`);
            if (blockEl) {
                blockEl.style.left = `${nx}px`;
                blockEl.style.top = `${ny}px`;
            }
            const b = state.layout.blocks.find((x) => x.id === blockDrag.id);
            if (b) {
                b.x = nx;
                b.y = ny;
            }
            rebuildLayoutFromStateBlocks();
            syncDatenlandkarteCanvasDom(state);
            return;
        }
        if (!panSession || ev.pointerId !== panSession.pointerId) return;
        state.viewTransform.x = panSession.tx + (ev.clientX - panSession.x);
        state.viewTransform.y = panSession.ty + (ev.clientY - panSession.y);
        applyCanvasTransform(state);
    });

    const endPan = (ev) => {
        if (blockDrag && ev.pointerId === blockDrag.pointerId) {
            const drag = blockDrag;
            blockDrag = null;
            const blockEl = root.querySelector(`[data-dl-block="${CSS.escape(drag.id)}"]`);
            if (blockEl) blockEl.classList.remove('dl-block--dragging');
            try {
                if (blockEl) blockEl.releasePointerCapture(ev.pointerId);
            } catch {
                /* ignore */
            }
            saveDatenlandkarteLayout(state.layout.blocks);
            state.savedLayout = loadDatenlandkarteLayout();
            state.layout.customized = true;
            if (!drag.moved) selectBlock(drag.id);
            return;
        }
        if (!panSession || ev.pointerId !== panSession.pointerId) return;
        const host = canvasHostEl();
        panSession = null;
        if (host) host.classList.remove('dl-panning');
        try {
            ev.target.releasePointerCapture(ev.pointerId);
        } catch {
            /* ignore */
        }
    };
    root.addEventListener('pointerup', endPan);
    root.addEventListener('pointercancel', endPan);
}

function canvasHostEl() {
    return root ? root.querySelector('#dlCanvasHost') : document.getElementById('dlCanvasHost');
}

function fitView() {
    const host = canvasHostEl();
    const inner = root ? root.querySelector('#dlCanvasInner') : document.getElementById('dlCanvasInner');
    if (!host || !inner) return;
    const cw = host.clientWidth - 40;
    const ch = host.clientHeight - 40;
    const iw = state.layout.width;
    const ih = state.layout.height;
    const k = Math.min(cw / iw, ch / ih, 1.1);
    state.viewTransform = {
        k,
        x: (host.clientWidth - iw * k) / 2,
        y: 24
    };
    applyCanvasTransform(state);
}

let globalListenersBound = false;

function bindGlobalListeners() {
    if (globalListenersBound) return;
    globalListenersBound = true;
    window.addEventListener('resize', () => fitView());
    window.addEventListener('ms365-spo-sync-status', () => {
        refreshLocalData();
        paint();
    });
    window.addEventListener('ms365-tenant-settings-changed', () => {
        refreshLocalData();
        paint();
    });
}

/**
 * @param {HTMLElement} container
 * @param {{ embedded?: boolean }} [options]
 */
export function mountDatenlandkarte(container, options) {
    if (!container) return;
    const embedded = !!(options && options.embedded);
    setDatenlandkarteHrefContext(embedded ? 'register' : 'tool');
    state.embedded = embedded;
    root = container;
    root.classList.add(embedded ? 'dl-embed' : 'dl-app-root');
    bindUiOnce();
    bindGlobalListeners();
    refreshAll();
    requestAnimationFrame(() => fitView());
    window.ms365DatenlandkarteFit = () => fitView();
    window.ms365DatenlandkarteRefresh = () => refreshAll();
}

function initStandalone() {
    const el = document.getElementById('dlApp');
    if (!el) return;
    mountDatenlandkarte(el, { embedded: false });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initStandalone);
} else {
    initStandalone();
}
