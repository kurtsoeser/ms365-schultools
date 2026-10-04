/**
 * Schuldaten-Karte – interaktiver Beziehungsgraph (Stammdaten).
 */
import { buildSchulGraph, layoutClusterGraph, normalizeGraphOptions, aspectsForPreset } from './schulgraph-logic.js';
import { ASPECT_CATALOG, GRAPH_PRESETS, detectPreset } from './schulgraph-aspects.js';
import { mergeGraphLayers, placeNodesAround } from './schulgraph-explore-logic.js';
import { expandNodeNeighborhood, catalogGroupIdSet } from './schulgraph-explore-api.js';
import {
    renderSchulgraphApp,
    paintSchulgraphCanvas,
    applyViewportTransform,
    readFiltersFromDom,
    renderNodeDetail
} from './schulgraph-ui.js';

const OPTIONS_KEY = 'ms365-schulgraph-options-v2';

/** @type {object} */
const state = {
    graph: { nodes: [], edges: [], stats: {}, warnings: [] },
    positions: new Map(),
    classCodes: [],
    options: normalizeGraphOptions(loadOptions()),
    selectedNodeId: '',
    focusNodeId: '',
    viewTransform: { x: 0, y: 0, k: 1 },
    displayGraph: { nodes: [], edges: [] },
    expansion: {
        anchorId: '',
        nodes: [],
        edges: [],
        loading: false,
        error: '',
        summary: '',
        warnings: []
    },
    exploreSources: { settings: {}, yearBucket: {}, setup: {} }
};

let root = null;
/** @type {{ x: number, y: number, tx: number, ty: number, pointerId: number } | null} */
let panSession = null;

function loadOptions() {
    try {
        const raw = localStorage.getItem(OPTIONS_KEY);
        if (raw) return JSON.parse(raw);
        const legacy = localStorage.getItem('ms365-schulgraph-options-v1');
        if (legacy) return JSON.parse(legacy);
    } catch {
        /* ignore */
    }
    return null;
}

function persistOptions() {
    try {
        localStorage.setItem(OPTIONS_KEY, JSON.stringify(state.options));
    } catch {
        /* ignore */
    }
}

function loadSources() {
    let settings = null;
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            settings = window.ms365TenantSettingsLoad();
        }
    } catch {
        settings = null;
    }

    let yearBucket = {};
    let setup = {};
    try {
        if (window.ms365AppDataV2) {
            if (typeof window.ms365AppDataV2.getSetup === 'function') {
                setup = window.ms365AppDataV2.getSetup() || {};
            }
            if (typeof window.ms365AppDataV2.getYearBucket === 'function') {
                const yb = window.ms365AppDataV2.getYearBucket();
                yearBucket = yb && yb.bucket ? yb.bucket : yb || {};
            }
        }
    } catch {
        /* ignore */
    }

    return { settings: settings || {}, yearBucket, setup };
}

function updateDisplayGraph() {
    const extra =
        state.expansion && (state.expansion.nodes.length || state.expansion.loading)
            ? { nodes: state.expansion.nodes, edges: state.expansion.edges }
            : { nodes: [], edges: [] };
    state.displayGraph = mergeGraphLayers(state.graph, extra);
}

function findNodeById(id) {
    return (
        state.displayGraph.nodes.find((n) => n.id === id) ||
        state.graph.nodes.find((n) => n.id === id) ||
        state.expansion.nodes.find((n) => n.id === id)
    );
}

function clearExpansion() {
    state.expansion = {
        anchorId: '',
        nodes: [],
        edges: [],
        loading: false,
        error: '',
        summary: '',
        warnings: []
    };
    updateDisplayGraph();
}

function exploreContext() {
    return {
        settings: state.exploreSources.settings,
        yearBucket: state.exploreSources.yearBucket,
        setup: state.exploreSources.setup,
        catalogGroupIds: catalogGroupIdSet(state.exploreSources.setup)
    };
}

function rebuildGraph() {
    const src = loadSources();
    state.classCodes = (src.settings.classes || [])
        .map((c) => String(c.code || '').trim().toUpperCase())
        .filter(Boolean)
        .sort();

    state.graph = buildSchulGraph({
        settings: src.settings,
        yearBucket: src.yearBucket,
        setup: src.setup,
        options: state.options
    });
    state.exploreSources = src;
    clearExpansion();

    const host = document.getElementById('sgCanvasHost');
    const w = host && host.clientWidth ? host.clientWidth : 900;
    const h = host && host.clientHeight ? host.clientHeight : 560;
    state.positions = layoutClusterGraph(state.graph.nodes, state.graph.edges, { width: w, height: h });
    updateDisplayGraph();
}

function refreshCanvasOnly() {
    paintSchulgraphCanvas(state);
}

function refreshAll() {
    if (!root) return;
    renderSchulgraphApp(state, root);
    refreshCanvasOnly();
}

function applyOptionsFromDom() {
    const raw = readFiltersFromDom();
    let preset = String(raw.preset || 'overview').trim();
    let aspects = raw.aspects;
    if (preset !== 'custom') {
        aspects = aspectsForPreset(preset);
    } else {
        aspects = normalizeGraphOptions({ aspects: raw.aspects }).aspects;
        preset = 'custom';
    }
    state.options = normalizeGraphOptions({
        preset,
        aspects,
        klasseFilter: raw.klasseFilter,
        peopleAutoOffThreshold: state.options.peopleAutoOffThreshold,
        maxStudents: state.options.maxStudents,
        maxGuardians: state.options.maxGuardians
    });
    persistOptions();
    state.selectedNodeId = '';
    state.focusNodeId = '';
    clearExpansion();
    rebuildGraph();
    refreshAll();
    syncAspectCheckboxesFromOptions();
    requestAnimationFrame(() => fitView());
}

function syncAspectCheckboxesFromOptions() {
    const asp = state.options.aspects || {};
    ASPECT_CATALOG.forEach(({ id }) => {
        const cb = document.getElementById(`sgAspect_${id}`);
        if (cb instanceof HTMLInputElement) cb.checked = !!asp[id];
    });
    const presetEl = document.getElementById('sgPreset');
    if (presetEl) {
        presetEl.value =
            state.options.preset === 'custom' ? 'custom' : detectPreset(asp) !== 'custom' ? detectPreset(asp) : 'custom';
    }
    const hint = document.getElementById('sgPresetHint');
    if (hint) {
        const key = presetEl ? presetEl.value : 'overview';
        const p = GRAPH_PRESETS[key] || GRAPH_PRESETS.custom;
        hint.textContent = p.description || '';
    }
}

function paint() {
    refreshAll();
    requestAnimationFrame(() => fitView());
}

function selectNode(id) {
    state.selectedNodeId = id;
    const node = findNodeById(id);
    if (!node) return;

    state.expansion = {
        anchorId: id,
        nodes: [],
        edges: [],
        loading: true,
        error: '',
        summary: '',
        warnings: []
    };
    updateDisplayGraph();
    refreshDetailPanel(node);
    refreshCanvasOnly();

    expandNodeNeighborhood(node, exploreContext())
        .then((result) => {
            if (state.selectedNodeId !== id) return;
            state.expansion = {
                anchorId: id,
                nodes: result.nodes || [],
                edges: result.edges || [],
                loading: false,
                error: '',
                summary: result.summary || '',
                warnings: result.warnings || []
            };
            updateDisplayGraph();
            const anchorPos = state.positions.get(id);
            if (anchorPos) {
                const newIds = state.expansion.nodes.map((n) => n.id).filter((nid) => !state.positions.has(nid));
                placeNodesAround(state.positions, anchorPos, newIds, 150);
            }
            refreshDetailPanel(findNodeById(id));
            refreshCanvasOnly();
            const btn = document.getElementById('sgBtnClearExplore');
            if (btn) btn.disabled = false;
        })
        .catch((e) => {
            if (state.selectedNodeId !== id) return;
            state.expansion = {
                anchorId: id,
                nodes: [],
                edges: [],
                loading: false,
                error: e && e.message ? e.message : String(e),
                summary: '',
                warnings: []
            };
            updateDisplayGraph();
            refreshDetailPanel(node);
            refreshCanvasOnly();
        });
}

function refreshDetailPanel(node) {
    const detail = document.getElementById('sgDetail');
    if (!detail || !node) return;
    detail.innerHTML = renderNodeDetail(node, state.displayGraph.edges, state);
}

function focusNode(id) {
    state.focusNodeId = id;
    state.selectedNodeId = id;
    refreshAll();
    requestAnimationFrame(() => fitView());
}

function bindUiOnce() {
    if (!root || root.dataset.sgUiBound === '1') return;
    root.dataset.sgUiBound = '1';

    root.addEventListener('click', (ev) => {
        const nodeEl = ev.target.closest('.sg-node');
        if (nodeEl) {
            ev.preventDefault();
            ev.stopPropagation();
            const id = nodeEl.getAttribute('data-node-id') || '';
            if (id) selectNode(id);
            return;
        }
        if (ev.target.closest('#sgBtnReload')) {
            applyOptionsFromDom();
            return;
        }
        if (ev.target.closest('#sgBtnClearExplore')) {
            clearExpansion();
            refreshCanvasOnly();
            const node = findNodeById(state.selectedNodeId);
            if (node) refreshDetailPanel(node);
            const btn = document.getElementById('sgBtnClearExplore');
            if (btn) btn.disabled = true;
            return;
        }
        if (ev.target.closest('#sgBtnClearFocus')) {
            state.focusNodeId = '';
            refreshAll();
            requestAnimationFrame(() => fitView());
            return;
        }
        if (ev.target.closest('#sgBtnFit')) {
            fitView();
            return;
        }
        if (ev.target.closest('#sgZoomIn')) {
            zoomBy(1.15);
            return;
        }
        if (ev.target.closest('#sgZoomOut')) {
            zoomBy(1 / 1.15);
        }
    });

    root.addEventListener('dblclick', (ev) => {
        const nodeEl = ev.target.closest('.sg-node');
        if (!nodeEl) return;
        ev.preventDefault();
        ev.stopPropagation();
        const id = nodeEl.getAttribute('data-node-id') || '';
        if (id) focusNode(id);
    });

    root.addEventListener('change', (ev) => {
        const t = ev.target;
        if (!(t instanceof HTMLElement)) return;
        if (t.id === 'sgPreset') {
            const val = String(t.value || 'overview');
            if (val !== 'custom') {
                const a = aspectsForPreset(val);
                ASPECT_CATALOG.forEach(({ id }) => {
                    const cb = document.getElementById(`sgAspect_${id}`);
                    if (cb instanceof HTMLInputElement) cb.checked = !!a[id];
                });
            }
            applyOptionsFromDom();
            return;
        }
        if (t.matches('[data-sg-aspect]')) {
            const presetEl = document.getElementById('sgPreset');
            if (presetEl) presetEl.value = 'custom';
            applyOptionsFromDom();
            return;
        }
        if (t.id === 'sgKlasse') {
            applyOptionsFromDom();
        }
    });

    root.addEventListener(
        'wheel',
        (ev) => {
            const host = ev.target.closest('#sgCanvasHost');
            if (!host) return;
            ev.preventDefault();
            const svg = document.getElementById('sgSvg');
            if (!svg) return;
            const rect = svg.getBoundingClientRect();
            const mx = ev.clientX - rect.left;
            const my = ev.clientY - rect.top;
            const delta = ev.deltaY < 0 ? 1.08 : 1 / 1.08;
            zoomAt(mx, my, delta);
        },
        { passive: false }
    );

    root.addEventListener('pointerdown', (ev) => {
        const host = ev.target.closest('#sgCanvasHost');
        if (!host) return;
        if (ev.target.closest('.sg-node')) return;
        panSession = {
            x: ev.clientX,
            y: ev.clientY,
            tx: state.viewTransform.x,
            ty: state.viewTransform.y,
            pointerId: ev.pointerId
        };
        host.classList.add('sg-panning');
        host.setPointerCapture(ev.pointerId);
    });

    root.addEventListener('pointermove', (ev) => {
        if (!panSession || ev.pointerId !== panSession.pointerId) return;
        const svg = document.getElementById('sgSvg');
        const viewport = document.getElementById('sgViewport');
        if (!svg || !viewport) return;
        state.viewTransform.x = panSession.tx + (ev.clientX - panSession.x);
        state.viewTransform.y = panSession.ty + (ev.clientY - panSession.y);
        applyViewportTransform(state, svg, viewport);
    });

    const endPan = (ev) => {
        if (!panSession || ev.pointerId !== panSession.pointerId) return;
        const host = document.getElementById('sgCanvasHost');
        panSession = null;
        if (host) host.classList.remove('sg-panning');
        try {
            ev.target.releasePointerCapture(ev.pointerId);
        } catch {
            /* ignore */
        }
    };
    root.addEventListener('pointerup', endPan);
    root.addEventListener('pointercancel', endPan);

    root.addEventListener('keydown', (ev) => {
        const nodeEl = ev.target.closest('.sg-node');
        if (!nodeEl) return;
        if (ev.key === 'Enter' || ev.key === ' ') {
            ev.preventDefault();
            const id = nodeEl.getAttribute('data-node-id') || '';
            if (id) selectNode(id);
        }
    });
}

function zoomBy(factor) {
    const svg = document.getElementById('sgSvg');
    if (!svg) return;
    const rect = svg.getBoundingClientRect();
    zoomAt(rect.width / 2, rect.height / 2, factor);
}

function zoomAt(mx, my, factor) {
    const svg = document.getElementById('sgSvg');
    const viewport = document.getElementById('sgViewport');
    if (!svg || !viewport) return;
    const t = state.viewTransform;
    const nx = mx - (mx - t.x) * factor;
    const ny = my - (my - t.y) * factor;
    t.k = Math.min(3, Math.max(0.25, t.k * factor));
    t.x = nx;
    t.y = ny;
    applyViewportTransform(state, svg, viewport);
}

function fitView() {
    const host = document.getElementById('sgCanvasHost');
    const svg = document.getElementById('sgSvg');
    const viewport = document.getElementById('sgViewport');
    if (!host || !svg || !viewport || !state.positions.size) return;

    let minX = Infinity;
    let minY = Infinity;
    let maxX = -Infinity;
    let maxY = -Infinity;
    state.positions.forEach((p) => {
        minX = Math.min(minX, p.x);
        minY = Math.min(minY, p.y);
        maxX = Math.max(maxX, p.x);
        maxY = Math.max(maxY, p.y);
    });
    const pad = 60;
    const bw = maxX - minX + pad * 2;
    const bh = maxY - minY + pad * 2;
    const w = host.clientWidth || 800;
    const h = host.clientHeight || 520;
    const k = Math.min(w / bw, h / bh, 1.2);
    state.viewTransform = {
        k,
        x: (w - (minX + maxX) * k) / 2,
        y: (h - (minY + maxY) * k) / 2
    };
    applyViewportTransform(state, svg, viewport);
}

function init() {
    root = document.getElementById('sgApp');
    if (!root) return;
    bindUiOnce();
    rebuildGraph();
    paint();
    window.addEventListener('resize', () => {
        rebuildGraph();
        refreshCanvasOnly();
        fitView();
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
