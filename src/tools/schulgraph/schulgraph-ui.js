/**
 * UI-Shell für die Schuldaten-Karte (Canvas + Sidebar).
 */
import { NODE_META, EDGE_LABELS } from './schulgraph-schema.js';
import { neighborhood } from './schulgraph-logic.js';
import { ASPECT_CATALOG, GRAPH_PRESETS, aspectsSummary } from './schulgraph-aspects.js';

/**
 * @param {object} state
 * @param {HTMLElement} root
 */
export function renderSchulgraphApp(state, root) {
    const g = state.graph || { nodes: [], edges: [], stats: {}, warnings: [] };
    const stats = g.stats || {};
    const klasseOptions = (state.classCodes || [])
        .map(
            (code) =>
                `<option value="${escapeAttr(code)}"${state.options.klasseFilter === code ? ' selected' : ''}>${escapeHtml(code)}</option>`
        )
        .join('');

    const warnings =
        (g.warnings || []).length > 0
            ? `<ul class="sg-warnings">${g.warnings.map((w) => `<li>${escapeHtml(w)}</li>`).join('')}</ul>`
            : '';

    const selected = state.selectedNodeId ? g.nodes.find((n) => n.id === state.selectedNodeId) : null;
    const detail = selected ? renderNodeDetail(selected, g.edges, state) : renderOverview(stats, g);

    const preset = state.options.preset || 'overview';
    const aspects = state.options.aspects || {};
    const presetOptions = Object.entries(GRAPH_PRESETS)
        .filter(([id]) => id !== 'custom')
        .map(
            ([id, p]) =>
                `<option value="${escapeAttr(id)}"${preset === id ? ' selected' : ''}>${escapeHtml(p.label)}</option>`
        )
        .join('');
    const aspectChecks = ASPECT_CATALOG.map(
        ({ id, label, hint }) => `
      <label class="sg-check sg-check--aspect" title="${escapeAttr(hint)}">
        <input type="checkbox" id="sgAspect_${escapeAttr(id)}" data-sg-aspect="${escapeAttr(id)}" ${aspects[id] ? 'checked' : ''}/>
        <span>${escapeHtml(label)}</span>
      </label>`
    ).join('');
    const presetHint = GRAPH_PRESETS[preset]?.description || GRAPH_PRESETS.custom.description;

    root.innerHTML = `
<section class="tm-hero sg-hero">
  <div class="tm-hero__icon" aria-hidden="true"><i class="bi bi-share"></i></div>
  <div>
    <p class="tm-hero__kicker">Stammdaten &amp; Relationen</p>
    <h2>Schuldaten-Karte</h2>
    <p>Zoombarer Graph: <strong>Klick auf Person, Gruppe oder Klasse</strong> lädt Nachbarn (M365-Mitgliedschaften bzw. Mitglieder) – plus Stammdaten-Beziehungen nach Aspekt.</p>
  </div>
  <div class="tm-hero__actions">
    <button type="button" class="btn btn-success" id="sgBtnReload"><i class="bi bi-arrow-clockwise"></i>Graph neu laden</button>
    <button type="button" class="btn" id="sgBtnFit"><i class="bi bi-aspect-ratio"></i>Ansicht zentrieren</button>
    <button type="button" class="btn btn-ghost" id="sgBtnClearExplore"${state.expansion && state.expansion.anchorId ? '' : ' disabled'}>Nachbarschaft schließen</button>
    <button type="button" class="btn btn-ghost" id="sgBtnClearFocus"${state.focusNodeId ? '' : ' disabled'}>Fokus aufheben</button>
  </div>
</section>

<div class="sg-layout">
  <aside class="sg-sidebar tm-panel" aria-label="Filter und Details">
    <h3><i class="bi bi-sliders"></i>Aspekte &amp; Filter</h3>
    <div class="sg-filters">
      <div class="tm-field">
        <label for="sgPreset">Ansicht (Preset)</label>
        <select id="sgPreset">
          ${presetOptions}
          <option value="custom"${preset === 'custom' ? ' selected' : ''}>Frei kombinieren…</option>
        </select>
        <p class="muted sg-preset-hint" id="sgPresetHint">${escapeHtml(presetHint)}</p>
      </div>
      <fieldset class="sg-aspect-fieldset">
        <legend>Beziehungs-Aspekte</legend>
        ${aspectChecks}
      </fieldset>
      <div class="tm-field">
        <label for="sgKlasse">Klasse (optional)</label>
        <select id="sgKlasse">
          <option value="">Alle Klassen</option>
          ${klasseOptions}
        </select>
      </div>
    </div>
    ${warnings}
    <hr class="sg-hr"/>
    <div id="sgDetail">${detail}</div>
    <hr class="sg-hr"/>
    <h4 class="sg-legend-title">Legende</h4>
    <ul class="sg-legend">${legendHtml(stats.byKind || {})}</ul>
    <p class="muted sg-meta">${stats.nodes || 0} Knoten · ${stats.edges || 0} Verbindungen · ${escapeHtml(aspectsSummary(aspects))}</p>
  </aside>
  <div class="sg-canvas-wrap tm-panel">
    <div class="sg-canvas-toolbar">
      <span class="muted">Klick: M365-Nachbarschaft · Doppelklick: Fokus · Mausrad: Zoom</span>
      <div class="sg-zoom-btns">
        <button type="button" class="btn btn-sm" id="sgZoomIn" title="Vergrößern">+</button>
        <button type="button" class="btn btn-sm" id="sgZoomOut" title="Verkleinern">−</button>
      </div>
    </div>
    <div class="sg-canvas-host" id="sgCanvasHost">
      <svg id="sgSvg" class="sg-svg" role="img" aria-label="Interaktiver Beziehungsgraph">
        <g id="sgViewport"></g>
      </svg>
    </div>
  </div>
</div>`;
}

function legendHtml(byKind) {
    const kinds = Object.keys(byKind || {}).filter((k) => k !== 'school' && byKind[k] > 0);
    const list = kinds.length ? kinds : Object.keys(NODE_META).filter((k) => k !== 'school');
    return list
        .map((kind) => {
            const meta = NODE_META[kind] || { label: kind };
            return `<li><span class="sg-legend-dot" data-kind="${escapeAttr(kind)}"></span>${escapeHtml(meta.label)}</li>`;
        })
        .join('');
}

function renderOverview(stats, g) {
    const kinds = stats.byKind || {};
    const rows = Object.entries(kinds)
        .map(([k, n]) => {
            const label = NODE_META[k]?.label || k;
            return `<tr><td>${escapeHtml(label)}</td><td>${n}</td></tr>`;
        })
        .join('');
    return `
    <h3>Überblick</h3>
    <p class="muted">Klicken Sie einen Knoten im Graph, um Nachbarn und Beziehungen zu sehen.</p>
    <table class="sg-mini-table"><tbody>${rows || '<tr><td colspan="2" class="muted">Keine Daten – Stammdaten pflegen.</td></tr>'}</tbody></table>
    <p class="muted sg-hint">Quellen: Tenant-Stammdaten, Schuljahr-Bucket (Schüler/Eltern), Unterrichtsbelegung, Einrichtungs-<code>catalogLinks</code>.</p>`;
}

export function renderNodeDetail(node, edges, state) {
    const meta = NODE_META[node.kind] || { label: node.kind, icon: 'bi-circle' };
    const hood = neighborhood(node.id, edges);
    const linked = (edges || []).filter((e) => e.source === node.id || e.target === node.id);
    const exp = state.expansion || {};
    const exploreBlock = exp.loading
        ? '<p class="sg-explore-status"><i class="bi bi-hourglass-split"></i> Lade M365-Nachbarschaft …</p>'
        : exp.error
          ? `<p class="sg-explore-status sg-explore-status--err">${escapeHtml(exp.error)}</p>`
          : exp.anchorId === node.id && exp.summary
            ? `<p class="sg-explore-status"><i class="bi bi-cloud-check"></i> ${escapeHtml(exp.summary)}</p>`
            : '';
    const edgeRows = linked
        .slice(0, 24)
        .map((e) => {
            const other = e.source === node.id ? e.target : e.source;
            const dir = e.source === node.id ? '→' : '←';
            const kindLabel = EDGE_LABELS[e.kind] || e.kind;
            return `<li><span class="muted">${escapeHtml(kindLabel)}</span> ${dir} <code>${escapeHtml(shortId(other))}</code></li>`;
        })
        .join('');

    return `
    <h3><i class="bi ${meta.icon}"></i> ${escapeHtml(node.label)}</h3>
    ${node.sublabel ? `<p class="muted">${escapeHtml(node.sublabel)}</p>` : ''}
    <p><span class="sg-badge">${escapeHtml(meta.label)}</span> · ${hood.size - 1} direkte Nachbarn</p>
    ${exploreBlock}
    <ul class="sg-edges">${edgeRows || '<li class="muted">Keine Kanten in diesem Layer.</li>'}</ul>
    ${state.focusNodeId === node.id ? '<p class="sg-focus-hint"><i class="bi bi-bullseye"></i> Fokus aktiv – nur Nachbarschaft sichtbar.</p>' : ''}`;
}

function shortId(id) {
    const s = String(id || '');
    const idx = s.indexOf(':');
    return idx >= 0 ? s.slice(idx + 1) : s;
}

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function escapeAttr(s) {
    return escapeHtml(s).replace(/'/g, '&#39;');
}

/**
 * SVG-Graph zeichnen.
 * @param {object} state
 */
export function paintSchulgraphCanvas(state) {
    const host = document.getElementById('sgCanvasHost');
    const svg = document.getElementById('sgSvg');
    const viewport = document.getElementById('sgViewport');
    if (!host || !svg || !viewport) return;

    const g = state.displayGraph || state.graph || { nodes: [], edges: [] };
    const positions = state.positions || new Map();
    const w = host.clientWidth || 800;
    const h = Math.max(420, host.clientHeight || 520);

    svg.setAttribute('viewBox', `0 0 ${w} ${h}`);
    svg.setAttribute('width', String(w));
    svg.setAttribute('height', String(h));

    const focus = state.focusNodeId;
    const visibleNodes = focus
        ? g.nodes.filter((n) => neighborhood(focus, g.edges).has(n.id))
        : g.nodes;
    const visibleIds = new Set(visibleNodes.map((n) => n.id));
    const visibleEdges = (g.edges || []).filter((e) => visibleIds.has(e.source) && visibleIds.has(e.target));

    const selected = state.selectedNodeId;
    const selectedHood = selected ? neighborhood(selected, g.edges) : null;

    let html = '';
    visibleEdges.forEach((e) => {
        const a = positions.get(e.source);
        const b = positions.get(e.target);
        if (!a || !b) return;
        const hi = selected && (e.source === selected || e.target === selected);
        html += `<line class="sg-edge${hi ? ' sg-edge--hi' : ''}" data-kind="${escapeAttr(e.kind)}" x1="${a.x}" y1="${a.y}" x2="${b.x}" y2="${b.y}"/>`;
    });

    visibleNodes.forEach((n) => {
        const p = positions.get(n.id);
        if (!p) return;
        const meta = NODE_META[n.kind] || NODE_META.class;
        const r = n.kind === 'school' ? 22 : n.kind === 'm365group' ? 11 : 14;
        const isSel = selected === n.id;
        const isDim = selected && selectedHood && !selectedHood.has(n.id);
        const label = n.label.length > 18 ? n.label.slice(0, 16) + '…' : n.label;
        html += `
<g class="sg-node${isSel ? ' sg-node--sel' : ''}${isDim ? ' sg-node--dim' : ''}" data-node-id="${escapeAttr(n.id)}" transform="translate(${p.x},${p.y})" tabindex="0" role="button" aria-label="${escapeAttr(n.label)}">
  <circle class="sg-node__circle" data-kind="${escapeAttr(n.kind)}" r="${r}"/>
  <text class="sg-node__label" y="${r + 14}" text-anchor="middle">${escapeHtml(label)}</text>
</g>`;
    });

    viewport.innerHTML = html;

    applyViewportTransform(state, svg, viewport);
}

/**
 * @param {object} state
 * @param {SVGSVGElement} svg
 * @param {SVGGElement} viewport
 */
export function applyViewportTransform(state, svg, viewport) {
    if (!svg || !viewport) return;
    const t = state.viewTransform || { x: 0, y: 0, k: 1 };
    viewport.setAttribute('transform', `translate(${t.x},${t.y}) scale(${t.k})`);
}

export function readFiltersFromDom() {
    const presetEl = document.getElementById('sgPreset');
    const preset = presetEl ? String(presetEl.value || 'overview').trim() : 'overview';

    /** @type {Record<string, boolean>} */
    const aspectsRaw = {};
    document.querySelectorAll('[data-sg-aspect]').forEach((el) => {
        if (!(el instanceof HTMLInputElement)) return;
        const key = el.getAttribute('data-sg-aspect');
        if (key) aspectsRaw[key] = el.checked;
    });

    const klasseFilter = String(document.getElementById('sgKlasse')?.value || '').trim();
    return { preset, aspects: aspectsRaw, klasseFilter };
}
