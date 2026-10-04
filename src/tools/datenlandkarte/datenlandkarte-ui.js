import { DATEN_LINKS } from './datenlandkarte-catalog.js';
import { formatBlockCount } from './datenlandkarte-metrics.js';
import { layoutDatenlandkarte, linksForSelection } from './datenlandkarte-layout.js';

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
 * @param {object} state
 * @param {HTMLElement} root
 */
export function renderDatenlandkarteApp(state, root) {
    const layout = state.layout || layoutDatenlandkarte();
    const metrics = state.metrics || {};
    const sel = state.selectedBlockId || '';
    const selLink = state.selectedLinkId || '';

    const linkLines = (sel ? linksForSelection(layout, sel, selLink) : layout.links)
        .map(
            (l) =>
                `<path class="dl-link${sel && (l.from === sel || l.to === sel) ? ' dl-link--hi' : ''}" data-dl-link="${escapeAttr(l.id)}" d="${l.path}" marker-end="url(#dlArrow)"/>`
        )
        .join('');

    const linkLabels = layout.links
        .filter((l) => !sel || l.from === sel || l.to === sel)
        .slice(0, 40)
        .map((l) => {
            const active = selLink === l.id ? ' dl-link-label--sel' : '';
            return `<button type="button" class="dl-link-label${active}" data-dl-link="${escapeAttr(l.id)}">${escapeHtml(l.label)}<span class="muted"> · ${escapeHtml(shortId(l.from))} → ${escapeHtml(shortId(l.to))}</span></button>`;
        })
        .join('');

    const customized = !!layout.customized;
    let lastLayer = '';
    const blockHtml = layout.blocks
        .map((b) => {
            const count = formatBlockCount(b.countKey, metrics);
            const isSel = sel === b.id;
            const layerHead =
                !customized &&
                b.layerLabel !== lastLayer
                    ? ((lastLayer = b.layerLabel), `<div class="dl-layer-label" style="left:${b.x}px;top:${b.y - 22}px">${escapeHtml(b.layerLabel)}</div>`)
                    : '';
            return `${layerHead}
<article class="dl-block${isSel ? ' dl-block--sel' : ''}" data-dl-block="${escapeAttr(b.id)}" style="left:${b.x}px;top:${b.y}px;width:${b.w}px;min-height:${b.h}px">
  <header class="dl-block__head" title="Ziehen zum Verschieben"><i class="bi bi-grip-vertical dl-block__grip" aria-hidden="true"></i><i class="bi ${escapeAttr(b.icon)}"></i><h3>${escapeHtml(b.title)}</h3></header>
  <p class="dl-block__count" title="${escapeAttr(count.hint)}"><span>${escapeHtml(count.text)}</span>${escapeHtml(count.unit || (typeof metrics[b.countKey]?.value === 'number' ? ' Einträge' : ''))}</p>
  <p class="dl-block__desc">${escapeHtml(b.description)}</p>
  ${b.href ? `<a class="dl-block__link" href="${escapeAttr(b.href)}"><i class="bi bi-box-arrow-up-right"></i>Öffnen</a>` : ''}
</article>`;
        })
        .join('');

    const detail = sel ? renderBlockDetail(layout, sel, metrics) : renderOverview(layout, metrics, state);
    const spoBanner = renderSpoStatus(state);

    root.innerHTML = `
<section class="tm-hero dl-hero">
  <div class="tm-hero__icon" aria-hidden="true"><i class="bi bi-grid-3x3-gap"></i></div>
  <div>
    <p class="tm-hero__kicker">Listen &amp; Datenquellen</p>
    <h2>Datenlandkarte</h2>
    <p>Blöcke = Listen oder Quellen · Linien = logische Verknüpfung. Lokale Zahlen aus diesem Browser; SharePoint-Blöcke mit Graph-<code>$count</code> nach Anmeldung.</p>
  </div>
  <div class="tm-hero__actions">
    <button type="button" class="btn btn-success" id="dlBtnReload"><i class="bi bi-arrow-clockwise"></i>Aktualisieren</button>
    <button type="button" class="btn" id="dlBtnFit"><i class="bi bi-aspect-ratio"></i>Ansicht zentrieren</button>
    <button type="button" class="btn btn-ghost" id="dlBtnClearSel"${sel ? '' : ' disabled'}>Auswahl aufheben</button>
    <button type="button" class="btn btn-ghost" id="dlBtnResetLayout"${layout.customized ? '' : ' disabled'} title="Gespeicherte Blockpositionen löschen"><i class="bi bi-layout-three-columns"></i>Standardlayout</button>
  </div>
</section>
${spoBanner}
<div class="dl-layout">
  <aside class="tm-panel dl-sidebar" aria-label="Details">
    <div id="dlDetail">${detail}</div>
    <hr class="dl-hr"/>
    <h4>Verbindungen</h4>
    <div class="dl-link-list">${linkLabels || '<p class="muted">Block wählen, um passende Linien zu sehen.</p>'}</div>
  </aside>
  <div class="dl-canvas-wrap tm-panel">
    <div class="dl-canvas-toolbar muted">Kopfzeile eines Blocks ziehen = verschieben · Leerer Bereich = Schwenken · Mausrad = Zoom · Anordnung wird in diesem Browser gespeichert</div>
    <div class="dl-canvas-host" id="dlCanvasHost">
      <div class="dl-canvas-inner" id="dlCanvasInner" style="width:${layout.width}px;height:${layout.height}px">
        <svg class="dl-svg" width="${layout.width}" height="${layout.height}" aria-hidden="true">
          <defs>
            <marker id="dlArrow" markerWidth="8" markerHeight="8" refX="6" refY="4" orient="auto">
              <path d="M0,0 L8,4 L0,8 z" fill="color-mix(in srgb, var(--text-secondary) 70%, transparent)"/>
            </marker>
          </defs>
          ${linkLines}
        </svg>
        <div class="dl-blocks">${blockHtml}</div>
      </div>
    </div>
  </div>
</div>`;
}

function shortId(blockId) {
    return String(blockId || '').replace(/^stamm-|^year-|^spo-|^plan-|^m365-/, '');
}

function renderOverview(layout, metrics, state) {
    const nBlocks = layout.blocks.length;
    const nLinks = layout.links.length;
    const spo = state && state.spoStatus;
    const spoNote =
        spo === 'loading'
            ? '<p class="muted"><i class="bi bi-hourglass-split"></i> SharePoint-Zähler werden geladen …</p>'
            : spo === 'ok'
              ? '<p class="muted"><i class="bi bi-check2-circle"></i> SharePoint-Zähler aktualisiert.</p>'
              : spo === 'partial' || spo === 'error'
                ? '<p class="muted"><i class="bi bi-exclamation-triangle"></i> SharePoint teilweise – Details oben.</p>'
                : '';
    return `
    <h3>Überblick</h3>
    <p class="muted">${nBlocks} Datenblöcke · ${nLinks} dokumentierte Verknüpfungen</p>
    ${spoNote}
    <p class="muted">Block-Kopfzeile ziehen, um die Anordnung anzupassen. Klick auf den Block (außer Kopfzeile) wählt ihn aus.</p>`;
}

function renderSpoStatus(state) {
    const status = state?.spoStatus || 'idle';
    const errors = Array.isArray(state?.spoErrors) ? state.spoErrors : [];
    if (status === 'idle' || (status === 'ok' && !errors.length)) return '';
    if (status === 'loading') {
        return `<p class="dl-spo-status dl-spo-status--load" role="status"><i class="bi bi-cloud-arrow-down"></i> SharePoint-Listen werden gezählt … (Microsoft-Anmeldung ggf. bestätigen)</p>`;
    }
    const cls = status === 'error' ? 'dl-spo-status--err' : 'dl-spo-status--warn';
    const head =
        status === 'error'
            ? 'SharePoint-Zähler nicht geladen'
            : 'SharePoint-Zähler mit Hinweisen';
    const list = errors.length
        ? `<ul class="dl-spo-status__list">${errors
              .slice(0, 6)
              .map((e) => `<li>${escapeHtml(e)}</li>`)
              .join('')}</ul>`
        : '';
    return `<div class="dl-spo-status ${cls}" role="status"><strong>${escapeHtml(head)}</strong>${list}</div>`;
}

function renderBlockDetail(layout, blockId, metrics) {
    const b = layout.blocks.find((x) => x.id === blockId);
    if (!b) return '';
    const count = formatBlockCount(b.countKey, metrics);
    const links = DATEN_LINKS.filter((l) => l.from === blockId || l.to === blockId);
    const rows = links
        .map((l) => {
            const other = l.from === blockId ? l.to : l.from;
            const dir = l.from === blockId ? '→' : '←';
            return `<li><strong>${escapeHtml(l.label)}</strong> ${dir} ${escapeHtml(shortId(other))}${l.detail ? `<br/><span class="muted">${escapeHtml(l.detail)}</span>` : ''}</li>`;
        })
        .join('');
    return `
    <h3><i class="bi ${escapeAttr(b.icon)}"></i> ${escapeHtml(b.title)}</h3>
    <p class="dl-detail-count">${escapeHtml(count.text)}${escapeHtml(count.unit || (typeof metrics[b.countKey]?.value === 'number' ? ' Einträge' : ''))}</p>
    <p>${escapeHtml(b.description)}</p>
    <ul class="dl-detail-links">${rows || '<li class="muted">Keine dokumentierte Verknüpfung.</li>'}</ul>`;
}

export function applyCanvasTransform(state) {
    const inner = document.getElementById('dlCanvasInner');
    if (!inner) return;
    const t = state.viewTransform || { x: 0, y: 0, k: 1 };
    inner.style.transform = `translate(${t.x}px, ${t.y}px) scale(${t.k})`;
    inner.style.transformOrigin = '0 0';
}
