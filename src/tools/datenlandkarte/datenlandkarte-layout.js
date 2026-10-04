/**
 * Layout: Blöcke pro Schicht, Kanten als Polylinien-Koordinaten.
 */
import { DATEN_BLOECKE, DATEN_LINKS, LAYER_LABELS, blocksInLayer } from './datenlandkarte-catalog.js';

const BLOCK_W = 200;
const BLOCK_H = 112;
const GAP_X = 48;
const GAP_Y = 36;
const LAYER_PAD = 24;

export { BLOCK_W, BLOCK_H };

/**
 * @param {Array<object>} placedBlocks
 * @param {Array<object>} links
 */
export function recomputeLayoutFromBlocks(placedBlocks, links) {
    /** @type {Map<string, { x: number, y: number, w: number, h: number }>} */
    const pos = new Map();
    (placedBlocks || []).forEach((b) => {
        if (!b || !b.id) return;
        pos.set(b.id, {
            x: b.x,
            y: b.y,
            w: b.w || BLOCK_W,
            h: b.h || BLOCK_H
        });
    });

    const linkGeom = (links || [])
        .map((link) => {
            const a = pos.get(link.from);
            const b = pos.get(link.to);
            if (!a || !b) return null;
            const x1 = a.x + a.w / 2;
            const y1 = a.y + a.h;
            const x2 = b.x + b.w / 2;
            const y2 = b.y;
            const midY = (y1 + y2) / 2;
            return {
                ...link,
                path: `M ${x1} ${y1} C ${x1} ${midY}, ${x2} ${midY}, ${x2} ${y2}`
            };
        })
        .filter(Boolean);

    let maxX = 800;
    let maxY = LAYER_PAD;
    pos.forEach((p) => {
        maxX = Math.max(maxX, p.x + p.w + LAYER_PAD);
        maxY = Math.max(maxY, p.y + p.h + LAYER_PAD);
    });

    return {
        blocks: placedBlocks,
        links: linkGeom,
        width: maxX,
        height: maxY + LAYER_PAD,
        layers: []
    };
}

function layoutAuto(blockDefs, linkDefs) {
    const blocks = blockDefs || DATEN_BLOECKE;
    const links = linkDefs || DATEN_LINKS;
    const layerOrder = ['stamm', 'schuljahr', 'sharepoint', 'planer', 'm365'];

    /** @type {Map<string, { x: number, y: number, w: number, h: number }>} */
    const pos = new Map();
    const placed = [];
    let yCursor = LAYER_PAD;
    let maxX = 800;

    layerOrder.forEach((layer) => {
        const row = blocksInLayer(layer, blocks);
        if (!row.length) return;
        const cols = 5;
        const rows = Math.ceil(row.length / cols);
        const rowW = Math.min(row.length, cols) * BLOCK_W + (Math.min(row.length, cols) - 1) * GAP_X;
        maxX = Math.max(maxX, rowW + LAYER_PAD * 2);
        row.forEach((def, i) => {
            const col = i % cols;
            const r = Math.floor(i / cols);
            const x = LAYER_PAD + Math.max(0, (maxX - rowW) / 2 - LAYER_PAD) + col * (BLOCK_W + GAP_X);
            const y = yCursor + r * (BLOCK_H + GAP_Y);
            pos.set(def.id, { x, y, w: BLOCK_W, h: BLOCK_H });
            placed.push({ ...def, x, y, w: BLOCK_W, h: BLOCK_H, layerLabel: LAYER_LABELS[layer] || layer });
        });
        yCursor += rows * (BLOCK_H + GAP_Y) + 28;
    });

    return recomputeLayoutFromBlocks(placed, links);
}

/**
 * @param {Array<object>|undefined} blockDefs
 * @param {Array<object>|undefined} linkDefs
 * @param {Record<string, { x: number, y: number }>|undefined} savedPositions
 * @returns {{ blocks: Array<object>, links: Array<object>, width: number, height: number, layers: string[], customized?: boolean }}
 */
export function layoutDatenlandkarte(blockDefs, linkDefs, savedPositions) {
    const auto = layoutAuto(blockDefs, linkDefs);
    const saved = savedPositions && typeof savedPositions === 'object' ? savedPositions : {};
    const keys = Object.keys(saved);
    if (!keys.length) {
        return { ...auto, customized: false };
    }

    const blocks = auto.blocks.map((b) => {
        const s = saved[b.id];
        if (s && Number.isFinite(s.x) && Number.isFinite(s.y)) {
            return { ...b, x: s.x, y: s.y };
        }
        return b;
    });

    const layout = recomputeLayoutFromBlocks(blocks, linkDefs || DATEN_LINKS);
    return { ...layout, customized: true };
}

/**
 * @param {object} layout
 * @param {string|null} selectedId
 * @param {string|null} hoverLinkId
 */
export function linksForSelection(layout, selectedId, hoverLinkId) {
    if (!selectedId) return layout.links;
    return layout.links.filter((l) => l.from === selectedId || l.to === selectedId || l.id === hoverLinkId);
}
