/**
 * Persönliches Kachel-Raster der Portal-Startseite.
 * Nur Positionen und Größen, keine Fachdaten.
 * Der gleiche Kachel-Satz kann später in den OneDrive-App-Ordner des Kontos wandern.
 */

export const PORTAL_CANVAS_COLS = 12;
export const PORTAL_CANVAS_ROW_PX = 72;
export const PORTAL_CANVAS_GAP_PX = 16;
export const PORTAL_CANVAS_MIN_W = 2;
export const PORTAL_CANVAS_MIN_H = 2;
export const PORTAL_CANVAS_MAX_H = 8;
export const PORTAL_CANVAS_LAYOUT_KEY = 'ms365-portal-canvas-layout-v2';

/**
 * @param {number} n
 * @param {number} min
 * @param {number} max
 */
function clamp(n, min, max) {
    const v = Number(n);
    if (!Number.isFinite(v)) return min;
    return Math.min(max, Math.max(min, Math.round(v)));
}

/**
 * @param {object} tile
 */
export function clampCanvasTile(tile) {
    const w = clamp(tile && tile.w, PORTAL_CANVAS_MIN_W, PORTAL_CANVAS_COLS);
    const h = clamp(tile && tile.h, PORTAL_CANVAS_MIN_H, PORTAL_CANVAS_MAX_H);
    const x = clamp(tile && tile.x, 0, PORTAL_CANVAS_COLS - w);
    const y = Math.max(0, clamp(tile && tile.y, 0, 80));
    return {
        id: String((tile && tile.id) || '').trim(),
        x,
        y,
        w,
        h
    };
}

/**
 * @param {object} a
 * @param {object} b
 */
export function canvasTilesOverlap(a, b) {
    if (!a || !b || a.id === b.id) return false;
    return a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
}

/**
 * @param {object[]} tiles
 */
export function sortCanvasTiles(tiles) {
    return [...(tiles || [])].sort((a, b) => a.y - b.y || a.x - b.x || String(a.id).localeCompare(String(b.id)));
}

/**
 * @param {string[]} ids
 * @param {Record<string, { w?: number, h?: number }>} [sizes]
 */
export function defaultCanvasTiles(ids, sizes) {
    /** @type {object[]} */
    const tiles = [];
    (ids || []).forEach((id) => {
        const key = String(id || '').trim();
        if (!key || tiles.some((t) => t.id === key)) return;
        const spec = sizes && sizes[key] ? sizes[key] : {};
        const w = clamp(spec.w || 6, PORTAL_CANVAS_MIN_W, PORTAL_CANVAS_COLS);
        const h = clamp(spec.h || 4, PORTAL_CANVAS_MIN_H, PORTAL_CANVAS_MAX_H);
        const spot = firstFreeSpot(tiles, w, h);
        tiles.push({ id: key, x: spot.x, y: spot.y, w, h });
    });
    return tiles;
}

/**
 * @param {object[]} tiles
 * @param {number} w
 * @param {number} h
 */
function firstFreeSpot(tiles, w, h) {
    const width = clamp(w, PORTAL_CANVAS_MIN_W, PORTAL_CANVAS_COLS);
    const height = clamp(h, PORTAL_CANVAS_MIN_H, PORTAL_CANVAS_MAX_H);
    for (let y = 0; y < 48; y += 1) {
        for (let x = 0; x <= PORTAL_CANVAS_COLS - width; x += 1) {
            const probe = { id: '__probe__', x, y, w: width, h: height };
            if (!(tiles || []).some((t) => canvasTilesOverlap(t, probe))) return { x, y };
        }
    }
    const bottom = (tiles || []).reduce((max, t) => Math.max(max, t.y + t.h), 0);
    return { x: 0, y: bottom };
}

/**
 * Gespeicherte Plätze behalten. Ohne strict kommen fehlende Kacheln dazu.
 * @param {string[]} ids
 * @param {object[]} saved
 * @param {{ strict?: boolean }} [opts]
 */
export function mergeCanvasTiles(ids, saved, opts) {
    const strict = !!(opts && opts.strict);
    const known = new Map();
    (saved || []).forEach((tile) => {
        const next = clampCanvasTile(tile);
        if (next.id) known.set(next.id, next);
    });
    /** @type {object[]} */
    const placed = [];
    (ids || []).forEach((id) => {
        const key = String(id || '').trim();
        if (!key || placed.some((t) => t.id === key)) return;
        if (known.has(key)) {
            placed.push(known.get(key));
            return;
        }
        if (strict) return;
        const spot = firstFreeSpot(placed, 6, 4);
        placed.push({ id: key, x: spot.x, y: spot.y, w: 6, h: 4 });
    });
    return placed;
}

/**
 * @param {object[]} tiles
 * @param {string} id
 * @param {{ w?: number, h?: number }} [size]
 */
export function addCanvasTile(tiles, id, size) {
    const key = String(id || '').trim();
    const list = (tiles || []).map((t) => clampCanvasTile(t)).filter((t) => t.id);
    if (!key || list.some((t) => t.id === key)) return list;
    const w = clamp((size && size.w) || 4, PORTAL_CANVAS_MIN_W, PORTAL_CANVAS_COLS);
    const h = clamp((size && size.h) || 3, PORTAL_CANVAS_MIN_H, PORTAL_CANVAS_MAX_H);
    const spot = firstFreeSpot(list, w, h);
    return [...list, { id: key, x: spot.x, y: spot.y, w, h }];
}

/**
 * @param {object[]} tiles
 * @param {string} id
 */
export function removeCanvasTile(tiles, id) {
    const key = String(id || '').trim();
    return (tiles || []).map((t) => clampCanvasTile(t)).filter((t) => t.id && t.id !== key);
}

/**
 * @param {object[]} tiles
 * @param {string} pinnedId
 */
function resolveOverlaps(tiles, pinnedId) {
    const list = (tiles || []).map((t) => clampCanvasTile(t)).filter((t) => t.id);
    let guard = 0;
    while (guard < 80) {
        guard += 1;
        let hit = false;
        for (let i = 0; i < list.length; i += 1) {
            for (let j = i + 1; j < list.length; j += 1) {
                if (!canvasTilesOverlap(list[i], list[j])) continue;
                hit = true;
                const pinI = list[i].id === pinnedId;
                const pinJ = list[j].id === pinnedId;
                if (pinI && !pinJ) list[j] = { ...list[j], y: list[i].y + list[i].h };
                else if (pinJ && !pinI) list[i] = { ...list[i], y: list[j].y + list[j].h };
                else if (list[i].y + list[i].h <= list[j].y + list[j].h) list[j] = { ...list[j], y: list[i].y + list[i].h };
                else list[i] = { ...list[i], y: list[j].y + list[j].h };
            }
        }
        if (!hit) break;
    }
    return list;
}

/**
 * @param {object[]} tiles
 * @param {string} id
 * @param {{ x?: number, y?: number, w?: number, h?: number }} patch
 */
export function placeCanvasTile(tiles, id, patch) {
    const key = String(id || '').trim();
    const next = (tiles || []).map((t) => clampCanvasTile(t)).filter((t) => t.id);
    const idx = next.findIndex((t) => t.id === key);
    if (idx < 0) return next;
    next[idx] = clampCanvasTile({ ...next[idx], ...(patch || {}), id: key });
    return resolveOverlaps(next, key);
}

/**
 * @param {object[]} tiles
 * @param {string} id
 * @param {number} toIndex
 */
export function reorderCanvasTiles(tiles, id, toIndex) {
    const ordered = sortCanvasTiles((tiles || []).map((t) => clampCanvasTile(t)).filter((t) => t.id));
    const from = ordered.findIndex((t) => t.id === id);
    if (from < 0) return ordered;
    const target = clamp(toIndex, 0, ordered.length - 1);
    const [item] = ordered.splice(from, 1);
    ordered.splice(target, 0, item);
    /** @type {object[]} */
    const placed = [];
    ordered.forEach((t) => {
        const spot = firstFreeSpot(placed, t.w, t.h);
        placed.push({ ...t, x: spot.x, y: spot.y });
    });
    return placed;
}

/**
 * @param {{ email?: string, tenantId?: string }} identity
 */
export function canvasLayoutUserKey(identity) {
    const id = identity || {};
    const tenant = String(id.tenantId || '').trim() || 'local';
    const email = String(id.email || '')
        .trim()
        .toLowerCase() || 'anon';
    return tenant + '|' + email;
}

/**
 * @param {string} raw
 */
function parseBag(raw) {
    try {
        const bag = JSON.parse(raw || '');
        if (!bag || typeof bag !== 'object' || !bag.users || typeof bag.users !== 'object') {
            return { v: 1, users: {} };
        }
        return { v: 1, users: bag.users };
    } catch {
        return { v: 1, users: {} };
    }
}

/**
 * @param {{ getItem: (key: string) => string|null }} storage
 * @param {{ email?: string, tenantId?: string }} identity
 */
export function readCanvasLayout(storage, identity) {
    const bag = parseBag(storage && storage.getItem(PORTAL_CANVAS_LAYOUT_KEY));
    const rec = bag.users[canvasLayoutUserKey(identity)];
    if (!rec || !Array.isArray(rec.tiles)) return [];
    return rec.tiles.map((tile) => clampCanvasTile(tile)).filter((tile) => tile.id);
}

/**
 * @param {{ getItem: (key: string) => string|null }} storage
 * @param {{ email?: string, tenantId?: string }} identity
 */
export function hasCanvasLayout(storage, identity) {
    const bag = parseBag(storage && storage.getItem(PORTAL_CANVAS_LAYOUT_KEY));
    const rec = bag.users[canvasLayoutUserKey(identity)];
    return !!(rec && Array.isArray(rec.tiles));
}

/**
 * @param {{ getItem: (key: string) => string|null, setItem: (key: string, value: string) => void }} storage
 * @param {{ email?: string, tenantId?: string, persona?: string }} identity
 * @param {object[]} tiles
 * @param {string} [updatedAt]
 */
export function writeCanvasLayout(storage, identity, tiles, updatedAt) {
    const bag = parseBag(storage && storage.getItem(PORTAL_CANVAS_LAYOUT_KEY));
    const id = identity || {};
    bag.users[canvasLayoutUserKey(id)] = {
        v: 1,
        tenantId: String(id.tenantId || '').trim(),
        persona: String(id.persona || '').trim(),
        updatedAt: updatedAt || new Date().toISOString(),
        tiles: (tiles || []).map((tile) => clampCanvasTile(tile)).filter((tile) => tile.id)
    };
    storage.setItem(PORTAL_CANVAS_LAYOUT_KEY, JSON.stringify(bag));
}
