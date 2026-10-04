const STORAGE_KEY = 'ms365-datenlandkarte-layout-v1';

/**
 * @returns {Record<string, { x: number, y: number }>}
 */
export function loadDatenlandkarteLayout() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return {};
        const data = JSON.parse(raw);
        if (!data || typeof data !== 'object') return {};
        /** @type {Record<string, { x: number, y: number }>} */
        const out = {};
        Object.keys(data).forEach((id) => {
            const p = data[id];
            if (!p || typeof p !== 'object') return;
            const x = Number(p.x);
            const y = Number(p.y);
            if (Number.isFinite(x) && Number.isFinite(y)) out[id] = { x, y };
        });
        return out;
    } catch {
        return {};
    }
}

/**
 * @param {Array<{ id: string, x: number, y: number }>} blocks
 */
export function saveDatenlandkarteLayout(blocks) {
    /** @type {Record<string, { x: number, y: number }>} */
    const data = {};
    (blocks || []).forEach((b) => {
        if (!b || !b.id) return;
        if (!Number.isFinite(b.x) || !Number.isFinite(b.y)) return;
        data[b.id] = { x: Math.round(b.x), y: Math.round(b.y) };
    });
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    } catch {
        /* ignore */
    }
}

export function clearDatenlandkarteLayout() {
    try {
        localStorage.removeItem(STORAGE_KEY);
    } catch {
        /* ignore */
    }
}

export function hasCustomDatenlandkarteLayout(saved) {
    return !!(saved && typeof saved === 'object' && Object.keys(saved).length > 0);
}
