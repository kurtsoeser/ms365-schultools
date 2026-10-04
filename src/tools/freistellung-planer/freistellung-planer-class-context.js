/**
 * Klassenteams / Setup aus ms365AppDataV2 (für Klasse + KV).
 */
function appDataApi() {
    try {
        if (typeof window !== 'undefined' && window.ms365AppDataV2) return window.ms365AppDataV2;
        if (typeof globalThis !== 'undefined' && globalThis.ms365AppDataV2) return globalThis.ms365AppDataV2;
    } catch {
        /* ignore */
    }
    return null;
}

export function loadClassTeamsContext() {
    try {
        const api = appDataApi();
        if (!api || typeof api.getContainer !== 'function') {
            return { classTeams: [], setup: {} };
        }
        const container = api.getContainer() || {};
        const raw = (container.core && container.core.classTeams) || [];
        const classTeams =
            typeof api.normalizeCoreClassTeams === 'function'
                ? api.normalizeCoreClassTeams(raw)
                : Array.isArray(raw)
                  ? raw
                  : [];
        const setup = typeof api.getSetup === 'function' ? api.getSetup() || {} : {};
        return { classTeams, setup };
    } catch {
        return { classTeams: [], setup: {} };
    }
}

/**
 * Klassenzeilen aus aktuellem State + allen Schuljahr-Buckets.
 * @param {{ stammdaten?: { classes?: object[] } }} state
 */
export function collectAllClassRows(state) {
    const rows = [];
    const seen = new Set();
    const push = (cl) => {
        if (!cl || typeof cl !== 'object') return;
        const code = String(cl.code || cl.name || '').trim();
        const key = code.toLowerCase();
        if (!code || seen.has(key)) return;
        seen.add(key);
        rows.push(cl);
    };
    (state && state.stammdaten && state.stammdaten.classes ? state.stammdaten.classes : []).forEach(push);
    try {
        const api = appDataApi();
        const c = api && typeof api.getContainer === 'function' ? api.getContainer() : null;
        const by = c && c.years && c.years.byLabel && typeof c.years.byLabel === 'object' ? c.years.byLabel : null;
        if (by) {
            Object.keys(by).forEach((lab) => {
                const bucket = by[lab];
                (bucket && bucket.classes ? bucket.classes : []).forEach(push);
            });
        }
    } catch {
        /* ignore */
    }
    return rows;
}
