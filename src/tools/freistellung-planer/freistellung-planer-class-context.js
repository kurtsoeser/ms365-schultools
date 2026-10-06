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
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Klassencode + Entra-Gruppe für Planer (SharePoint-Sync / Schüler ohne IT-Stammdaten).
 * @returns {{ code: string, groupId: string, name?: string }[]}
 */
export function classTeamLinksFromAppData() {
    const { classTeams, setup } = loadClassTeamsContext();
    const out = [];
    const seen = new Set();
    const push = (code, groupId, name) => {
        const c = String(code || '').trim();
        const g = String(groupId || '').trim();
        if (!c || !GUID_RE.test(g)) return;
        const key = c.toLowerCase() + '|' + g.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push({ code: c, groupId: g, name: String(name || c).trim() });
    };
    (classTeams || []).forEach((t) => {
        if (!t) return;
        push(t.classCode || t.displayName, t.graphGroupId, t.displayName || t.classCode);
    });
    const map =
        setup && setup.classGroupMatchByKey && typeof setup.classGroupMatchByKey === 'object'
            ? setup.classGroupMatchByKey
            : {};
    Object.entries(map).forEach(([key, entry]) => {
        if (!entry || typeof entry !== 'object') return;
        push(key, entry.groupId || entry.graphGroupId, entry.displayName || key);
    });
    return out;
}

/**
 * @param {{ code?: string, groupId?: string, name?: string }[]} links
 * @returns {{ code: string, name: string }[]}
 */
export function classRowsFromTeamLinks(links) {
    const rows = [];
    const seen = new Set();
    (links || []).forEach((entry) => {
        const code = String((entry && entry.code) || '').trim();
        if (!code || seen.has(code.toLowerCase())) return;
        seen.add(code.toLowerCase());
        rows.push({ code, name: String((entry && entry.name) || code).trim() });
    });
    return rows;
}

/**
 * Schülerzeilen aus Tenant-Settings und allen Schuljahr-Buckets (App-Daten).
 * @returns {object[]}
 */
export function collectAllStudentRows() {
    const rows = [];
    const seen = new Set();
    const push = (s) => {
        if (!s || typeof s !== 'object') return;
        const email = String(s.email || s.mail || s.upn || s.userPrincipalName || '')
            .trim()
            .toLowerCase();
        const klasse = String(s.klasse || s.class || s.classCode || s.Klasse || '').trim();
        const name = String(s.name || '').trim();
        if (!email && !klasse && !name) return;
        const key = email || name.toLowerCase() + '|' + klasse.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        rows.push({
            email,
            mail: email,
            klasse,
            name,
            class: klasse,
            classCode: klasse
        });
    };
    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const core = window.ms365TenantSettingsLoad();
            const data = (core && core.data) || core || {};
            (Array.isArray(data.students) ? data.students : []).forEach(push);
        }
    } catch {
        /* ignore */
    }
    try {
        const api = appDataApi();
        const c = api && typeof api.getContainer === 'function' ? api.getContainer() : null;
        const by =
            c && c.years && c.years.byLabel && typeof c.years.byLabel === 'object' ? c.years.byLabel : null;
        if (by) {
            Object.keys(by).forEach((lab) => {
                const bucket = by[lab];
                (bucket && bucket.students ? bucket.students : []).forEach(push);
            });
        }
    } catch {
        /* ignore */
    }
    return rows;
}

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
