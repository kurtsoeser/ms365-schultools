/**
 * Schul-spezifische Overrides für Dashboard-Sichtbarkeit (Stammdaten → Werkzeuge).
 */
import { DASHBOARD_TOOL_RULES, normalizeAudiences } from './dashboard-audience-catalog.js';

export const DASHBOARD_TOOL_ACCESS_STORAGE_KEY = 'ms365-dashboard-tool-access-v1';

/**
 * @typedef {{ audience: import('./dashboard-audience-catalog.js').DashboardAudience[], planner?: import('./dashboard-audience-catalog.js').DashboardToolRule['planner']|null }} StoredToolRule
 * @typedef {{ version: number, tools: Record<string, StoredToolRule> }} DashboardToolAccessConfig
 */

function safeParse(raw) {
    try {
        return JSON.parse(String(raw));
    } catch {
        return null;
    }
}

/**
 * @returns {DashboardToolAccessConfig}
 */
export function loadDashboardToolAccessConfig() {
    try {
        const raw = localStorage.getItem(DASHBOARD_TOOL_ACCESS_STORAGE_KEY);
        const data = safeParse(raw);
        if (data && typeof data === 'object' && data.tools && typeof data.tools === 'object') {
            return { version: 1, tools: { ...data.tools } };
        }
    } catch {
        /* ignore */
    }
    return { version: 1, tools: {} };
}

/**
 * @param {DashboardToolAccessConfig} config
 */
export function saveDashboardToolAccessConfig(config) {
    const payload = {
        version: 1,
        tools: config && config.tools && typeof config.tools === 'object' ? config.tools : {}
    };
    try {
        localStorage.setItem(DASHBOARD_TOOL_ACCESS_STORAGE_KEY, JSON.stringify(payload));
    } catch {
        /* ignore */
    }
    try {
        window.dispatchEvent(new CustomEvent('ms365-dashboard-tool-access-changed'));
    } catch {
        /* ignore */
    }
    return payload;
}

/**
 * @param {string} toolId
 * @param {DashboardToolAccessConfig} [config]
 * @returns {import('./dashboard-audience-catalog.js').DashboardToolRule}
 */
export function getEffectiveToolRule(toolId, config) {
    const base = DASHBOARD_TOOL_RULES[toolId] || { audience: 'it' };
    const cfg = config || loadDashboardToolAccessConfig();
    const ov = cfg.tools && cfg.tools[toolId];
    if (!ov) return base;
    const audience = Array.isArray(ov.audience) && ov.audience.length ? ov.audience : normalizeAudiences(base);
    const planner = ov.planner === null ? undefined : ov.planner !== undefined ? ov.planner : base.planner;
    return { audience, planner };
}

/**
 * @param {string} toolId
 * @param {DashboardToolAccessConfig} [config]
 * @returns {{ lehrer: boolean, schueler: boolean, planner: boolean }}
 */
export function getToolAccessFlags(toolId, config) {
    const rule = getEffectiveToolRule(toolId, config);
    const aud = normalizeAudiences(rule);
    return {
        lehrer: aud.includes('lehrer') || aud.includes('all'),
        schueler: aud.includes('schueler') || aud.includes('all'),
        planner: !!(rule.planner && rule.planner.roles && rule.planner.roles.length)
    };
}

/**
 * @param {string} toolId
 * @param {{ lehrer?: boolean, schueler?: boolean }} flags
 * @param {DashboardToolAccessConfig} [baseConfig]
 */
export function setToolAccessFlags(toolId, flags, baseConfig) {
    const cfg = baseConfig || loadDashboardToolAccessConfig();
    const defaults = getToolAccessFlags(toolId, { version: 1, tools: {} });
    const lehrer = flags.lehrer !== undefined ? !!flags.lehrer : defaults.lehrer;
    const schueler = flags.schueler !== undefined ? !!flags.schueler : defaults.schueler;

    const base = DASHBOARD_TOOL_RULES[toolId] || { audience: 'it' };
    const baseAud = normalizeAudiences(base);
    const baseLehrer = baseAud.includes('lehrer') || baseAud.includes('all');
    const baseSchueler = baseAud.includes('schueler') || baseAud.includes('all');

    const nextTools = { ...(cfg.tools || {}) };
    if (lehrer === baseLehrer && schueler === baseSchueler) {
        delete nextTools[toolId];
    } else {
        /** @type {import('./dashboard-audience-catalog.js').DashboardAudience[]} */
        const audience = [];
        if (lehrer) audience.push('lehrer');
        if (schueler) audience.push('schueler');
        if (!audience.length) audience.push('it');
        const entry = { audience };
        if (base.planner) entry.planner = base.planner;
        nextTools[toolId] = entry;
    }
    return saveDashboardToolAccessConfig({ version: 1, tools: nextTools });
}

export function resetDashboardToolAccessToDefaults() {
    return saveDashboardToolAccessConfig({ version: 1, tools: {} });
}
