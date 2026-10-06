/**
 * Zentrale Dashboard-Persona: einmal nach Login auflösen, app-weit nutzen.
 */
import { resolveDashboardPersonas } from './dashboard-audience-resolve.js';
import { clearGlobalAdministratorCache } from '../tools/schularbeiten-planer/schularbeiten-planer-entra-role.js';

const CACHE_KEY = 'ms365-dashboard-persona-cache-v2';
const CACHE_TTL_MS = 20 * 60 * 1000;

/** @type {Promise<import('./dashboard-audience-resolve.js').DashboardPersonas>|null} */
let inFlight = null;

function accountCacheKey() {
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            const oid = String((info && info.oid) || '').trim().toLowerCase();
            const upn = String((info && (info.upn || info.username)) || '').trim().toLowerCase();
            return oid || upn || '';
        }
    } catch {
        /* ignore */
    }
    return '';
}

function readCache(key) {
    try {
        const raw = sessionStorage.getItem(CACHE_KEY);
        if (!raw) return null;
        const data = JSON.parse(raw);
        if (!data || data.key !== key || !data.at || !data.personas) return null;
        if (Date.now() - Number(data.at) > CACHE_TTL_MS) return null;
        return data.personas;
    } catch {
        return null;
    }
}

function writeCache(key, personas) {
    try {
        sessionStorage.setItem(CACHE_KEY, JSON.stringify({ key, at: Date.now(), personas }));
    } catch {
        /* ignore */
    }
}

export function clearDashboardPersonaCache() {
    inFlight = null;
    clearGlobalAdministratorCache();
    try {
        sessionStorage.removeItem(CACHE_KEY);
    } catch {
        /* ignore */
    }
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export async function refreshDashboardPersonaSession(opts) {
    const force = !!(opts && opts.force);
    const key = accountCacheKey();

    if (!force && key) {
        const cached = readCache(key);
        if (cached) return cached;
    }

    if (inFlight && !force) return inFlight;

    inFlight = (async () => {
        let personas = await resolveDashboardPersonas({ demoMode: false });
        if (
            personas &&
            personas.loggedIn &&
            !personas.isIt &&
            !personas.globalAdmin &&
            !personas.designatedSchoolIt
        ) {
            clearGlobalAdministratorCache();
            personas = await resolveDashboardPersonas({ demoMode: false });
        }
        if (key) writeCache(key, personas);
        try {
            window.__ms365DashboardPersonasLast = personas;
        } catch {
            /* ignore */
        }
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-dashboard-persona-ready', { detail: { personas } })
            );
        } catch {
            /* ignore */
        }
        return personas;
    })();

    try {
        return await inFlight;
    } finally {
        inFlight = null;
    }
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export async function getDashboardPersonas(opts) {
    return refreshDashboardPersonaSession(opts);
}

function bootSessionListeners() {
    if (typeof window === 'undefined' || window.__ms365DashboardPersonaBoot) return;
    window.__ms365DashboardPersonaBoot = true;

    const rerun = () => {
        clearDashboardPersonaCache();
        refreshDashboardPersonaSession({ force: true });
    };

    window.addEventListener('ms365-auth-state-changed', rerun);
    window.addEventListener('ms365-dashboard-audience-groups-changed', rerun);
    window.addEventListener('ms365-dashboard-tool-access-changed', rerun);
    window.addEventListener('storage', (ev) => {
        if (ev && ev.key === 'ms365-schooltool-data-v2') rerun();
    });
    window.addEventListener('ms365-auth-widget-ready', () => {
        refreshDashboardPersonaSession({ force: false });
    });
}

bootSessionListeners();
