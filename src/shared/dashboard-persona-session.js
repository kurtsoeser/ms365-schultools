/**
 * Zentrale Dashboard-Persona: einmal nach Login auflösen, app-weit nutzen.
 */
import {
    landingShotDashboardPersonas,
    resolveDashboardPersonas
} from './dashboard-audience-resolve.js';
import { clearGlobalAdministratorCache } from '../tools/schularbeiten-planer/schularbeiten-planer-entra-role.js';
import { ensureSignedInMs365DemoPortalAccountInLocalStammdaten } from './demo-portal-teacher.js';
import { ensureLiveSchoolStammdatenForBrowser } from './ensure-live-school-stammdaten.js';
import { loadEffectiveTeachersForPersona } from './live-school-stammdaten-cache.js';

const CACHE_KEY = 'ms365-dashboard-persona-cache-v2';
const CACHE_TTL_MS = 20 * 60 * 1000;
/** Kurz cachen, solange Lizenz/Operator/GA noch nachziehen können. */
const CACHE_TTL_INCONCLUSIVE_MS = 15 * 1000;
/**
 * Ohne Konto ist accountCacheKey() leer. Diesen Stand trotzdem cachen,
 * sonst startet jedes persona-ready (Portal-Render) eine neue Auflösung
 * und die Startseite blockiert den Hauptthread.
 */
const ANON_CACHE_KEY = 'anonymous';

/** @type {Promise<import('./dashboard-audience-resolve.js').DashboardPersonas>|null} */
let inFlight = null;
let resolveGen = 0;
/** Screenshot-URLs feuern sonst bei jedem Listener erneut persona-ready (Sync-Rekursion). */
let landingShotNotified = false;

function accountCacheKey() {
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            const oid = String((info && info.oid) || '').trim().toLowerCase();
            const upn = String((info && (info.upn || info.username)) || '')
                .trim()
                .toLowerCase();
            return oid || upn || '';
        }
    } catch {
        /* ignore */
    }
    return '';
}

/** Kontoschlüssel oder anonymer Platzhalter – nie leer, sonst kein Cache-Treffer. */
function personaCacheKey() {
    return accountCacheKey() || ANON_CACHE_KEY;
}

function readCache(key) {
    try {
        const raw = sessionStorage.getItem(CACHE_KEY);
        if (!raw) return null;
        const data = JSON.parse(raw);
        if (!data || data.key !== key || !data.at || !data.personas) return null;
        const ttl = Number(data.ttl) > 0 ? Number(data.ttl) : CACHE_TTL_MS;
        if (Date.now() - Number(data.at) > ttl) return null;
        return data.personas;
    } catch {
        return null;
    }
}

function writeCache(key, personas, ttlMs) {
    try {
        sessionStorage.setItem(
            CACHE_KEY,
            JSON.stringify({
                key,
                at: Date.now(),
                ttl: Number(ttlMs) > 0 ? Number(ttlMs) : CACHE_TTL_MS,
                personas
            })
        );
    } catch {
        /* ignore */
    }
}

/**
 * true wenn Nicht-IT noch „vorläufig“ sein kann (License.Access / Operator / GA-Token fehlt).
 * @param {import('./dashboard-audience-resolve.js').DashboardPersonas|null|undefined} personas
 */
export function isDashboardItResolutionInconclusive(personas) {
    if (!personas || !personas.loggedIn || personas.isIt) return false;

    try {
        const core = window.ms365LicenseGateCore;
        if (core && typeof core.isLicenseApiConfigured === 'function' && core.isLicenseApiConfigured()) {
            const cache = typeof core.readCache === 'function' ? core.readCache() : null;
            if (!cache) return true;
        }
    } catch {
        /* ignore */
    }

    try {
        const oa = window.ms365OperatorAccess;
        if (
            oa &&
            typeof oa.shouldShowAdminMenuLink === 'function' &&
            oa.shouldShowAdminMenuLink() &&
            typeof oa.isCurrentUserOperator === 'function' &&
            !oa.isCurrentUserOperator()
        ) {
            return true;
        }
    } catch {
        /* ignore */
    }

    return false;
}

export function clearDashboardPersonaCache() {
    inFlight = null;
    resolveGen += 1;
    landingShotNotified = false;
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
    const landingShot = landingShotDashboardPersonas();
    if (landingShot) {
        try {
            window.__ms365DashboardPersonasLast = landingShot;
        } catch {
            /* ignore */
        }
        if (!landingShotNotified) {
            landingShotNotified = true;
            try {
                window.dispatchEvent(
                    new CustomEvent('ms365-dashboard-persona-ready', { detail: { personas: landingShot } })
                );
            } catch {
                /* ignore */
            }
        }
        return landingShot;
    }

    const force = !!(opts && opts.force);
    const key = personaCacheKey();

    if (!force) {
        const cached = readCache(key);
        if (cached) return cached;
    }

    if (inFlight && !force) return inFlight;

    const myGen = ++resolveGen;

    const promise = (async () => {
        // Demo-Lehrer/-Schüler (@ms365.schule) vor Persona-Match in Stammdaten absichern
        try {
            ensureSignedInMs365DemoPortalAccountInLocalStammdaten();
        } catch {
            /* ignore */
        }

        // Lehrer-Browser ohne IT-Lokalstand: Lehrerinnen von der Tool-Site nachziehen
        try {
            if (!loadEffectiveTeachersForPersona().length) {
                await ensureLiveSchoolStammdatenForBrowser({ force: false });
            }
        } catch {
            /* ignore */
        }

        // Operator/License.Access oft erst nach License-Gate – kurz mitziehen, bevor IT verworfen wird
        try {
            const oa = window.ms365OperatorAccess;
            if (oa && typeof oa.refreshOperatorStatus === 'function') {
                await oa.refreshOperatorStatus({ force: false });
            }
        } catch {
            /* ignore */
        }

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
        // Nach Live-Hydrate nochmals Persona, falls vorher „Noch nicht zugeordnet“
        if (personas && personas.loggedIn && personas.unknownUser) {
            try {
                const ensured = await ensureLiveSchoolStammdatenForBrowser({ force: true });
                if (ensured && ensured.teacherCount > 0) {
                    personas = await resolveDashboardPersonas({ demoMode: false });
                }
            } catch {
                /* ignore */
            }
        }

        // Neuere Auflösung läuft – kein veralteter Cache / Event
        if (myGen !== resolveGen) {
            return personas;
        }

        if (key) {
            const ttl = isDashboardItResolutionInconclusive(personas)
                ? CACHE_TTL_INCONCLUSIVE_MS
                : CACHE_TTL_MS;
            writeCache(key, personas, ttl);
        }
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

    inFlight = promise;

    try {
        return await promise;
    } finally {
        if (inFlight === promise) inFlight = null;
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
    // Bootstrap / IT-Sicherung / Stammdaten-Pull: Persona neu auflösen (kein „Hängenbleiben“)
    window.addEventListener('ms365-tenant-settings-changed', rerun);
    window.addEventListener('ms365-school-data-written', rerun);
    // License.Access / isDesignatedSchoolIt / Operator oft erst nach dem ersten Persona-Lauf
    window.addEventListener('ms365-license-changed', rerun);
    window.addEventListener('storage', (ev) => {
        if (ev && ev.key === 'ms365-schooltool-data-v2') rerun();
    });
    window.addEventListener('ms365-auth-widget-ready', () => {
        refreshDashboardPersonaSession({ force: false });
    });
}

bootSessionListeners();
