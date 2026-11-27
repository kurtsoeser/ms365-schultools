/**
 * Wahlfach-Zuteilung im Lehrer-Portal nur für Direktion, Klassenvorstand und Schul-IT.
 */
import { registerDashboardMenuHook } from '../../shared/dashboard-audience-catalog.js';
import { accountIsPlannerUserInList } from '../freistellung-planer/freistellung-planer-direktion-users.js';
import { loadEffectivePermissionsConfig as loadSaPerms } from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import { probeWahlfachDirektion, readWahlfachDemoRole } from './wahlfach-entra-role.js';
import {
    kvKlassenForEmail,
    loadWzStammdaten,
    readWzAccount,
    shouldShowWahlfachZuteilungMenu
} from './wahlfach-zuteilung-state.js';

const CACHE_KEY = 'ms365-wahlfach-menu-role-v1';
const PERSONA_CACHE_KEY = 'ms365-dashboard-persona-cache-v2';

function readPersona() {
    try {
        const raw = sessionStorage.getItem(PERSONA_CACHE_KEY);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        return parsed && parsed.personas ? parsed.personas : null;
    } catch {
        return null;
    }
}

/**
 * @param {string} email
 * @returns {'direktion'|'lehrer'|''}
 */
function readCachedRole(email) {
    try {
        const raw = sessionStorage.getItem(CACHE_KEY);
        if (!raw) return '';
        const parsed = JSON.parse(raw);
        if (!parsed || parsed.email !== email) return '';
        return parsed.role === 'direktion' ? 'direktion' : parsed.role === 'lehrer' ? 'lehrer' : '';
    } catch {
        return '';
    }
}

/**
 * @param {string} email
 * @param {'direktion'|'lehrer'} role
 */
export function rememberWahlfachMenuRole(email, role) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return;
    try {
        sessionStorage.setItem(
            CACHE_KEY,
            JSON.stringify({ email: em, role: role === 'direktion' ? 'direktion' : 'lehrer' })
        );
    } catch {
        /* ignore */
    }
}

export function wahlfachMenuContext() {
    const email = readWzAccount().email;
    const personas = readPersona();
    let localAdmin = false;
    try {
        const sa = loadSaPerms();
        localAdmin = !!(email && accountIsPlannerUserInList(email, (sa && sa.adminUsers) || []));
    } catch {
        localAdmin = false;
    }
    return {
        isIt: !!(personas && personas.isIt),
        email,
        demoRole: readWahlfachDemoRole(),
        kvKlassen: kvKlassenForEmail(loadWzStammdaten(), email),
        localAdmin,
        cachedRole: readCachedRole(email)
    };
}

function wahlfachMenuAllows(toolId, view) {
    if (toolId !== 'wahlfach-zuteilung' || view !== 'lehrer') return null;
    return shouldShowWahlfachZuteilungMenu(wahlfachMenuContext());
}

registerDashboardMenuHook(wahlfachMenuAllows);

/**
 * Entra-Direktion nachziehen. true, wenn sich die Kachel dadurch ändert.
 * @returns {Promise<boolean>}
 */
export async function refreshWahlfachMenuRole() {
    const before = wahlfachMenuContext();
    if (!before.email || before.isIt || before.demoRole) return false;
    if (before.kvKlassen.length || before.localAdmin || before.cachedRole) return false;
    let hit = { ok: false };
    try {
        hit = await probeWahlfachDirektion();
    } catch {
        hit = { ok: false };
    }
    rememberWahlfachMenuRole(before.email, hit && hit.ok ? 'direktion' : 'lehrer');
    const afterShow = shouldShowWahlfachZuteilungMenu(wahlfachMenuContext());
    const beforeShow = shouldShowWahlfachZuteilungMenu(before);
    return afterShow !== beforeShow;
}
