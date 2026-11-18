/**
 * Wahlfach-Zuteilung (Admin/SLG): Zugang über Schul-IT / SA-Admin / Schulleitung.
 * Kein Rollen-Switcher – Schüler nutzen die separate Anmelde-Seite.
 * IT-Vorschau: ?demoRole=1
 */
import { readDemoRoleOverrideFromUrl } from '../../shared/app-role-switcher-policy.js';
import { userIsDesignatedSchoolIt } from '../../shared/designated-school-it.js';
import {
    fetchUserMemberGroupIds,
    userIsEntraGlobalAdministrator
} from '../schularbeiten-planer/schularbeiten-planer-entra-role.js';
import { loadEffectivePermissionsConfig as loadSaPerms } from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import { accountIsPlannerUserInList } from '../freistellung-planer/freistellung-planer-direktion-users.js';

/**
 * @returns {Promise<{ ok: boolean, reason: string, demo: boolean }>}
 */
export async function resolveWahlfachAdminAccess() {
    const demo = readDemoRoleOverrideFromUrl(false);
    if (demo) {
        return { ok: true, reason: 'IT-Vorschau (?demoRole=1)', demo: true };
    }

    try {
        if (await userIsEntraGlobalAdministrator()) {
            return { ok: true, reason: 'Globaler Administrator', demo: false };
        }
    } catch {
        /* ignore */
    }
    try {
        if (await userIsDesignatedSchoolIt()) {
            return { ok: true, reason: 'Schul-IT', demo: false };
        }
    } catch {
        /* ignore */
    }

    let sa = { groupAdminId: '', adminUsers: [] };
    try {
        sa = loadSaPerms();
    } catch {
        /* ignore */
    }

    if (sa.groupAdminId) {
        try {
            const member = await fetchUserMemberGroupIds([sa.groupAdminId]);
            if (member.has(String(sa.groupAdminId).toLowerCase())) {
                return { ok: true, reason: 'Schulleitung / Admin-Gruppe', demo: false };
            }
        } catch {
            /* ignore */
        }
    }

    let email = '';
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const a = window.ms365AuthGetAccountInfo();
            email = String((a && (a.username || a.upn)) || '')
                .trim()
                .toLowerCase();
        }
    } catch {
        /* ignore */
    }
    if (email && accountIsPlannerUserInList(email, sa.adminUsers || [])) {
        return { ok: true, reason: 'Admin-Einzelperson', demo: false };
    }

    // Ohne konfigurierte Admin-Gruppe: zugänglich (Setup-Phase), Hinweis bleibt
    if (!sa.groupAdminId && !(sa.adminUsers && sa.adminUsers.length)) {
        return {
            ok: true,
            reason: 'Noch keine Admin-Gruppe konfiguriert (Schularbeiten-Setup)',
            demo: false
        };
    }

    return {
        ok: false,
        reason: 'Kein Zugang zur Wahlfach-Zuteilung. Nur Schulleitung/Admin bzw. Schul-IT.',
        demo: false
    };
}
