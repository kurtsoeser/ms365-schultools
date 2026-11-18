/**
 * Ablauf-Apps (Einsammler, Elternbrief, Krankmeldung, Kollegium):
 * Admin = Schulleitung / SA-Admin / Verwaltung / Schul-IT – nicht nur Mail-Listen.
 * Docs: rollen-vertrag.md Welle 3.
 */
import { readDemoRoleOverrideFromUrl } from './app-role-switcher-policy.js';
import { userIsDesignatedSchoolIt } from './designated-school-it.js';
import { loadSchoolAdminGroups } from './school-admin-groups.js';
import { loadSchoolAudienceGroups } from './school-audience-groups.js';
import {
    fetchUserMemberGroupIds,
    userIsEntraGlobalAdministrator
} from '../tools/schularbeiten-planer/schularbeiten-planer-entra-role.js';
import { loadEffectivePermissionsConfig as loadSaPerms } from '../tools/schularbeiten-planer/schularbeiten-planer-permissions.js';
import { loadEffectivePermissionsConfig as loadFrPerms } from '../tools/freistellung-planer/freistellung-planer-permissions.js';
import { accountIsPlannerUserInList } from '../tools/freistellung-planer/freistellung-planer-direktion-users.js';

/**
 * @returns {string[]}
 */
export function ablaufAdminGroupIds() {
    /** @type {string[]} */
    const ids = [];
    try {
        const aud = loadSchoolAdminGroups();
        if (aud.schulleitungGroupId) ids.push(aud.schulleitungGroupId);
        if (aud.verwaltungGroupId) ids.push(aud.verwaltungGroupId);
    } catch {
        /* ignore */
    }
    try {
        const sa = loadSaPerms();
        if (sa.groupAdminId) ids.push(sa.groupAdminId);
    } catch {
        /* ignore */
    }
    try {
        const fr = loadFrPerms();
        if (fr.groupDirektionId) ids.push(fr.groupDirektionId);
    } catch {
        /* ignore */
    }
    return [...new Set(ids.map((x) => String(x).trim()).filter(Boolean))];
}

/**
 * @returns {string[]}
 */
export function ablaufLehrerGroupIds() {
    /** @type {string[]} */
    const ids = [];
    try {
        const stamm = loadSchoolAudienceGroups();
        if (stamm.groupLehrerId) ids.push(stamm.groupLehrerId);
    } catch {
        /* ignore */
    }
    try {
        const sa = loadSaPerms();
        if (sa.groupLehrerId) ids.push(sa.groupLehrerId);
    } catch {
        /* ignore */
    }
    return [...new Set(ids.map((x) => String(x).trim()).filter(Boolean))];
}

/**
 * @returns {string[]}
 */
export function ablaufKvGroupIds() {
    /** @type {string[]} */
    const ids = [];
    try {
        const fr = loadFrPerms();
        if (fr.groupKvId) ids.push(fr.groupKvId);
    } catch {
        /* ignore */
    }
    try {
        const g =
            typeof globalThis !== 'undefined'
                ? globalThis
                : typeof window !== 'undefined'
                  ? window
                  : {};
        const api = g.ms365AppDataV2;
        if (api && typeof api.getSetup === 'function') {
            const matched = (api.getSetup() || {}).matched || {};
            if (matched.kvGroupId) ids.push(String(matched.kvGroupId).trim());
        }
    } catch {
        /* ignore */
    }
    return [...new Set(ids.map((x) => String(x).trim()).filter(Boolean))];
}

/**
 * Stammdaten-Administration (Schulleitung / Verwaltung / IT) – Legacy-Fallback.
 * @param {string} email
 * @param {{ includeIt?: boolean, includeVerwaltung?: boolean }} [opts]
 */
export function isAblaufAdminFromStammdaten(email, opts) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return false;
    const includeIt = !opts || opts.includeIt !== false;
    const includeVerwaltung = !opts || opts.includeVerwaltung !== false;
    try {
        if (typeof window === 'undefined' || typeof window.ms365TenantSettingsLoad !== 'function') {
            return false;
        }
        const core = window.ms365TenantSettingsLoad();
        const data = (core && core.data) || core || {};
        for (const a of Array.isArray(data.admin) ? data.admin : []) {
            if (String((a && a.email) || '').trim().toLowerCase() === em) return true;
        }
        for (const a of Array.isArray(data.administration) ? data.administration : []) {
            const mail = String((a && a.email) || '')
                .trim()
                .toLowerCase();
            if (mail !== em) continue;
            const role = String((a && (a.role || a.group || a.code || a.tier)) || '').toLowerCase();
            if (!role) return true;
            if (role.includes('schulleitung') || role.includes('direktion') || role.includes('admin')) {
                return true;
            }
            if (includeVerwaltung && (role.includes('verwaltung') || role.includes('sekretariat'))) {
                return true;
            }
            if (includeIt && (role.includes('it') || role.includes('technik'))) return true;
        }
    } catch {
        /* ignore */
    }
    return false;
}

/**
 * @param {string} [email]
 * @param {{
 *   legacyAdminEmails?: string[],
 *   includeVerwaltung?: boolean,
 *   includeIt?: boolean
 * }} [opts]
 * @returns {Promise<{ isAdmin: boolean, sources: string[], demo: boolean, entraConfigured: boolean }>}
 */
export async function resolveAblaufAdminAccess(email, opts) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    const demo = readDemoRoleOverrideFromUrl(false);
    if (demo) {
        return {
            isAdmin: true,
            sources: ['demo'],
            demo: true,
            entraConfigured: true
        };
    }

    /** @type {string[]} */
    const sources = [];
    const groupIds = ablaufAdminGroupIds();
    const entraConfigured = groupIds.length > 0;

    try {
        if (await userIsEntraGlobalAdministrator()) {
            sources.push('global-admin');
        }
    } catch {
        /* ignore */
    }
    try {
        if (await userIsDesignatedSchoolIt()) {
            sources.push('school-it');
        }
    } catch {
        /* ignore */
    }

    if (groupIds.length) {
        try {
            const member = await fetchUserMemberGroupIds(groupIds);
            const want = new Set(groupIds.map((id) => String(id).toLowerCase()));
            for (const id of member) {
                if (want.has(String(id).toLowerCase())) {
                    sources.push('entra');
                    break;
                }
            }
        } catch {
            /* ignore */
        }
    }

    try {
        const sa = loadSaPerms();
        if (em && accountIsPlannerUserInList(em, sa.adminUsers || [])) {
            sources.push('admin-user');
        }
    } catch {
        /* ignore */
    }

    if (
        isAblaufAdminFromStammdaten(em, {
            includeIt: !opts || opts.includeIt !== false,
            includeVerwaltung: !opts || opts.includeVerwaltung !== false
        })
    ) {
        sources.push('stammdaten');
    }

    const legacy = (opts && opts.legacyAdminEmails) || [];
    if (em && legacy.map((x) => String(x).toLowerCase()).includes(em)) {
        sources.push('legacy-admin-emails');
    }

    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365OperatorAccess &&
            typeof window.ms365OperatorAccess.isOperatorCached === 'function' &&
            window.ms365OperatorAccess.isOperatorCached()
        ) {
            sources.push('operator');
        }
    } catch {
        /* ignore */
    }

    return {
        isAdmin: sources.length > 0,
        sources,
        demo: false,
        entraConfigured
    };
}

/**
 * Mitgliedschaft Lehrer-Sammelgruppe (für Kollegium-Zielgruppe).
 * @param {string[]} [extraGroupIds]
 * @returns {Promise<boolean>}
 */
export async function resolveIsAblaufTeacherMember(extraGroupIds) {
    const ids = [...ablaufLehrerGroupIds(), ...(extraGroupIds || [])].filter(Boolean);
    if (!ids.length) return false;
    try {
        const member = await fetchUserMemberGroupIds(ids);
        return ids.some((id) => member.has(String(id).toLowerCase()));
    } catch {
        return false;
    }
}

/**
 * Mitgliedschaft KV-Gruppe.
 * @returns {Promise<boolean>}
 */
export async function resolveIsAblaufKvMember() {
    const ids = ablaufKvGroupIds();
    if (!ids.length) return false;
    try {
        const member = await fetchUserMemberGroupIds(ids);
        return ids.some((id) => member.has(String(id).toLowerCase()));
    } catch {
        return false;
    }
}
