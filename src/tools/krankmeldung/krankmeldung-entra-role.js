/**
 * Krankmeldung Inbox: Sekretariat/Admin vs. KV (Stammdaten).
 */
import {
    resolveAblaufAdminAccess,
    resolveIsAblaufKvMember
} from '../../shared/ablauf-admin-entra.js';
import { getKrankmeldungAccount, resolveKvKlassenForEmail } from './krankmeldung-klassen.js';

/** @type {{ email: string, isAdmin: boolean, isKv: boolean, sources: string[], demo: boolean, resolved: boolean }} */
let cache = {
    email: '',
    isAdmin: false,
    isKv: false,
    sources: [],
    demo: false,
    resolved: false
};

/**
 * @param {string} [email]
 */
export async function resolveKrankmeldungRoles(email) {
    const acc = getKrankmeldungAccount();
    const em = String(email || acc.email || '')
        .trim()
        .toLowerCase();
    const admin = await resolveAblaufAdminAccess(em, {
        includeIt: true,
        includeVerwaltung: true
    });
    const kvKlassen = resolveKvKlassenForEmail(em);
    let isKv = !!(kvKlassen && kvKlassen.length);
    if (!isKv) {
        try {
            isKv = await resolveIsAblaufKvMember();
        } catch {
            /* ignore */
        }
    }
    cache = {
        email: em,
        isAdmin: !!admin.isAdmin,
        isKv,
        sources: admin.sources || [],
        demo: !!admin.demo,
        resolved: true
    };
    return {
        isAdmin: cache.isAdmin,
        isKv: cache.isKv,
        sources: cache.sources.slice(),
        demo: cache.demo,
        kvKlassen: kvKlassen || []
    };
}

export function getKrankmeldungRoleCache() {
    return {
        email: cache.email,
        isAdmin: cache.isAdmin,
        isKv: cache.isKv,
        sources: cache.sources.slice(),
        demo: cache.demo,
        resolved: cache.resolved
    };
}

export function isKrankmeldungAdminCached(email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (cache.resolved && cache.email === em) return cache.isAdmin;
    return false;
}
