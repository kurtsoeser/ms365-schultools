/**
 * Kollegiumsumfragen: Admin (Schulleitung/…) vs. Lehrkraft/KV-Audience.
 */
import {
    resolveAblaufAdminAccess,
    resolveIsAblaufKvMember,
    resolveIsAblaufTeacherMember
} from '../../shared/ablauf-admin-entra.js';
import { resolveTeacherContext } from './kollegium-umfragen-logic.js';

/**
 * @returns {{ email: string, name: string, loggedIn: boolean }}
 */
function getKollegiumAccount() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            if (info) {
                const email = String(info.username || info.upn || '')
                    .trim()
                    .toLowerCase();
                return {
                    email,
                    name: String(info.name || '').trim(),
                    loggedIn: !!email
                };
            }
        }
    } catch {
        /* ignore */
    }
    return { email: '', name: '', loggedIn: false };
}

/** @type {{ email: string, isAdmin: boolean, isTeacher: boolean, isKv: boolean, sources: string[], demo: boolean, resolved: boolean }} */
let cache = {
    email: '',
    isAdmin: false,
    isTeacher: false,
    isKv: false,
    sources: [],
    demo: false,
    resolved: false
};

/**
 * @param {string} [email]
 */
export async function resolveKollegiumRoles(email) {
    const acc = getKollegiumAccount();
    const em = String(email || acc.email || '')
        .trim()
        .toLowerCase();
    const stamm = resolveTeacherContext(em);
    const admin = await resolveAblaufAdminAccess(em, {
        includeIt: true,
        includeVerwaltung: true
    });
    let isTeacher = !!stamm.isTeacher;
    let isKv = !!stamm.isKv;
    try {
        if (!isTeacher && (await resolveIsAblaufTeacherMember())) isTeacher = true;
    } catch {
        /* ignore */
    }
    try {
        if (!isKv && (await resolveIsAblaufKvMember())) {
            isKv = true;
            isTeacher = true;
        }
    } catch {
        /* ignore */
    }
    if (admin.isAdmin) isTeacher = true;
    cache = {
        email: em,
        isAdmin: !!admin.isAdmin,
        isTeacher,
        isKv,
        sources: admin.sources || [],
        demo: !!admin.demo,
        resolved: true
    };
    return {
        isAdmin: cache.isAdmin,
        isTeacher: cache.isTeacher,
        isKv: cache.isKv,
        sources: cache.sources.slice(),
        demo: cache.demo,
        abteilung: stamm.abteilung,
        kvKlassen: stamm.kvKlassen.slice()
    };
}

export function getKollegiumRoleCache() {
    return {
        email: cache.email,
        isAdmin: cache.isAdmin,
        isTeacher: cache.isTeacher,
        isKv: cache.isKv,
        sources: cache.sources.slice(),
        demo: cache.demo,
        resolved: cache.resolved
    };
}

export function isKollegiumAdminCached(email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (cache.resolved && cache.email === em) return cache.isAdmin;
    return false;
}

/**
 * Teacher-Kontext inkl. Entra-Cache für Audience-Checks.
 * @param {string} [email]
 */
export function kollegiumTeacherCtxFromCache(email) {
    const em = String(email || cache.email || '')
        .trim()
        .toLowerCase();
    const stamm = resolveTeacherContext(em);
    if (cache.resolved && cache.email === em) {
        return {
            email: em,
            isKv: cache.isKv || stamm.isKv,
            isTeacher: cache.isTeacher || stamm.isTeacher,
            abteilung: stamm.abteilung,
            kvKlassen: stamm.kvKlassen
        };
    }
    return stamm;
}
