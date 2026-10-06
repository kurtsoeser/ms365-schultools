/**
 * Schul-IT laut Lizenz-API: eingetragene IT-Kontakt(e) + Global Admin (in resolve).
 */

function accountUpn() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetUserPrincipalName === 'function') {
            return String(window.ms365AuthGetUserPrincipalName() || '').trim().toLowerCase();
        }
    } catch {
        /* ignore */
    }
    return '';
}

function readLicenseCache() {
    try {
        const core = typeof window !== 'undefined' ? window.ms365LicenseGateCore : null;
        if (!core || typeof core.readCache !== 'function') return null;
        return core.readCache();
    } catch {
        return null;
    }
}

function cacheMatchesCurrentUser(cache) {
    try {
        const core = window.ms365LicenseGateCore;
        if (!core || typeof core.accountKeyFromAuth !== 'function' || typeof core.cacheMatchesAccount !== 'function') {
            return !!cache;
        }
        const key = core.accountKeyFromAuth();
        return core.cacheMatchesAccount(cache, key);
    } catch {
        return false;
    }
}

/**
 * true wenn /license/me den angemeldeten UPN als IT-Kontakt der Schule gemeldet hat.
 */
export function userIsDesignatedSchoolItFromCache() {
    const cache = readLicenseCache();
    if (!cache || !cache.allowed || !cacheMatchesCurrentUser(cache)) return false;
    if (cache.isDesignatedSchoolIt === true) return true;
    const upn = accountUpn();
    if (!upn) return false;
    const lic = cache.license && typeof cache.license === 'object' ? cache.license : {};
    const emails = Array.isArray(lic.contactEmails)
        ? lic.contactEmails
        : lic.contactEmail
          ? [lic.contactEmail]
          : [];
    return emails.some((e) => String(e || '').trim().toLowerCase() === upn);
}

/**
 * @returns {Promise<boolean>}
 */
export async function userIsDesignatedSchoolIt() {
    if (userIsDesignatedSchoolItFromCache()) return true;
    const core = typeof window !== 'undefined' ? window.ms365LicenseGateCore : null;
    if (!core || !core.isLicenseApiConfigured || !core.isLicenseApiConfigured()) return false;
    const api = window.ms365LicenseApi;
    if (!api || typeof api.acquireLicenseToken !== 'function' || typeof api.fetchLicenseMe !== 'function') {
        return false;
    }
    try {
        const token = await api.acquireLicenseToken({ silentOnly: true });
        const data = await api.fetchLicenseMe(token);
        if (data && typeof core.writeCache === 'function') {
            const key = core.accountKeyFromAuth ? core.accountKeyFromAuth() : '';
            if (key) core.writeCache(data, key);
        }
        if (data && data.isDesignatedSchoolIt === true) return true;
        const upn = accountUpn();
        const lic = data && data.license ? data.license : {};
        const emails = Array.isArray(lic.contactEmails)
            ? lic.contactEmails
            : lic.contactEmail
              ? [lic.contactEmail]
              : [];
        return emails.some((e) => String(e || '').trim().toLowerCase() === upn);
    } catch {
        return false;
    }
}
