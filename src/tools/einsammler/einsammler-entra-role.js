/**
 * Einsammler: Admin aus Entra/Schulleitung (Ablauf-Helper).
 */
import { resolveAblaufAdminAccess } from '../../shared/ablauf-admin-entra.js';
import { loadEinsammlerSettings } from './einsammler-settings.js';
import {
    getEinsammlerAccount,
    parseAdminMails,
    setEinsammlerAdminAccessCache
} from './einsammler-scope.js';
import { announceStaffAccess } from '../../shared/app-role-chrome.js';

/**
 * @param {string} [email]
 */
export async function resolveEinsammlerAdmin(email) {
    const em = String(email || getEinsammlerAccount().email || '')
        .trim()
        .toLowerCase();
    const settings = loadEinsammlerSettings();
    const result = await resolveAblaufAdminAccess(em, {
        legacyAdminEmails: parseAdminMails(settings.adminEmails),
        includeIt: true,
        includeVerwaltung: true
    });
    setEinsammlerAdminAccessCache({
        email: em,
        isAdmin: !!result.isAdmin,
        sources: result.sources || [],
        demo: !!result.demo
    });
    announceStaffAccess('einsammler', {
        email: em,
        isAdmin: !!result.isAdmin,
        sources: result.sources || []
    });
    return result;
}
