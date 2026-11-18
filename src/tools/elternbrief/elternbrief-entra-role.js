/**
 * Elternbrief: Admin aus Entra/Schulleitung (Ablauf-Helper).
 */
import { resolveAblaufAdminAccess } from '../../shared/ablauf-admin-entra.js';
import { loadElternbriefSettings } from './elternbrief-settings.js';
import {
    getElternbriefAccount,
    parseAdminMails,
    setElternbriefAdminAccessCache
} from './elternbrief-scope.js';

/**
 * @param {string} [email]
 */
export async function resolveElternbriefAdmin(email) {
    const em = String(email || getElternbriefAccount().email || '')
        .trim()
        .toLowerCase();
    const settings = loadElternbriefSettings();
    const result = await resolveAblaufAdminAccess(em, {
        legacyAdminEmails: parseAdminMails(settings.adminEmails),
        includeIt: true,
        includeVerwaltung: true
    });
    setElternbriefAdminAccessCache({
        email: em,
        isAdmin: !!result.isAdmin,
        sources: result.sources || [],
        demo: !!result.demo
    });
    return result;
}
