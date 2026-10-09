/**
 * SharePoint-Berechtigungen Lehrer-Freistellungsliste (Entra-Gruppen aus lfr-permissions).
 */
import {
    applyEntraGroupListPermissions,
    acquireSpoContext,
    roleDefIdForLevel
} from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import { loadLfrPermissions, normalizeLfrPermissions } from './lfr-permissions.js';

const LIST_PROFILE = {
    admin: 'fullControl',
    lehrer: 'contribute',
    direktion: 'design'
};

/**
 * @param {string} siteWebUrl
 * @param {string} listTitle
 * @param {string} serviceAccountEmail
 * @param {{ listId?: string }} [opts]
 * @param {(msg: string) => void} [logFn]
 */
export async function grantLfrFlowServiceAccountOnList(
    siteWebUrl,
    listTitle,
    serviceAccountEmail,
    opts,
    logFn
) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const mail = String(serviceAccountEmail || '').trim().toLowerCase();
    if (!mail) {
        write('! Technik-Konto fehlt – kein Vollzugriff für den Flow.');
        return { skipped: true, reason: 'no-account' };
    }
    const title = String(listTitle || '').trim() || 'Lehrer-Freistellungen';
    const G = window.ms365SpoGraph;
    if (!G || typeof G.spoEnsureUser !== 'function' || typeof G.spoAddRoleAssignment !== 'function') {
        write('! SharePoint-Hilfsfunktionen fehlen.');
        return { skipped: true, reason: 'no-spo-api' };
    }
    write('Technik-Konto: Vollzugriff auf „' + title + '“ (' + mail + ') …');
    const ctx = await acquireSpoContext(siteWebUrl);
    const principal = await G.spoEnsureUser(ctx.url, ctx.spoToken, ctx.digest, mail);
    await G.spoAddRoleAssignment(
        ctx.url,
        ctx.spoToken,
        ctx.digest,
        title,
        principal.id,
        roleDefIdForLevel(LIST_PROFILE.admin)
    );
    write('Technik-Konto zugewiesen.');
    return { ok: true };
}

/**
 * @param {string} siteWebUrl
 * @param {string} listTitle
 * @param {{ listId?: string, flowServiceAccount?: string }} [opts]
 * @param {(msg: string) => void} [logFn]
 */
export async function applyLfrListPermissions(siteWebUrl, listTitle, opts, logFn) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const perms = normalizeLfrPermissions(loadLfrPermissions());
    const groups = [];
    if (perms.groupLehrerId) {
        groups.push({ id: perms.groupLehrerId, name: perms.groupLehrerName, level: LIST_PROFILE.lehrer });
    }
    if (perms.groupDirektionId) {
        groups.push({
            id: perms.groupDirektionId,
            name: perms.groupDirektionName,
            level: LIST_PROFILE.direktion
        });
    }
    if (!groups.length) {
        write('Keine Entra-Gruppen in Schritt 4 – Berechtigungen übersprungen.');
        return { skipped: true };
    }
    await applyEntraGroupListPermissions(siteWebUrl, listTitle, groups, write);
    const flowAccount = opts && opts.flowServiceAccount ? String(opts.flowServiceAccount).trim() : '';
    if (flowAccount) {
        await grantLfrFlowServiceAccountOnList(siteWebUrl, listTitle, flowAccount, opts, write);
    }
    return { ok: true };
}
