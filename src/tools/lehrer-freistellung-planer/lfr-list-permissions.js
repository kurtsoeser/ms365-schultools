/**
 * SharePoint-Berechtigungen Lehrer-Freistellungsliste (Gruppen + Einzelpersonen).
 */
import {
    applyEntraGroupListPermissions,
    acquireSpoContext,
    roleDefIdForLevel
} from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import { loadLfrPermissions, normalizeLfrPermissions, lfrPermsStepConfigured } from './lfr-permissions.js';

/** Direktion → Design; Lehrkräfte → Contribute (wie bisher). */
const LIST_PROFILE = {
    admin: 'design',
    lehrer: 'contribute',
    schueler: null
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
        roleDefIdForLevel('fullControl')
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
    const title = String(listTitle || '').trim() || 'Lehrer-Freistellungen';
    const flowAccount = opts && opts.flowServiceAccount ? String(opts.flowServiceAccount).trim() : '';

    if (!lfrPermsStepConfigured(perms)) {
        write('Keine Gruppen/Personen in Schritt 4 – nur Technik-Konto falls gesetzt.');
        if (flowAccount) {
            await grantLfrFlowServiceAccountOnList(siteWebUrl, title, flowAccount, opts, write);
        }
        return { skipped: true, reason: 'no-audience' };
    }

    const config = {
        groupAdmin: perms.groupDirektionName,
        groupAdminId: perms.groupDirektionId,
        groupLehrer: perms.groupLehrerName,
        groupLehrerId: perms.groupLehrerId,
        groupSchueler: '',
        groupSchuelerId: '',
        adminUsers: perms.direktionUsers,
        lehrerUsers: perms.lehrerUsers,
        schuelerUsers: []
    };

    write('Berechtigungen auf Liste „' + title + '“ (Gruppen + Einzelpersonen) …');
    const ctx = await acquireSpoContext(siteWebUrl);
    await applyEntraGroupListPermissions(
        ctx.url,
        ctx.spoToken,
        ctx.digest,
        title,
        ctx.graphToken,
        config,
        LIST_PROFILE,
        write
    );

    if (flowAccount) {
        await grantLfrFlowServiceAccountOnList(siteWebUrl, title, flowAccount, opts, write);
    }
    return { ok: true };
}
