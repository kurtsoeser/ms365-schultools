/**
 * SharePoint-Berechtigungen Freistellungsliste (Entra-Gruppen aus freistellung-planer-permissions).
 */
import {
    applyEntraGroupListPermissions,
    acquireSpoContext,
    roleDefIdForLevel
} from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import { loadPermissionsConfig, normalizePermissionsConfig } from './freistellung-planer-permissions.js';

/** @type {{ admin: 'fullControl', lehrer: 'contribute', schueler: 'contribute' }} */
const FREISTELLUNG_LIST_PROFILE = {
    admin: 'fullControl',
    lehrer: 'contribute',
    schueler: 'contribute'
};

/**
 * @param {ReturnType<typeof normalizePermissionsConfig>} fr
 */
export function mapFreistellungConfigForSpo(fr) {
    const c = normalizePermissionsConfig(fr);
    return {
        groupAdmin: c.groupDirektion,
        groupAdminId: c.groupDirektionId,
        groupLehrer: c.groupKv,
        groupLehrerId: c.groupKvId,
        groupSchueler: c.groupSchueler,
        groupSchuelerId: c.groupSchuelerId,
        skipPerms: !!c.skipPerms
    };
}

/**
 * @param {string} siteWebUrl
 * @param {string} listTitle
 * @param {object} [configPatch]
 * @param {(msg: string) => void} [logFn]
 */
export async function applyFreistellungListPermissions(siteWebUrl, listTitle, configPatch, logFn) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const fr = normalizePermissionsConfig({ ...loadPermissionsConfig(), ...(configPatch || {}) });
    if (fr.skipPerms) {
        write('Berechtigungen übersprungen (Checkbox aktiv).');
        return { skipped: true };
    }
    if (!fr.groupDirektionId && !fr.groupKvId && !fr.groupSchuelerId) {
        write('! Berechtigungen: keine Entra-Gruppen gewählt – bitte im Setup eintragen.');
        return { skipped: true, reason: 'no-groups' };
    }

    const title = String(listTitle || '').trim() || 'Freistellungen';
    const mapped = mapFreistellungConfigForSpo(fr);
    write('Berechtigungen auf Liste „' + title + '“ …');
    const ctx = await acquireSpoContext(siteWebUrl);
    await applyEntraGroupListPermissions(
        ctx.url,
        ctx.spoToken,
        ctx.digest,
        title,
        ctx.graphToken,
        mapped,
        FREISTELLUNG_LIST_PROFILE,
        write
    );

    const G = window.ms365SpoGraph;
    if (G && !fr.skipPerms) {
        async function grantUsers(users, level, roleLabel) {
            for (const u of users || []) {
                const mail = String(u.mail || '').trim();
                if (!mail) continue;
                try {
                    const principal = await G.spoEnsureUser(ctx.url, ctx.spoToken, ctx.digest, mail);
                    await G.spoAddRoleAssignment(
                        ctx.url,
                        ctx.spoToken,
                        ctx.digest,
                        title,
                        principal.id,
                        roleDefIdForLevel(level)
                    );
                    write(
                        '  + „' +
                            title +
                            '": ' +
                            roleLabel +
                            ' → ' +
                            level +
                            ' (' +
                            (principal.title || u.displayName || mail) +
                            ')'
                    );
                } catch (e) {
                    write('  ! Einzelperson ' + mail + ': ' + (e && e.message ? e.message : e));
                }
            }
        }
        await grantUsers(fr.direktionUsers, FREISTELLUNG_LIST_PROFILE.admin, 'Direktion');
        await grantUsers(fr.kvUsers, FREISTELLUNG_LIST_PROFILE.lehrer, 'KV');
        await grantUsers(fr.schuelerUsers, FREISTELLUNG_LIST_PROFILE.schueler, 'Schüler');
    }

    write('Berechtigungen fertig.');
    return { skipped: false };
}
