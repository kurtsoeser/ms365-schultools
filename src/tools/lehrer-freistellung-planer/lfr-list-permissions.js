/**
 * SharePoint-Berechtigungen Lehrer-Freistellungsliste (Gruppen + Einzelpersonen).
 */
import {
    applyEntraGroupListPermissions,
    acquireSpoContext,
    roleDefIdForLevel
} from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import {
    loadLfrPermissions,
    normalizeLfrPermissions,
    lfrPermsStepConfigured,
    resolveLfrApproverEmail
} from './lfr-permissions.js';
import { resolveFlowServiceAccount } from '../../shared/flow-service-account-list-access.js';
import {
    direktionUsersFromTenantStammdaten,
    normalizePlannerUsers
} from '../freistellung-planer/freistellung-planer-direktion-users.js';

/** Direktion → fullControl (alle Anträge sichtbar); Lehrkräfte → Contribute. */
const LIST_PROFILE = {
    admin: 'fullControl',
    lehrer: 'contribute',
    schueler: null
};

/**
 * SharePoint „Erweiterte Einstellungen“: Lesen/Schreiben nur eigene Elemente.
 * Full Control (Direktion / Flow) sieht weiterhin alles; Contribute nur Author-eigene.
 */
export const LFR_LIST_ITEM_LEVEL = {
    readSecurity: 2,
    writeSecurity: 2
};

/**
 * Direktion für Listenrechte: Setup-Personen + Approver-Mail + Schulleitung (Stammdaten).
 * Sonst hat der Direktor App-Rolle „Direktion“, aber nach Vererbungsbruch keine SPO-Rechte.
 * @param {ReturnType<typeof normalizeLfrPermissions>} [perms]
 */
export function effectiveLfrDirektionUsersForListAccess(perms) {
    const p = normalizeLfrPermissions(perms || loadLfrPermissions());
    /** @type {Array<{ id?: string, displayName?: string, mail?: string }>} */
    const merged = [...(p.direktionUsers || [])];
    const approver = resolveLfrApproverEmail(p);
    if (approver) {
        merged.push({ mail: approver, displayName: approver, id: '' });
    }
    try {
        direktionUsersFromTenantStammdaten().forEach((u) => merged.push(u));
    } catch {
        /* ignore */
    }
    return normalizePlannerUsers(merged);
}

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
    const flowAccount = resolveFlowServiceAccount(
        'lfr',
        opts && opts.flowServiceAccount ? opts.flowServiceAccount : ''
    );

    if (!lfrPermsStepConfigured(perms)) {
        write('Keine Gruppen/Personen in Schritt 4 – Technik-Konto wird trotzdem abgesichert.');
        await grantLfrFlowServiceAccountOnList(siteWebUrl, title, flowAccount, opts, write);
        return { skipped: true, reason: 'no-audience' };
    }

    const adminUsers = effectiveLfrDirektionUsersForListAccess(perms);
    if (adminUsers.length && !(perms.direktionUsers && perms.direktionUsers.length)) {
        write(
            '  · Direktion-Einzelpersonen aus Approver/Stammdaten Schulleitung ergänzt (' +
                adminUsers.length +
                ').'
        );
    }

    const config = {
        groupAdmin: perms.groupDirektionName,
        groupAdminId: perms.groupDirektionId,
        groupLehrer: perms.groupLehrerName,
        groupLehrerId: perms.groupLehrerId,
        groupSchueler: '',
        groupSchuelerId: '',
        adminUsers,
        lehrerUsers: perms.lehrerUsers,
        schuelerUsers: []
    };

    write('Berechtigungen auf Liste „' + title + '“ (Gruppen + Einzelpersonen) …');
    const ctx = await acquireSpoContext(siteWebUrl);
    const listId = opts && opts.listId ? String(opts.listId).trim() : '';
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

    const G = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (G && typeof G.spoSetListItemLevelPermissions === 'function') {
        try {
            await G.spoSetListItemLevelPermissions(
                ctx.url,
                ctx.spoToken,
                ctx.digest,
                title,
                LFR_LIST_ITEM_LEVEL.readSecurity,
                LFR_LIST_ITEM_LEVEL.writeSecurity,
                listId || undefined
            );
            let verified = false;
            if (typeof G.spoGetListItemLevelPermissions === 'function') {
                const cur = await G.spoGetListItemLevelPermissions(
                    ctx.url,
                    ctx.spoToken,
                    ctx.digest,
                    title,
                    listId || undefined
                );
                verified =
                    cur.readSecurity === LFR_LIST_ITEM_LEVEL.readSecurity &&
                    cur.writeSecurity === LFR_LIST_ITEM_LEVEL.writeSecurity;
                if (verified) {
                    write(
                        '  „' +
                            title +
                            '": Elementberechtigungen aktiv (Lehrkräfte lesen/schreiben nur eigene Anträge; Direktion mit Vollzugriff sieht alle).'
                    );
                } else {
                    write(
                        '  ! Elementberechtigungen nicht übernommen (ReadSecurity=' +
                            cur.readSecurity +
                            ', WriteSecurity=' +
                            cur.writeSecurity +
                            '). Manuell: Listen-Einstellungen → Erweitert → „Elemente lesen: nur eigene“ / „Elemente erstellen und bearbeiten: nur eigene“.'
                    );
                }
            } else {
                write(
                    '  „' +
                        title +
                        '": Elementberechtigungen gesetzt (nur eigene Elemente für Mitwirkende).'
                );
            }
        } catch (e) {
            const msg = e && e.message ? e.message : String(e);
            write('  ! Elementberechtigungen fehlgeschlagen: ' + msg);
            write(
                '  → Manuell in SharePoint: Liste → Einstellungen → Erweiterte Einstellungen → nur eigene Elemente.'
            );
        }
    } else {
        write(
            '  ! SPO-API für Elementberechtigungen fehlt – bitte manuell „nur eigene Elemente“ setzen.'
        );
    }

    // Immer nach Rechte-Schreiben: Automate-/Technik-Konto behält Vollzugriff
    await grantLfrFlowServiceAccountOnList(siteWebUrl, title, flowAccount, opts, write);
    return { ok: true };
}
