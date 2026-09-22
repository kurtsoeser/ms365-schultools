'use strict';

const { getConfig, isTenantAllowed } = require('./config');
const { validateCallerToken } = require('./validate-token');
const { graphJson } = require('./graph-client');
const { getAppOnlyToken } = require('./msal-app-only');

/**
 * @param {Array<{ roleTemplateId?: string }>} directoryRoles
 * @param {string[]} allowedIds
 */
function hasOperatorRole(directoryRoles, allowedIds) {
    const allow = new Set((allowedIds || []).map((s) => String(s).toLowerCase()));
    return (directoryRoles || []).some((role) =>
        allow.has(String(role && role.roleTemplateId ? role.roleTemplateId : '').toLowerCase())
    );
}

/**
 * @param {string} token
 * @returns {Promise<{ tid: string, oid: string, upn: string }>}
 */
async function requireKursteamCaller(token) {
    const cfg = getConfig();
    return validateCallerToken(token, cfg.tokenAudiences);
}

/**
 * @param {{ tid: string, oid: string }} caller
 */
async function assertCallerMayCreateTeams(caller) {
    const cfg = getConfig();
    if (!isTenantAllowed(caller.tid)) {
        const err = new Error('Dieser Mandant ist für das Kursteams-Backend nicht freigeschaltet.');
        err.status = 403;
        throw err;
    }

    let roles = [];
    try {
        const graphToken = await getAppOnlyToken(caller.tid);
        roles = await listDirectoryRoles(graphToken, caller.oid);
    } catch (e) {
        const err = new Error('Berechtigung konnte im Mandanten nicht geprüft werden.');
        err.status = 403;
        err.cause = e;
        throw err;
    }

    if (!hasOperatorRole(roles, cfg.operatorRoleTemplateIds)) {
        const err = new Error('Dieses Konto darf keine Kursteams anlegen.');
        err.status = 403;
        throw err;
    }
}

/**
 * @param {string} graphToken
 * @param {string} oid
 */
async function listDirectoryRoles(graphToken, oid) {
    let url =
        '/users/' +
        encodeURIComponent(oid) +
        '/transitiveMemberOf/microsoft.graph.directoryRole?$select=roleTemplateId,displayName';
    const out = [];
    for (let i = 0; i < 5 && url; i++) {
        const page = await graphJson('GET', url, graphToken);
        if (page && Array.isArray(page.value)) out.push(...page.value);
        url = page && page['@odata.nextLink'] ? String(page['@odata.nextLink']) : '';
    }
    return out;
}

/**
 * @param {{ createdByOid?: string, tenantId?: string } | null} job
 * @param {{ oid?: string, tid?: string } | null} caller
 */
function jobVisibleToCaller(job, caller) {
    if (!job || !caller) return false;
    const oid = String(caller.oid || '').toLowerCase();
    const tid = String(caller.tid || '').toLowerCase();
    if (!oid || !tid) return false;
    if (String(job.createdByOid || '').toLowerCase() !== oid) return false;
    if (String(job.tenantId || '').toLowerCase() !== tid) return false;
    return true;
}

module.exports = {
    requireKursteamCaller,
    assertCallerMayCreateTeams,
    hasOperatorRole,
    jobVisibleToCaller
};
