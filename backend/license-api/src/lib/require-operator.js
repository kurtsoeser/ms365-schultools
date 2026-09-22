'use strict';

const { getConfig } = require('./config');
const { validateCallerToken } = require('./validate-token');
const { lookupLicenseFields } = require('./sharepoint-license');
const { evaluateLicense } = require('./evaluate-license');

/**
 * @param {{ upn?: string, oid?: string }} caller
 * @param {{ operatorUpns: string[], operatorOids: string[] }} cfg
 */
function isOperatorCaller(caller, cfg) {
    const oid = String(caller && caller.oid ? caller.oid : '')
        .trim()
        .toLowerCase();
    const upn = String(caller && caller.upn ? caller.upn : '')
        .trim()
        .toLowerCase();
    if (oid && cfg.operatorOids.includes(oid)) return true;
    if (upn && cfg.operatorUpns.includes(upn)) return true;
    return false;
}

/**
 * @param {string} token
 * @returns {Promise<{ tid: string, oid: string, upn: string, name: string }>}
 */
async function requireOperatorCaller(token) {
    const cfg = getConfig();
    const caller = await validateCallerToken(token, cfg.tokenAudiences);
    if (!isOperatorCaller(caller, cfg)) {
        const err = new Error('Nur Betreiber-Konten dürfen Lizenzen verwalten.');
        err.status = 403;
        throw err;
    }
    return caller;
}

/**
 * Lesen des zentralen Katalogs: Betreiber oder freigeschaltete Schule.
 * @param {string} token
 * @returns {Promise<{ caller: { tid: string, oid: string, upn: string, name: string }, isOperator: boolean }>}
 */
async function requireCatalogReader(token) {
    const cfg = getConfig();
    const caller = await validateCallerToken(token, cfg.tokenAudiences);
    if (isOperatorCaller(caller, cfg)) {
        return { caller, isOperator: true };
    }
    const fields = await lookupLicenseFields(caller.tid);
    const result = evaluateLicense({
        tenantId: caller.tid,
        fields,
        allowedStatuses: cfg.allowedStatuses
    });
    if (!result.allowed) {
        const err = new Error(result.message || 'Dieser Mandant ist nicht freigeschaltet.');
        err.status = 403;
        throw err;
    }
    return { caller, isOperator: false };
}

module.exports = { requireOperatorCaller, requireCatalogReader, isOperatorCaller };
